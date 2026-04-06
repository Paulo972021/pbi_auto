"""
executor_real.py — Etapa 9: Executor real (integração com Power BI)

Ponte entre o pipeline de templates e o script de automação existente.

Estratégia de integração:
  - NÃO duplica nem reescreve o script principal
  - Configura os parâmetros de runtime (FILTER_PLAN, TARGET_PAGE)
    diretamente no módulo importado antes de chamar run_export()
  - Restaura o estado original após a execução
  - Retorna resultado estruturado

Uso:
    from executor_real import run_template_real
    result = await run_template_real(template, catalog)
"""

import asyncio
import logging
import os
import sys
import importlib
import time
import shutil

# Pipeline imports
from validator import validate_template
from codegen import generate_template_code
from storage import prepare_output_folder

log = logging.getLogger("executor_real")

_DOWNLOAD_SETTLE_WAIT = 3.0


def _get_downloads_folder() -> str:
    """Retorna a pasta padrão de downloads do usuário atual."""
    return os.path.join(os.path.expanduser("~"), "Downloads")


def _snapshot_downloads(downloads_folder: str) -> set:
    """Cria snapshot dos arquivos atuais da pasta de downloads."""
    try:
        return set(os.listdir(downloads_folder))
    except Exception:
        return set()


def _detect_new_files(before_files: set, after_files: set) -> list:
    """Detecta novos arquivos entre dois snapshots."""
    return sorted(list(after_files - before_files))


def _wait_for_downloads_complete(downloads_folder: str, max_wait: int = 15) -> None:
    """
    Aguarda término de downloads temporários (crdownload/part/tmp).
    Mantém comportamento de exportação intacto; apenas sincroniza captura.
    """
    start = time.time()
    transient_suffixes = (".crdownload", ".part", ".tmp")
    while time.time() - start < max_wait:
        try:
            files = os.listdir(downloads_folder)
        except Exception:
            return
        if not any(name.lower().endswith(transient_suffixes) for name in files):
            return
        time.sleep(0.5)


def _move_files_to_output(new_files: list, downloads_folder: str, output_folder: str) -> list:
    """Move arquivos novos para a pasta de saída do template."""
    os.makedirs(output_folder, exist_ok=True)
    moved = []
    for name in new_files:
        src = os.path.join(downloads_folder, name)
        if not os.path.isfile(src):
            continue
        dst = os.path.join(output_folder, name)
        try:
            shutil.move(src, dst)
            moved.append(name)
        except Exception:
            continue
    return moved


# ---------------------------------------------------------------------------
# Função auxiliar: converte template em FILTER_PLAN do script
# ---------------------------------------------------------------------------

def build_runtime_filter_plan_from_template(template: dict) -> dict:
    """
    Converte os filtros de um template no formato FILTER_PLAN
    consumido pelo script de automação.

    Entrada (do template):
      {"entrante": "1", "epn_final": "ALMAVIVA"}

    Saída (FILTER_PLAN):
      {
          "entrante": {
              "mode": "single",
              "clear_first": True,
              "target_values": ["1"],
              "required": True,
          },
          "epn_final": {
              "mode": "single",
              "clear_first": True,
              "target_values": ["ALMAVIVA"],
              "required": True,
          },
      }
    """
    filters = template.get("filters", {})
    plan = {}

    for filter_key, filter_value in filters.items():
        normalized_key = filter_key.strip().lower()

        # Suporta valor único (string) ou lista
        if isinstance(filter_value, list):
            target_values = [str(v) for v in filter_value]
            mode = "multi" if len(target_values) > 1 else "single"
        else:
            target_values = [str(filter_value)]
            mode = "single"

        plan[normalized_key] = {
            "mode": mode,
            "clear_first": True,
            "target_values": target_values,
            "required": True,
        }

    return plan


# ---------------------------------------------------------------------------
# Importação controlada do script principal
# ---------------------------------------------------------------------------

def _import_pbi_module(module_name: str = "pbi_auto_v06"):
    """
    Importa o módulo do script principal de forma controlada.

    Tenta importar pelo nome. Se não estiver no path, adiciona
    o diretório pai ao sys.path.
    """
    try:
        return importlib.import_module(module_name)
    except ModuleNotFoundError:
        # Tenta adicionar diretório pai ao path
        parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if parent_dir not in sys.path:
            sys.path.insert(0, parent_dir)
        try:
            return importlib.import_module(module_name)
        except ModuleNotFoundError:
            log.error(
                f"❌ Módulo '{module_name}' não encontrado. "
                f"Verifique se o arquivo {module_name}.py está acessível."
            )
            return None


# ---------------------------------------------------------------------------
# Executor real
# ---------------------------------------------------------------------------

async def run_template_real(
    template: dict,
    catalog: dict,
    pbi_module_name: str = "pbi_auto_v06",
) -> dict:
    """
    Executa um template real contra o Power BI.

    Fluxo:
      A. Validar entrada
      B. Gerar metadados (template_code, output_folder)
      C. Preparar contexto runtime (FILTER_PLAN, TARGET_PAGE)
      D. Chamar fluxo real existente (run_export)
      E. Consolidar resultado

    Retorna:
      {
          "template_id": str,
          "template_code": str,
          "page": str,
          "success": bool,
          "filters_applied": list,
          "export_ok": bool,
          "output_folder": str,
          "error": str | None,
      }
    """
    tid = template.get("template_id", "unknown")
    page = template.get("page", "")
    filters = template.get("filters", {})

    result = {
        "template_id": tid,
        "template_code": "",
        "page": page,
        "success": False,
        "filters_applied": [],
        "export_ok": False,
        "output_folder": "",
        "error": None,
    }

    log.info(f"{'=' * 60}")
    log.info(f"[TEMPLATE_EXEC_START] template_id={tid}")
    log.info(f"  page={page}")
    log.info(f"  filters={filters}")
    log.info(f"{'=' * 60}")

    # ══════════════════════════════════════════════════════════════
    # ETAPA A — Validar entrada
    # ══════════════════════════════════════════════════════════════
    validation = validate_template(template, catalog)
    if not validation["valid"]:
        error_msg = "; ".join(validation["errors"])
        log.error(f"[TEMPLATE_VALIDATION_FAILED] {error_msg}")
        result["error"] = f"validation_failed: {error_msg}"
        return result

    log.info(f"[TEMPLATE_VALIDATION_OK] template_id={tid}")

    # ══════════════════════════════════════════════════════════════
    # ETAPA B — Gerar metadados
    # ══════════════════════════════════════════════════════════════
    template_code = generate_template_code(template)
    output_folder = prepare_output_folder(template_code)
    result["template_code"] = template_code
    result["output_folder"] = output_folder

    log.info(f"[TEMPLATE_METADATA]")
    log.info(f"  template_code={template_code}")
    log.info(f"  output_folder={output_folder}")

    # ══════════════════════════════════════════════════════════════
    # ETAPA C — Preparar contexto runtime
    # ══════════════════════════════════════════════════════════════
    runtime_filter_plan = build_runtime_filter_plan_from_template(template)

    log.info(f"[RUNTIME_FILTER_PLAN]")
    for fk, fv in runtime_filter_plan.items():
        log.info(f"  {fk}: mode={fv['mode']} target={fv['target_values']}")

    # Importar módulo principal
    pbi = _import_pbi_module(pbi_module_name)
    if pbi is None:
        result["error"] = f"module_not_found: {pbi_module_name}"
        return result

    # Salvar estado global original
    original_filter_plan = getattr(pbi, "FILTER_PLAN", {})
    original_target_page = getattr(pbi, "TARGET_PAGE", "")

    # Configurar runtime no módulo (monkey-patch controlado)
    pbi.FILTER_PLAN = runtime_filter_plan
    pbi.TARGET_PAGE = page

    log.info(f"[RUNTIME_CONFIG_APPLIED]")
    log.info(f"  TARGET_PAGE={page}")
    log.info(f"  FILTER_PLAN keys={list(runtime_filter_plan.keys())}")

    # ══════════════════════════════════════════════════════════════
    # ETAPA D — Chamar fluxo real
    # ══════════════════════════════════════════════════════════════
    try:
        url = getattr(pbi, "POWERBI_URL", "")
        browser_path = getattr(pbi, "BROWSER_PATH", "")

        if not url or not browser_path:
            result["error"] = "missing_config: POWERBI_URL ou BROWSER_PATH vazio"
            return result

        log.info(f"[TEMPLATE_EXEC_RUN] Chamando run_export...")
        log.info(f"  url={url[:80]}...")
        log.info(f"  browser={browser_path}")

        export_ok = await pbi.run_export(
            url=url,
            browser_path=browser_path,
            target_page=page,
            stop_after_filters=False,
        )

        result["export_ok"] = bool(export_ok)
        result["success"] = bool(export_ok)
        result["filters_applied"] = list(runtime_filter_plan.keys())

        if export_ok:
            log.info(f"[TEMPLATE_EXEC_OK] template_id={tid}")
        else:
            log.warning(f"[TEMPLATE_EXEC_FAIL] template_id={tid} export_ok=False")
            result["error"] = "export_returned_false"

    except Exception as e:
        log.error(f"[TEMPLATE_EXEC_ERROR] template_id={tid} error={e}")
        result["error"] = f"exception: {str(e)}"

    finally:
        # ══════════════════════════════════════════════════════════
        # Restaurar estado global original
        # ══════════════════════════════════════════════════════════
        pbi.FILTER_PLAN = original_filter_plan
        pbi.TARGET_PAGE = original_target_page
        log.info(f"[RUNTIME_CONFIG_RESTORED]")

    # ══════════════════════════════════════════════════════════════
    # ETAPA E — Consolidar resultado
    # ══════════════════════════════════════════════════════════════
    log.info(f"[TEMPLATE_EXEC_RESULT]")
    log.info(f"  template_id={result['template_id']}")
    log.info(f"  template_code={result['template_code']}")
    log.info(f"  success={result['success']}")
    log.info(f"  export_ok={result['export_ok']}")
    log.info(f"  filters_applied={result['filters_applied']}")
    log.info(f"  output_folder={result['output_folder']}")
    log.info(f"  error={result['error']}")

    return result


# ---------------------------------------------------------------------------
# Conveniência: execução síncrona (para chamar de scripts simples)
# ---------------------------------------------------------------------------

def run_template_real_sync(
    template: dict,
    catalog: dict,
    pbi_module_name: str = "pbi_auto_v06",
) -> dict:
    """Wrapper síncrono de run_template_real."""
    return asyncio.run(run_template_real(template, catalog, pbi_module_name))
