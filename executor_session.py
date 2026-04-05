"""
executor_session.py — Executor com sessão compartilhada por página

Arquitetura:
  1 página = 1 browser/session
  N templates da mesma página = mesma sessão

Camada A (sessão): browser + aba + navegação → uma vez por página
Camada B (execução): filtros + exportação + downloads → por template

Redução de foco:
  - safe_focus_tab chamado apenas 1x no início da sessão
  - todas as operações subsequentes via DOM/CDP/evaluate
  - nenhum tab.activate() entre templates

NÃO altera: lógica de exportação, lógica de filtros, pbi_auto_v06.py
"""

import asyncio
import logging
import os
import shutil
import sys
import time
import importlib

from validator import validate_template
from codegen import generate_template_code
from storage import prepare_output_folder
from executor_real import (
    build_runtime_filter_plan_from_template,
    _get_downloads_folder,
    _snapshot_downloads,
    _detect_new_files,
    _wait_for_downloads_complete,
    _move_files_to_output,
    _DOWNLOAD_SETTLE_WAIT,
)

log = logging.getLogger("executor_session")


# ---------------------------------------------------------------------------
# Importação do módulo PBI
# ---------------------------------------------------------------------------

def _import_pbi(module_name: str = "pbi_auto_v06"):
    """Importa o módulo PBI com fallback de path."""
    try:
        return importlib.import_module(module_name)
    except ModuleNotFoundError:
        parent = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if parent not in sys.path:
            sys.path.insert(0, parent)
        try:
            return importlib.import_module(module_name)
        except ModuleNotFoundError:
            return None


# ---------------------------------------------------------------------------
# Transição entre templates (cleanup)
# ---------------------------------------------------------------------------

async def _template_transition_cleanup(
    tab, pbi, page: str,
    from_tid: str, to_tid: str,
) -> dict:
    """
    Limpa estado entre templates da mesma página.

    Faz:
      1. ESC para fechar menus/overlays
      2. cleanup_residual_ui
      3. dismiss popups
      4. espera render estabilizar
      5. verifica visuais ainda presentes

    Retorna: {"ui_clean": bool, "filters_reset_ok": bool, "report_stable": bool}
    """
    log.info(f"[TEMPLATE_TRANSITION_CLEANUP]")
    log.info(f"  page={page}")
    log.info(f"  from_template={from_tid}")
    log.info(f"  to_template={to_tid}")

    # 1. ESC
    await pbi.press_escape(tab, times=3, wait_each=0.3)
    await asyncio.sleep(0.5)

    # 2. Cleanup UI
    await pbi.cleanup_residual_ui(tab, stage_label=f"transition {from_tid}→{to_tid}", aggressive=True)
    await asyncio.sleep(0.5)

    # 3. Dismiss popups
    await pbi.dismiss_sensitive_data_popup(tab)

    # 4. Esperar render
    await asyncio.sleep(2)

    # 5. Verificar visuais
    try:
        visual_count = await tab.evaluate(
            "document.querySelectorAll('visual-container').length"
        )
        visual_count = int(visual_count or 0)
    except Exception:
        visual_count = 0

    # 6. Verificar overlays
    try:
        overlay_count = await tab.evaluate("""
            (() => {
                const visible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };
                return Array.from(document.querySelectorAll(
                    '[role="menu"], [role="dialog"], [role="listbox"], ' +
                    '[aria-modal="true"], .cdk-overlay-pane, .contextMenu, .popup, .modal'
                )).filter(visible).length;
            })()
        """)
        overlay_count = int(overlay_count or 0)
    except Exception:
        overlay_count = -1

    ui_clean = overlay_count == 0
    report_stable = visual_count > 0

    result = {
        "ui_clean": ui_clean,
        "filters_reset_ok": True,  # filtros serão limpos pelo clear_first do próximo template
        "report_stable": report_stable,
    }

    log.info(f"  ui_clean={ui_clean}")
    log.info(f"  visible_overlays={overlay_count}")
    log.info(f"  visual_count={visual_count}")
    log.info(f"  report_stable={report_stable}")

    return result


# ---------------------------------------------------------------------------
# Execução de UM template dentro de sessão existente
# ---------------------------------------------------------------------------

async def _run_single_template_in_session(
    tab, pbi, template: dict, catalog: dict,
) -> dict:
    """
    Executa um template usando tab/browser já abertos.

    NÃO abre browser. NÃO navega para página (já está nela).
    Apenas: configura filtros → scan_slicers → exporta → captura downloads.
    """
    tid = template.get("template_id", "unknown")
    page = template.get("page", "")
    template_code = generate_template_code(template)
    output_folder = prepare_output_folder(template_code)
    runtime_plan = build_runtime_filter_plan_from_template(template)

    result = {
        "template_id": tid,
        "template_code": template_code,
        "page": page,
        "success": False,
        "export_ok": False,
        "filters_applied": list(runtime_plan.keys()),
        "output_folder": output_folder,
        "files_moved": 0,
        "files_list": [],
        "error": None,
    }

    log.info(f"[PAGE_SESSION_TEMPLATE_START]")
    log.info(f"  page={page}")
    log.info(f"  template_id={tid}")
    log.info(f"  template_code={template_code}")
    log.info(f"  output_folder={output_folder}")
    log.info(f"  filter_plan={list(runtime_plan.keys())}")

    # ── Configurar FILTER_PLAN no módulo ──
    original_plan = getattr(pbi, "FILTER_PLAN", {})
    pbi.FILTER_PLAN = runtime_plan

    downloads_folder = _get_downloads_folder()
    before_files = _snapshot_downloads(downloads_folder)

    try:
        # ── Bloquear links externos ──
        await pbi.block_microsoft_learn_and_external_links(tab)
        await pbi.cleanup_residual_ui(tab, stage_label=f"pre-template {tid}", aggressive=True)

        # ── Verificar visuais presentes ──
        loaded = await pbi.wait_for_visual_containers(tab, retries=6, wait_seconds=3)
        if not loaded:
            result["error"] = "no_visual_containers"
            return result

        # ── FASE 1: Scan slicers (inclui enumeração + aplicação de filtros) ──
        log.info(f"  📌 Scan de slicers + aplicação de filtros para {tid}...")
        slicers = await pbi.scan_slicers(tab)
        pbi._display_slicers_inline(slicers)

        # ── POST-FILTER GATE ──
        filters_expected = list(runtime_plan.keys())
        filters_ok = []
        filters_failed = []
        critical_incidents = []

        for s in slicers:
            norm_title = pbi.normalize_slicer_name(s.get("title", ""))
            if norm_title not in runtime_plan:
                continue
            plan_cfg = runtime_plan[norm_title]
            target_set = {v.lower() for v in plan_cfg.get("target_values", [])}
            selected_set = {v.lower() for v in (s.get("selectedValues") or [])}
            if target_set.issubset(selected_set):
                filters_ok.append(norm_title)
            else:
                filters_failed.append(norm_title)
                if plan_cfg.get("required", True):
                    critical_incidents.append(
                        f"required_filter_failed:{norm_title}"
                    )

        # Cleanup pós-filtros
        await pbi.press_escape(tab, times=3, wait_each=0.3)
        await asyncio.sleep(0.5)
        await pbi.cleanup_residual_ui(tab, stage_label="post_filter", aggressive=True)
        await asyncio.sleep(3)  # render wait

        if critical_incidents:
            log.warning(f"  [POST_FILTER_GATE] critical_incidents={critical_incidents}")
            result["error"] = f"filter_gate_failed: {critical_incidents}"
            return result

        log.info(f"  [POST_FILTER_GATE] filters_ok={filters_ok} export_release=True")

        # ── FASE 2: Scan e exportação de visuais ──
        log.info(f"  📌 Escaneando visuais para exportação...")
        slicer_indexes = {s.get("index") for s in slicers}
        visuals = await pbi.scan_visuals(tab, slicer_indexes=slicer_indexes)

        exportable = [i for i, v in enumerate(visuals) if v.get("hasExportData")]
        if not exportable:
            log.warning(f"  ⚠️ Nenhum visual exportável encontrado")
            result["error"] = "no_exportable_visuals"
            return result

        log.info(f"  📥 Exportando {len(exportable)} visuais...")
        export_results = await pbi.export_selected_visuals(tab, visuals, exportable)
        pbi.display_export_summary(export_results)

        export_ok = any(r.get("success") for r in export_results)
        result["export_ok"] = export_ok

    except Exception as e:
        log.error(f"  [TEMPLATE_EXEC_ERROR] {tid}: {e}")
        result["error"] = f"exception: {str(e)}"
    finally:
        pbi.FILTER_PLAN = original_plan

    # ── Capturar downloads ──
    if result["export_ok"]:
        log.info(f"  [DOWNLOAD_SETTLE_WAIT] {_DOWNLOAD_SETTLE_WAIT}s...")
        await asyncio.sleep(_DOWNLOAD_SETTLE_WAIT)
        _wait_for_downloads_complete(downloads_folder, max_wait=15)

        after_files = _snapshot_downloads(downloads_folder)
        new_files = _detect_new_files(before_files, after_files)

        if new_files:
            moved = _move_files_to_output(new_files, downloads_folder, output_folder)
            result["files_moved"] = len(moved)
            result["files_list"] = moved
            log.info(f"  [DOWNLOAD_CAPTURE] files_moved={len(moved)} files={moved}")
        else:
            log.warning(f"  [DOWNLOAD_CAPTURE_WARNING] no_new_files_detected")

    result["success"] = bool(result["export_ok"] and result.get("files_moved", 0) > 0)

    log.info(f"[PAGE_SESSION_TEMPLATE_END]")
    log.info(f"  page={page}")
    log.info(f"  template_id={tid}")
    log.info(f"  success={result['success']}")
    log.info(f"  export_ok={result['export_ok']}")
    log.info(f"  files_moved={result['files_moved']}")

    return result


# ---------------------------------------------------------------------------
# Sessão compartilhada por página
# ---------------------------------------------------------------------------

async def run_templates_for_page_in_shared_session(
    page: str,
    templates: list,
    catalog: dict,
    pbi_module_name: str = "pbi_auto_v06",
) -> dict:
    """
    Executa N templates da mesma página em uma única sessão de browser.

    Fluxo:
      1. Inicia browser (uma vez)
      2. Abre relatório (uma vez)
      3. Navega para a página (uma vez)
      4. Para cada template:
         a. Configura filtros
         b. Exporta visuais
         c. Captura downloads
         d. Cleanup transição
      5. Fecha browser

    FOCUS_POLICY:
      - safe_focus_tab chamado apenas 1x no início
      - todas as operações por DOM/CDP/evaluate
    """
    import tempfile

    pbi = _import_pbi(pbi_module_name)
    if pbi is None:
        return {"page": page, "error": f"module_not_found: {pbi_module_name}",
                "results": [], "success": 0, "failed": len(templates), "skipped": 0}

    url = getattr(pbi, "POWERBI_URL", "")
    browser_path = getattr(pbi, "BROWSER_PATH", "")

    if not url or not browser_path:
        return {"page": page, "error": "missing_config",
                "results": [], "success": 0, "failed": len(templates), "skipped": 0}

    results = []
    success_count = 0
    failed_count = 0
    skipped_count = 0
    total_files = 0

    log.info(f"[PAGE_SESSION_START]")
    log.info(f"  page={page}")
    log.info(f"  templates_count={len(templates)}")
    log.info(f"  browser_reused=False")  # nova sessão para esta página

    # ── FOCUS_POLICY: início da sessão ──
    log.info(f"[FOCUS_POLICY] stage=session_start safe_focus_called=True reason=initial_session_setup fallback_only=False")

    browser = None
    tab = None
    owned_tab_refs = set()

    try:
        # ══════════════════════════════════════════════════════════
        # CAMADA A — Sessão (uma vez por página)
        # ══════════════════════════════════════════════════════════

        browser_path_norm = pbi.normalize_browser_path(browser_path)
        browser, _profile_dir = await pbi.start_isolated_browser(browser_path_norm)

        tab = await pbi.open_report_tab(browser, url, owned_tab_refs)

        # Focus ÚNICO no início da sessão
        await pbi.safe_focus_tab(tab)
        await pbi.close_extra_tabs_created_by_script(browser, tab, owned_tab_refs)
        await pbi.allow_multiple_downloads(tab)

        log.info(f"  ⏳ Aguardando carregamento inicial ({pbi.PAGE_LOAD_WAIT}s)...")
        await asyncio.sleep(pbi.PAGE_LOAD_WAIT)

        await pbi.block_microsoft_learn_and_external_links(tab)
        await pbi.cleanup_residual_ui(tab, stage_label="session_init", aggressive=True)

        # Verificar URL
        try:
            current_url = await pbi.get_tab_url(tab)
            if current_url and "powerbi.com" not in current_url.lower():
                await tab.get(url)
                await asyncio.sleep(pbi.LONG_WAIT)
        except Exception:
            pass

        # Aguardar visuais
        loaded = await pbi.wait_for_visuals_or_abort(
            tab, stage_label="session_init", retries=10, wait_seconds=3
        )
        if not loaded:
            return {"page": page, "error": "report_not_loaded",
                    "results": [], "success": 0, "failed": len(templates), "skipped": 0}

        # Navegar para a página (uma vez)
        original_target_page = getattr(pbi, "TARGET_PAGE", "")
        pbi.TARGET_PAGE = page

        if page:
            page_ok = await pbi.navigate_to_report_page(tab, page)
            if not page_ok:
                return {"page": page, "error": f"navigation_failed:{page}",
                        "results": [], "success": 0, "failed": len(templates), "skipped": 0}

            loaded2 = await pbi.wait_for_visuals_or_abort(
                tab, stage_label=f"navigate_{page}", retries=8, wait_seconds=3
            )
            if not loaded2:
                return {"page": page, "error": "visuals_not_loaded_after_navigation",
                        "results": [], "success": 0, "failed": len(templates), "skipped": 0}

        await pbi.close_open_menus_and_overlays(tab, aggressive=True)

        # ══════════════════════════════════════════════════════════
        # CAMADA B — Execução por template (N vezes)
        # ══════════════════════════════════════════════════════════

        previous_tid = None

        for i, tpl in enumerate(templates):
            tid = tpl.get("template_id", "?")

            # Validar
            validation = validate_template(tpl, catalog)
            if not validation["valid"]:
                error_msg = "; ".join(validation["errors"])
                log.info(f"[BATCH_LIVE_TEMPLATE_SKIP] {tid}: {error_msg}")
                results.append({
                    "template_id": tid,
                    "success": False,
                    "skipped": True,
                    "error": f"validation_failed: {error_msg}",
                })
                skipped_count += 1
                continue

            # Transição entre templates
            if previous_tid is not None:
                cleanup = await _template_transition_cleanup(
                    tab, pbi, page, previous_tid, tid
                )
                if not cleanup.get("report_stable"):
                    log.warning(f"  ⚠️ Relatório instável após transição, aguardando mais...")
                    await asyncio.sleep(3)

            # FOCUS_POLICY: NÃO chamar safe_focus entre templates
            log.info(f"[FOCUS_POLICY] stage=template_{tid} safe_focus_called=False reason=reusing_session fallback_only=False")

            # Executar template
            tpl_result = await _run_single_template_in_session(tab, pbi, tpl, catalog)
            results.append(tpl_result)

            if tpl_result.get("success"):
                success_count += 1
                total_files += tpl_result.get("files_moved", 0)
            else:
                failed_count += 1

            previous_tid = tid

    except Exception as e:
        log.error(f"[PAGE_SESSION_ERROR] page={page} error={e}")
        # Templates não executados contam como failed
        remaining = len(templates) - len(results)
        failed_count += remaining

    finally:
        # Restaurar TARGET_PAGE
        if pbi:
            pbi.TARGET_PAGE = original_target_page if 'original_target_page' in dir() else ""

        # Fechar browser
        if tab or browser:
            try:
                await pbi.graceful_browser_shutdown(tab, browser)
            except Exception:
                pass

    log.info(f"[PAGE_SESSION_END]")
    log.info(f"  page={page}")
    log.info(f"  templates_total={len(templates)}")
    log.info(f"  success={success_count}")
    log.info(f"  failed={failed_count}")
    log.info(f"  skipped={skipped_count}")
    log.info(f"  total_files={total_files}")

    return {
        "page": page,
        "results": results,
        "success": success_count,
        "failed": failed_count,
        "skipped": skipped_count,
        "total_files": total_files,
        "error": None,
    }
