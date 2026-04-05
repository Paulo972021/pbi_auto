"""
batch_runner.py — Etapa 6+7: Execução em lote com agrupamento por tela

Regras:
  - executa todos os templates válidos
  - ordem sequencial, agrupados por página
  - template inválido → abortar apenas ele
  - continuar execução dos demais
"""

from collections import OrderedDict

from validator import validate_template
from codegen import generate_template_code
from storage import prepare_output_folder
from executor_mock import run_template_mock


def group_templates_by_page(templates: list) -> dict:
    """
    Agrupa templates por página, mantendo ordem de inserção.

    Retorna: {"COMPARATIVO": [tpl1, tpl2], "OUTRA_TELA": [tpl3]}
    """
    groups = OrderedDict()
    for tpl in templates:
        page = tpl.get("page", "UNKNOWN")
        if page not in groups:
            groups[page] = []
        groups[page].append(tpl)
    return groups


def run_all_templates(templates: list, catalog: dict) -> dict:
    """
    Executa todos os templates em lote.

    Fluxo:
      1. Agrupa por página
      2. Para cada página:
         a. Para cada template:
            - valida
            - gera código
            - prepara pasta
            - executa (mock)
      3. Template inválido → pula, não aborta o lote

    Retorna:
      {"results": [...], "total": int, "success": int,
       "failed": int, "skipped": int, "by_page": {...}}
    """
    results = []
    groups = group_templates_by_page(templates)

    print("\n" + "=" * 70)
    print("📦 EXECUÇÃO EM LOTE")
    print("=" * 70)
    print(f"  Templates: {len(templates)}")
    print(f"  Páginas: {list(groups.keys())}")

    for page, tpls in groups.items():
        print(f"\n  ─── Página: {page} ({len(tpls)} templates) ───")

    print()

    success_count = 0
    failed_count = 0
    skipped_count = 0

    for page, page_templates in groups.items():
        print(f"\n{'─' * 60}")
        print(f"📄 PÁGINA: {page}")
        print(f"{'─' * 60}")

        for i, tpl in enumerate(page_templates, 1):
            tid = tpl.get("template_id", "?")
            print(f"\n  ▶ [{i}/{len(page_templates)}] Template: {tid}")

            # Validar
            validation = validate_template(tpl, catalog)
            if not validation["valid"]:
                print(f"  ❌ Template inválido:")
                for err in validation["errors"]:
                    print(f"     {err}")
                results.append({
                    "template_id": tid,
                    "status": "skipped",
                    "reason": "validation_failed",
                    "errors": validation["errors"],
                })
                skipped_count += 1
                continue

            print(f"  ✅ Validação OK")

            # Gerar código
            template_code = generate_template_code(tpl)
            print(f"  📝 Código: {template_code}")

            # Preparar pasta
            output_folder = prepare_output_folder(template_code)
            print(f"  📂 Pasta: {output_folder}")

            # Executar (mock)
            try:
                exec_result = run_template_mock(tpl, template_code, output_folder)
                if exec_result.get("success"):
                    success_count += 1
                    results.append({
                        "template_id": tid,
                        "status": "success",
                        "template_code": template_code,
                        "output_folder": output_folder,
                        "files": exec_result.get("files", []),
                        "duration_ms": exec_result.get("duration_ms", 0),
                    })
                else:
                    failed_count += 1
                    results.append({
                        "template_id": tid,
                        "status": "failed",
                        "template_code": template_code,
                        "reason": "execution_failed",
                    })
            except Exception as e:
                failed_count += 1
                results.append({
                    "template_id": tid,
                    "status": "failed",
                    "template_code": template_code,
                    "reason": f"exception: {str(e)}",
                })
                print(f"  ❌ Erro na execução: {e}")

    # Resumo
    print(f"\n{'=' * 70}")
    print("📊 RESUMO DO LOTE")
    print(f"{'=' * 70}")
    print(f"  Total:    {len(templates)}")
    print(f"  Sucesso:  {success_count}")
    print(f"  Falha:    {failed_count}")
    print(f"  Pulados:  {skipped_count}")

    return {
        "results": results,
        "total": len(templates),
        "success": success_count,
        "failed": failed_count,
        "skipped": skipped_count,
        "by_page": {page: len(tpls) for page, tpls in groups.items()},
    }
