#!/usr/bin/env python3
"""
test_pipeline.py — Etapa 8: Script de teste completo (sem Power BI)

Executa todas as etapas do pipeline offline:
  1. Carrega CATALOG
  2. Carrega TEMPLATES
  3. Valida templates
  4. Gera código para cada template
  5. Prepara pasta de saída
  6. Roda executor_mock por template
  7. Roda batch_runner para lote completo

Critério: deve funcionar 100% sem acessar Power BI.
"""

import sys
import os

# Garante que o diretório do pipeline está no path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config_model import CATALOG, TEMPLATES
from validator import validate_template, validate_all_templates
from codegen import generate_template_code
from storage import prepare_output_folder, list_output_folders
from executor_mock import run_template_mock
from batch_runner import run_all_templates, group_templates_by_page


def separator(title: str):
    print(f"\n{'=' * 70}")
    print(f"  {title}")
    print(f"{'=' * 70}")


def main():
    exit_code = 0

    # ── 1. Catálogo ──
    separator("ETAPA 1 — CATÁLOGO")
    print(f"  Páginas disponíveis: {list(CATALOG.keys())}")
    for page_name, page_def in CATALOG.items():
        filters = page_def.get("filters", {})
        print(f"\n  📄 {page_name} ({page_def.get('label', '')})")
        for fk, fv in filters.items():
            print(f"     🎚️  {fk} ({fv.get('label', '')}): {fv.get('values', [])}")
    print(f"\n  [TEST] Catálogo carregado: {len(CATALOG)} página(s)")

    # ── 2. Templates ──
    separator("ETAPA 2 — TEMPLATES")
    print(f"  Templates definidos: {len(TEMPLATES)}")
    for tpl in TEMPLATES:
        print(f"    {tpl.get('template_id', '?')}: {tpl.get('page', '?')} → {tpl.get('filters', {})}")
    print(f"\n  [TEST] Templates carregados: {len(TEMPLATES)}")

    # ── 3. Validação individual ──
    separator("ETAPA 3 — VALIDAÇÃO INDIVIDUAL")
    for tpl in TEMPLATES:
        tid = tpl.get("template_id", "?")
        result = validate_template(tpl, CATALOG)
        if result["valid"]:
            print(f"  ✅ [{tid}] Template válido")
        else:
            print(f"  ❌ [{tid}] Template inválido:")
            for err in result["errors"]:
                print(f"     {err}")
            exit_code = 1

    # Validação em lote
    batch_val = validate_all_templates(TEMPLATES, CATALOG)
    print(f"\n  [TEST] Validação em lote: {batch_val['valid_count']}/{batch_val['total']} válidos")

    # ── 4. Geração de código ──
    separator("ETAPA 4 — GERAÇÃO DE CÓDIGO")
    codes = {}
    for tpl in TEMPLATES:
        tid = tpl.get("template_id", "?")
        code = generate_template_code(tpl)
        codes[tid] = code
        print(f"  [TEST] {tid} → Código gerado: {code}")

    # ── 5. Gestão de pastas ──
    separator("ETAPA 5 — PREPARAÇÃO DE PASTAS")
    folders = {}
    for tid, code in codes.items():
        folder = prepare_output_folder(code)
        folders[tid] = folder
        exists = os.path.isdir(folder)
        print(f"  [TEST] {tid} → Pasta criada: {folder} (exists={exists})")

    # Listar pastas criadas
    all_folders = list_output_folders()
    print(f"\n  [TEST] Pastas de saída existentes: {all_folders}")

    # ── 6. Executor mock individual ──
    separator("ETAPA 6 — EXECUTOR MOCK (INDIVIDUAL)")
    for tpl in TEMPLATES[:1]:  # Executa só o primeiro como teste individual
        tid = tpl.get("template_id", "?")
        code = codes.get(tid, "?")
        folder = folders.get(tid, "")
        print(f"\n  ▶ Executando mock para {tid}:")
        result = run_template_mock(tpl, code, folder)
        print(f"  [TEST] Resultado: success={result['success']} files={result['files']}")

    # ── 7. Agrupamento por página ──
    separator("ETAPA 7 — AGRUPAMENTO POR PÁGINA")
    groups = group_templates_by_page(TEMPLATES)
    for page, tpls in groups.items():
        tids = [t.get("template_id", "?") for t in tpls]
        print(f"  📄 {page}: {tids}")

    # ── 8. Batch runner (completo) ──
    separator("ETAPA 8 — BATCH RUNNER (LOTE COMPLETO)")

    # Adiciona um template inválido para testar rejeição
    templates_with_invalid = TEMPLATES + [
        {
            "template_id": "tpl_invalid",
            "page": "PAGINA_QUE_NAO_EXISTE",
            "filters": {"foo": "bar"},
        }
    ]

    batch_result = run_all_templates(templates_with_invalid, CATALOG)

    separator("RESULTADO FINAL")
    print(f"  Total:   {batch_result['total']}")
    print(f"  Sucesso: {batch_result['success']}")
    print(f"  Falha:   {batch_result['failed']}")
    print(f"  Pulados: {batch_result['skipped']}")
    print(f"  Por página: {batch_result['by_page']}")

    print()
    for r in batch_result["results"]:
        icon = {"success": "✅", "failed": "❌", "skipped": "⏭️"}.get(r["status"], "?")
        line = f"  {icon} {r['template_id']}: {r['status']}"
        if r.get("template_code"):
            line += f" → {r['template_code']}"
        if r.get("files"):
            line += f" ({len(r['files'])} arquivos)"
        if r.get("errors"):
            line += f" errors={r['errors']}"
        print(line)

    # Verificação de saída
    final_folders = list_output_folders()
    print(f"\n  Pastas de saída finais: {final_folders}")

    if batch_result["success"] == len(TEMPLATES) and batch_result["skipped"] == 1:
        print("\n  ✅ PIPELINE OFFLINE COMPLETO — TUDO OK")
    else:
        print("\n  ⚠️  Resultado inesperado — verificar")
        exit_code = 1

    return exit_code


if __name__ == "__main__":
    sys.exit(main())
