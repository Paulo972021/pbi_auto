#!/usr/bin/env python3
"""
test_batch_real.py — Teste do batch runner com sessão compartilhada

Modos:
  --dry-run (padrão): valida cadeia + mostra agrupamento
  --live: executa todos os templates (1 sessão por página)

Uso:
    python test_batch_real.py
    python test_batch_real.py --live
"""

import sys
import os
import logging

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config_model import CATALOG, TEMPLATES
from validator import validate_template
from codegen import generate_template_code
from storage import prepare_output_folder
from batch_runner_real import run_all_templates_live, group_templates_by_page

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)


def separator(title: str):
    print(f"\n{'=' * 60}")
    print(f"  {title}")
    print(f"{'=' * 60}")


def test_dry_run():
    """Valida cadeia + mostra agrupamento por página."""
    separator("DRY-RUN: Sessão compartilhada por página")

    groups = group_templates_by_page(TEMPLATES)
    print(f"\n  Templates: {len(TEMPLATES)}")
    print(f"  Páginas: {list(groups.keys())}")
    print(f"  Sessões necessárias: {len(groups)} (1 por página)\n")

    for page, tpls in groups.items():
        print(f"  📄 {page} — {len(tpls)} templates, 1 sessão:")
        for tpl in tpls:
            tid = tpl.get("template_id", "?")
            val = validate_template(tpl, CATALOG)
            code = generate_template_code(tpl)
            icon = "✅" if val["valid"] else "❌"
            print(f"    {icon} {tid} → {code}")
            if not val["valid"]:
                for err in val["errors"]:
                    print(f"       ⚠️  {err}")

    separator("RESULTADO DRY-RUN")
    print(f"  ✅ {len(TEMPLATES)} templates validados")
    print(f"  📦 {len(groups)} sessão(ões) de browser necessária(s)")
    print(f"\n  Para executar de verdade:")
    print(f"    python test_batch_real.py --live")
    return True


def test_live():
    """Executa todos os templates com sessão compartilhada."""
    separator("LIVE: Sessão compartilhada por página")

    try:
        import configpbi
        url = getattr(configpbi, "url", "")
        browser = getattr(configpbi, "browser", "")
        if not url or not browser:
            print("  ❌ configpbi.url ou browser vazio")
            return False
        print(f"  configpbi OK")
    except ImportError:
        print("  ❌ configpbi.py não encontrado")
        return False

    groups = group_templates_by_page(TEMPLATES)
    print(f"  Templates: {len(TEMPLATES)}")
    print(f"  Sessões: {len(groups)} (1 por página)")
    print()

    result = run_all_templates_live(TEMPLATES, CATALOG, "pbi_auto_v06")

    separator("RESULTADO FINAL")
    print(f"  Total:    {result['total']}")
    print(f"  Sucesso:  {result['success']}")
    print(f"  Falha:    {result['failed']}")
    print(f"  Sessões:  {result.get('total_sessions', '?')}")
    print(f"  Arquivos: {result.get('total_files_moved', 0)}")
    print(f"  Duração:  {result.get('duration_s', '?')}s")

    return result["failed"] == 0 and result["skipped"] == 0


def main():
    ok = test_live() if "--live" in sys.argv else test_dry_run()
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
