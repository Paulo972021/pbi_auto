#!/usr/bin/env python3
"""
test_executor_real.py — Teste do executor real

Dois modos:
  1. --dry-run (padrão): valida toda a ponte sem abrir o Power BI
  2. --live: executa de verdade contra o Power BI

Uso:
    python test_executor_real.py              # dry-run
    python test_executor_real.py --live       # execução real
"""

import sys
import os
import logging

# Garante path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config_model import CATALOG, TEMPLATES
from validator import validate_template
from codegen import generate_template_code
from storage import prepare_output_folder
from executor_real import build_runtime_filter_plan_from_template

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("test_executor_real")


def separator(title: str):
    print(f"\n{'=' * 60}")
    print(f"  {title}")
    print(f"{'=' * 60}")


def test_dry_run():
    """
    Testa toda a cadeia SEM abrir o Power BI:
      - validação do template
      - geração de código
      - build_runtime_filter_plan
      - preparação de pasta
    """
    separator("DRY-RUN: Validação da ponte template → runtime")

    template = TEMPLATES[0]
    tid = template.get("template_id", "?")

    print(f"\n  [TEST_REAL] template_id={tid}")
    print(f"  [TEST_REAL] page={template.get('page', '?')}")
    print(f"  [TEST_REAL] filters={template.get('filters', {})}")

    # 1. Validação
    val = validate_template(template, CATALOG)
    print(f"\n  1. Validação: {'✅ OK' if val['valid'] else '❌ FALHOU'}")
    if not val["valid"]:
        for err in val["errors"]:
            print(f"     {err}")
        return False

    # 2. Template code
    code = generate_template_code(template)
    print(f"  2. Template code: {code}")

    # 3. Output folder
    folder = prepare_output_folder(code)
    print(f"  3. Output folder: {folder} (exists={os.path.isdir(folder)})")

    # 4. Runtime filter plan
    plan = build_runtime_filter_plan_from_template(template)
    print(f"  4. Runtime FILTER_PLAN:")
    for fk, fv in plan.items():
        print(f"     {fk}: mode={fv['mode']} clear_first={fv['clear_first']} "
              f"target_values={fv['target_values']} required={fv['required']}")

    # 5. Verificar compatibilidade com o formato esperado pelo script
    for fk, fv in plan.items():
        assert isinstance(fv, dict), f"filtro {fk} não é dict"
        assert "mode" in fv, f"filtro {fk} sem 'mode'"
        assert "clear_first" in fv, f"filtro {fk} sem 'clear_first'"
        assert "target_values" in fv, f"filtro {fk} sem 'target_values'"
        assert isinstance(fv["target_values"], list), f"filtro {fk} target_values não é list"
        assert "required" in fv, f"filtro {fk} sem 'required'"
    print(f"  5. Formato FILTER_PLAN: ✅ compatível com o script")

    # 6. Testar todos os templates
    print(f"\n  6. Validando todos os {len(TEMPLATES)} templates:")
    all_ok = True
    for tpl in TEMPLATES:
        t_val = validate_template(tpl, CATALOG)
        t_code = generate_template_code(tpl)
        t_plan = build_runtime_filter_plan_from_template(tpl)
        icon = "✅" if t_val["valid"] else "❌"
        print(f"     {icon} {tpl['template_id']} → {t_code} "
              f"(filtros: {list(t_plan.keys())})")
        if not t_val["valid"]:
            all_ok = False

    separator("RESULTADO DRY-RUN")
    if all_ok:
        print("  ✅ Ponte validada — pronto para execução real")
        print(f"\n  Para executar de verdade:")
        print(f"    python test_executor_real.py --live")
    else:
        print("  ❌ Alguns templates falharam na validação")

    return all_ok


def test_live():
    """
    Executa um template de verdade contra o Power BI.
    Requer configpbi.py com url e browser configurados.
    """
    separator("LIVE: Execução real contra Power BI")

    template = TEMPLATES[0]
    tid = template.get("template_id", "?")

    print(f"\n  [TEST_REAL] template_id={tid}")
    print(f"  [TEST_REAL] page={template.get('page', '?')}")
    print(f"  [TEST_REAL] filters={template.get('filters', {})}")

    # Verificar se configpbi existe
    try:
        import configpbi
        url = getattr(configpbi, "url", "")
        browser = getattr(configpbi, "browser", "")
        if not url or not browser:
            print("\n  ❌ configpbi.py encontrado mas url ou browser está vazio")
            return False
        print(f"\n  configpbi.url = {url[:60]}...")
        print(f"  configpbi.browser = {browser}")
    except ImportError:
        print("\n  ❌ configpbi.py não encontrado.")
        print("     Crie com: url = '...' e browser = '...'")
        return False

    # Verificar se o módulo principal existe
    try:
        from executor_real import _import_pbi_module
        pbi = _import_pbi_module("pbi_auto_v06")
        if pbi is None:
            print("  ❌ pbi_auto_v06.py não encontrado no path")
            return False
        print(f"  pbi_auto_v06 importado com sucesso")
    except Exception as e:
        print(f"  ❌ Erro ao importar: {e}")
        return False

    # Executar
    print(f"\n  ▶ Executando template {tid} (LIVE)...\n")

    from executor_real import run_template_real_sync
    result = run_template_real_sync(template, CATALOG, "pbi_auto_v06")

    separator("RESULTADO LIVE")
    print(f"  [TEST_REAL] result =")
    for k, v in result.items():
        print(f"    {k}: {v}")

    if result.get("success"):
        print(f"\n  ✅ Template {tid} executado com sucesso!")
    else:
        print(f"\n  ❌ Template {tid} falhou: {result.get('error', '?')}")

    return result.get("success", False)


def main():
    live_mode = "--live" in sys.argv

    if live_mode:
        ok = test_live()
    else:
        ok = test_dry_run()

    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
