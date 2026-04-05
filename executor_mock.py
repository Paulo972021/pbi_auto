"""
executor_mock.py — Etapa 5: Executor simulado (sem Power BI)

Simula o fluxo completo de um template:
  1. Abrir tela
  2. Aplicar filtros
  3. Exportar visuais
  4. Salvar em pasta

Usado para validação offline do pipeline antes da integração real.
"""

import os
import time


def run_template_mock(template: dict, template_code: str, output_folder: str) -> dict:
    """
    Executa um template em modo mock.

    Retorna:
      {"success": True/False, "template_id": str, "template_code": str,
       "steps": [...], "output_folder": str, "duration_ms": int}
    """
    tid = template.get("template_id", "?")
    page = template.get("page", "?")
    filters = template.get("filters", {})
    steps = []
    start = time.time()

    # 1. Abrir tela
    steps.append(f"[MOCK] Abrindo tela '{page}'...")
    print(f"  [MOCK] Abrindo tela '{page}'...")

    # 2. Aplicar filtros
    for filter_key, filter_value in filters.items():
        msg = f"[MOCK] Aplicando filtro '{filter_key}' = '{filter_value}'"
        steps.append(msg)
        print(f"  {msg}")

    # 3. Validar filtros
    steps.append("[MOCK] Validando filtros aplicados... OK")
    print("  [MOCK] Validando filtros aplicados... OK")

    # 4. Limpar UI
    steps.append("[MOCK] Limpando UI residual...")
    print("  [MOCK] Limpando UI residual...")

    # 5. Aguardar estabilização
    steps.append("[MOCK] Aguardando estabilização do relatório...")
    print("  [MOCK] Aguardando estabilização do relatório...")

    # 6. Exportar visuais
    mock_visuals = ["Tabela_Resumo", "Grafico_Barras", "KPI_Total"]
    for visual in mock_visuals:
        msg = f"[MOCK] Exportando visual '{visual}'..."
        steps.append(msg)
        print(f"  {msg}")

        # Cria arquivo mock na pasta de saída
        mock_file = os.path.join(output_folder, f"{visual}.xlsx")
        with open(mock_file, "w") as f:
            f.write(f"mock export: {visual} | template: {tid} | code: {template_code}\n")

    # 7. Resultado
    duration_ms = int((time.time() - start) * 1000)
    steps.append(f"[MOCK] Exportação concluída em {duration_ms}ms")
    print(f"  [MOCK] Exportação concluída em {duration_ms}ms")

    files_created = os.listdir(output_folder)
    steps.append(f"[MOCK] Arquivos criados: {files_created}")
    print(f"  [MOCK] Arquivos criados: {files_created}")

    return {
        "success": True,
        "template_id": tid,
        "template_code": template_code,
        "steps": steps,
        "output_folder": output_folder,
        "files": files_created,
        "duration_ms": duration_ms,
    }
