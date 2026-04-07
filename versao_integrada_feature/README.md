# Versão integrada da feature (conservadora)

Arquivo principal: `versao_antiga_com_feature.py`

## O que entrega
- Nova função para executar múltiplos templates na mesma página/sessão:
  - `run_templates_for_page_shared_session`
  - `run_templates_for_page_shared_session_sync`
- Simulação de foco com fallback usando funções do módulo PBI (`ensure_focus_visibility_emulation` + `safe_focus_tab`).
- Transferência automática de arquivos exportados para pastas geradas por template (`page + filtros`) via:
  - `generate_template_code`
  - `prepare_output_folder`

## Como usar
```python
from versao_integrada_feature.versao_antiga_com_feature import run_templates_for_page_shared_session_sync

result = run_templates_for_page_shared_session_sync(
    page="COMPARATIVO",
    templates=[...],
    catalog={...},
    pbi_module_name="pbi_nico"  # ou "versao_antiga" quando existir
)
```
