# 📦 Versão integrada (dentro de `versoes/`) — múltiplos templates por sessão

Esta pasta contém a versão integrada da feature pedida, em formato conservador:

- execução de múltiplos templates na mesma seção/página sem reabrir sessão toda hora;
- simulação de foco com fallback (`ensure_focus_visibility_emulation` -> `safe_focus_tab`);
- transferência de arquivos exportados para pastas automáticas por template (página + filtros).

## Arquivo principal
- `versao_antiga_com_feature.py`

## Documentação
- `docs/dependencias.md`
- `docs/fluxo_geral.md`
- `docs/microfluxos.md`
- `docs/legenda_navegacao.md`
- `docs/tecnologias_e_bibliotecas.md`
- `docs/etapas_criticas/` (parte crítica)

## Uso rápido
```python
from versoes.feature_multi_template_integrada.versao_antiga_com_feature import run_templates_for_page_shared_session_sync

result = run_templates_for_page_shared_session_sync(
    page="COMPARATIVO",
    templates=[...],
    catalog={...},
    pbi_module_name="pbi_nico"  # ou "versao_antiga"
)
```
