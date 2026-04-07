# 01) `scan_slicers`

## Finalidade
Escanear slicers do relatório, coletar metadados e acionar a aplicação de filtros configurados no `FILTER_PLAN`.

## Quem chama
- `run_export`.
- `executor_session.py` também chama `pbi.scan_slicers(...)` no fluxo de sessão compartilhada.

## Quais funções ela chama (núcleo)
- `get_report_dom_context_info`
- `cleanup_residual_ui`
- `_simple_keyboard_probe`
- `press_escape`
- `_enumerate_slicer_via_micro_scroll` (para enrich de enumeração)
- `apply_filter_plan` (quando há plano de filtro para o slicer)

## Entradas e saídas
- **Entrada:** `tab`.
- **Saída:** `list[dict]` com descrição dos slicers e estado de seleção.

## Ações principais (fluxo)
1. Limpa estado visual residual (menus/overlays).
2. Executa scan de slicers no DOM e normaliza títulos/campos.
3. Para cada slicer com filtro no plano, tenta enumeração auxiliar.
4. Chama `apply_filter_plan` e atualiza estado do slicer no resultado final.
5. Faz fechamento/normalização de UI com `ESC` no final.

## Pontos de atenção
- Função longa e central; mistura descoberta de slicers com aplicação de filtro.
- Alto impacto de regressão se alterar ordem de limpeza/scan/apply.

