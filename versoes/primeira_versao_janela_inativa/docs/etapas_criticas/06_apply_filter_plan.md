# 06) `apply_filter_plan`

## Finalidade
Aplicar o plano declarativo de filtro (`FILTER_PLAN`) a um slicer específico com validação final.

## Quem chama
- `scan_slicers`.

## Quais funções ela chama
- `_click_slicer_header_safe`, `_cdp_click`
- `read_current_selection`
- `clear_slicer_selection`
- `_scroll_to_find_value_in_slicer`
- `validate_filter_final`

## Entradas e saídas
- **Entrada:** `tab`, `slicer` (dict), `plan_cfg` (dict), `enum_hint` opcional.
- **Saída:** `dict` com status, ações executadas, valores aplicados/falhos e validação.

## Ações principais
1. Garante ativação do slicer.
2. Lê estado atual e avalia necessidade de clear (`clear_first`).
3. Para cada target, localiza via scroll estratificado e tenta clique.
4. Revalida seleção após cada ação.
5. Executa validação final (`validate_filter_final`) e retorna relatório detalhado.

## Pontos de atenção
- É o coração da aplicação de filtros no fluxo ativo.
- Depende de múltiplos helpers críticos; qualquer mudança de contrato em helper quebra o fluxo.

