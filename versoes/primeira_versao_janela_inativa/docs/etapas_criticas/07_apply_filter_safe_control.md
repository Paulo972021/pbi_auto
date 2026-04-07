# 07) `apply_filter_safe_control`

## Finalidade
Versão alternativa "safe control" para aplicar filtro com `_read_filter_state`, `_safe_clear_filter` e `_apply_single_filter_value`.

## Quem chama
- **Nenhum chamador ativo neste snapshot**.

## Quais funções ela chama
- `_click_slicer_header_safe`, `_cdp_click`
- `_read_filter_state`
- `_safe_clear_filter`
- `_apply_single_filter_value`
- `press_escape`

## Entradas e saídas
- **Entrada:** `tab`, `slicer`, `plan_cfg`.
- **Saída:** `dict` estruturado com status/diagnóstico.

## Ações principais
1. Ativa slicer e lê estado.
2. Opcionalmente limpa seleção.
3. Aplica valores individualmente com revalidação.
4. Retorna estado final com sucesso/falha.

## Observação crítica
Você está correto em estranhar: por análise de chamadas deste snapshot, essa função está definida, mas não entra no caminho principal atual (`run_export -> scan_slicers -> apply_filter_plan`).

## Risco
- Pode ficar desatualizada silenciosamente por não ser exercitada no fluxo padrão.

