# 03) `clear_slicer_selection`

## Finalidade
Limpar a seleção atual de um slicer antes de aplicar novos valores.

## Quem chama
- `apply_filter_plan`.

## Quais funções ela chama
- `_cdp_click` (clique robusto em botão/área de clear)
- `read_current_selection` (rechecagens antes/depois)

## Entradas e saídas
- **Entrada:** `tab`, `idx`, `slicer_title`, `state`.
- **Saída:** `bool` (`True` se limpeza confirmada).

## Ações principais
1. Tenta rota preferencial de clear via elemento de limpeza.
2. Se necessário, usa fallback por clique em item selecionado.
3. Revalida estado até confirmar ausência de seleção.

## Pontos de atenção
- Sensível a variações de UI do Power BI.
- Pode aparentar sucesso visual sem refletir no estado lógico se não houver revalidação.

