# 05) `_scroll_to_find_value_in_slicer`

## Finalidade
Localizar um valor alvo em slicer usando estratégia em camadas (viewport atual, hint de enumeração e busca por scroll dirigido).

## Quem chama
- `apply_filter_plan`.

## Quais funções ela chama
- `_find_in_viewport` (inner)
- `_set_scroll_position` (inner)
- `_micro_scroll` (inner)
- `_click_slicer_header_safe` e `_cdp_click` (reativação/fallback)

## Entradas e saídas
- **Entrada:** `tab`, `idx`, `target_value`, contexto do slicer, `enum_hint` opcional, limites de busca.
- **Saída:** `dict` com `found`, coordenadas de clique, estratégia usada, passos e telemetria.

## Ações principais
1. **Camada A:** tenta localizar no viewport sem rolar.
2. **Camada B:** se houver `enum_hint`, tenta pular para posição provável.
3. **Camada C:** busca incremental para frente e para trás com micro-scroll.
4. Retorna ponto clicável e contexto do encontro.

## Pontos de atenção
- Função extensa porque encapsula muitas contingências de virtualização.
- Alterações de heurística podem quebrar apenas certos slicers (regressão difícil de detectar sem logs).

