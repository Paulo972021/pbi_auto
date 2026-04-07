# 04) `_enumerate_slicer_via_micro_scroll`

## Finalidade
Enumerar valores de slicer com scroll incremental para lidar com virtualização e listas longas.

## Quem chama
- `scan_slicers`.

## Quais funções ela chama
- `_click_slicer_header_safe`
- `_find_scroll_container`
- `_read_slicer_selected_snapshot`
- `_read_visible_box`
- `_micro_scroll_step`
- `_cdp_click` (reativação quando necessário)

## Entradas e saídas
- **Entrada:** `tab`, `idx`, `slicer_title`, `field_name`, parâmetros de limite/step.
- **Saída:** `dict` com valores encontrados, selecionados, passos, sucesso, condições de parada e diagnóstico.

## Ações principais
1. Ativa/garante foco do slicer.
2. Detecta container de rolagem válido.
3. Lê janela visível inicial.
4. Executa loop de micro-scroll com deduplicação e heurísticas de estagnação.
5. Retorna conjunto consolidado para apoiar localização de target.

## Pontos de atenção
- Função grande e cara (tempo de execução), porém útil para reduzir falso "valor não encontrado".
- Depende de heurísticas de scroll/visibilidade.

