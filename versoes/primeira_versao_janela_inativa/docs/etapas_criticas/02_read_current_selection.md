# 02) `read_current_selection`

## Finalidade
Ler o estado atual de seleção de um slicer específico (valores selecionados, opções visíveis e metadados de localização).

## Quem chama
- `apply_filter_plan`
- `clear_slicer_selection`
- `validate_filter_final`

## Quais funções ela chama
- Leitura direta via `tab.evaluate(...)` para extrair estado do DOM do slicer.

## Entradas e saídas
- **Entrada:** `tab`, `idx` (índice do slicer), `slicer_title`.
- **Saída:** `dict` com `selectedValues`, `availableValues`, flags e coordenadas auxiliares.

## Ações principais
1. Localiza o slicer pelo índice/título.
2. Extrai seleção atual e valores visíveis.
3. Retorna estrutura padronizada para validação/aplicação de filtro.

## Pontos de atenção
- É base para validação de sucesso de clique/clear.
- Pequenas mudanças no formato de retorno quebram múltiplos chamadores.

