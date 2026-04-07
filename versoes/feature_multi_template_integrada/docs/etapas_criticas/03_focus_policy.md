# 03) Etapa crítica: política de foco

## Estratégia
1. usar `ensure_focus_visibility_emulation` (DOM-first)
2. fallback para `safe_focus_tab`
3. revalidar foco após fallback

## Motivo
Minimizar dependência de janela ativa para rodar em background.

## Risco
Se o módulo PBI não tiver essas funções, a feature falha por contrato ausente.
