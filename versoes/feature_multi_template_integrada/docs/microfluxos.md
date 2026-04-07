# 🔬 Microfluxos

## A) Inicialização de sessão
- Importa módulo PBI.
- Valida URL e caminho do browser.
- Inicia browser/tab.

## B) Política de foco
- Tenta `ensure_focus_visibility_emulation`.
- Se falhar, aplica `safe_focus_tab` e revalida.

## C) Execução por template
- Valida template.
- Gera pasta de saída por `page+filtros`.
- Aplica runtime plan e executa `run_export`.

## D) Captura de arquivos
- Snapshot antes/depois em Downloads.
- Espera finalização de arquivos transitórios.
- Move arquivos novos para pasta do template.

## E) Encerramento
- Restaura `FILTER_PLAN` e `TARGET_PAGE`.
- Fecha sessão com `graceful_browser_shutdown`.
