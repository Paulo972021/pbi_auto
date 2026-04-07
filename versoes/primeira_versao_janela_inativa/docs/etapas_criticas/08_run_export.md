# 08) `run_export`

## Finalidade
Orquestrar o processo completo: abrir browser/relatório, navegar, aplicar filtros, exportar visuais e encerrar com limpeza.

## Quem chama
- `executor_real.py` (`run_template_real`) chama `pbi.run_export(...)`.
- `main()` de `pbi_nico.py` também chama `run_export`.

## Quais funções ela chama (macro)
- Inicialização: `start_isolated_browser`, `open_report_tab`
- Robustez de foco: `ensure_focus_visibility_emulation`, `safe_focus_tab`
- Estabilidade: `wait_for_visuals_or_abort`, `cleanup_residual_ui`, `close_open_menus_and_overlays`
- Filtros: `scan_slicers` (que internamente chama `apply_filter_plan`)
- Exportação: `scan_visuals`, `ask_user_visual_selection`, `export_selected_visuals`
- Encerramento: `graceful_browser_shutdown`

## Entradas e saídas
- **Entrada:** `url`, `browser_path`, `target_page`, `stop_after_filters`.
- **Saída:** `bool` de sucesso do processo.

## Ações principais
1. Abre browser isolado e tab do relatório.
2. Garante política de foco/visibilidade para background.
3. Navega para página alvo e valida visual-containers.
4. Executa etapa de filtros via `scan_slicers`.
5. Executa etapa de exportação.
6. Encerra sessão de forma graciosa no `finally`.

## Pontos de atenção
- Função longa porque agrega o fluxo de ponta a ponta.
- Qualquer alteração de ordem entre "limpeza -> scan -> filtro -> export" pode gerar efeito cascata.

