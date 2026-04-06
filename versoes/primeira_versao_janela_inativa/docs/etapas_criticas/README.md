# 📌 Etapas Críticas — Slicers/Filtros e `run_export`

Esta pasta documenta as funções críticas pedidas:

- `scan_slicers`
- `read_current_selection`
- `clear_slicer_selection`
- `_enumerate_slicer_via_micro_scroll`
- `_scroll_to_find_value_in_slicer`
- `apply_filter_plan`
- `apply_filter_safe_control`
- `run_export`

## Navegação rápida

1. [`01_scan_slicers.md`](./01_scan_slicers.md)
2. [`02_read_current_selection.md`](./02_read_current_selection.md)
3. [`03_clear_slicer_selection.md`](./03_clear_slicer_selection.md)
4. [`04_enumerate_slicer_via_micro_scroll.md`](./04_enumerate_slicer_via_micro_scroll.md)
5. [`05_scroll_to_find_value_in_slicer.md`](./05_scroll_to_find_value_in_slicer.md)
6. [`06_apply_filter_plan.md`](./06_apply_filter_plan.md)
7. [`07_apply_filter_safe_control.md`](./07_apply_filter_safe_control.md)
8. [`08_run_export.md`](./08_run_export.md)

## Observação importante sobre chamada

- `apply_filter_safe_control` **está definida**, mas não possui chamada ativa neste snapshot.
- O fluxo efetivo de filtros passa por `scan_slicers -> apply_filter_plan`.

