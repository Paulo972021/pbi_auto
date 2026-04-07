# 🔗 Dependências

## Arquivo principal
- `versao_antiga_com_feature.py`

## Dependências diretas
- `validator.validate_template`
- `codegen.generate_template_code`
- `storage.prepare_output_folder`
- módulo PBI carregado dinamicamente (`pbi_nico` ou `versao_antiga`)

## Cadeia de chamada principal
```text
run_templates_for_page_shared_session
  ├─ _import_pbi_module
  ├─ pbi.start_isolated_browser
  ├─ pbi.open_report_tab
  ├─ pbi.ensure_focus_visibility_emulation
  ├─ pbi.wait_for_visuals_or_abort
  ├─ loop templates:
  │   ├─ validate_template
  │   ├─ generate_template_code
  │   ├─ prepare_output_folder
  │   ├─ build_runtime_filter_plan_from_template
  │   ├─ pbi.run_export
  │   ├─ _snapshot_downloads / _detect_new_files
  │   └─ _move_files_to_output
  └─ pbi.graceful_browser_shutdown
```
