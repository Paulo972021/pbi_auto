# 01) Função crítica: run_templates_for_page_shared_session

## Finalidade
Executar N templates da mesma página em sessão única, com foco estável e captura de arquivos por template.

## Quem chama
- aplicação externa (ou wrapper sync).

## Funções chamadas
- start/open/navigate/wait/focus do módulo PBI
- validate_template
- generate_template_code
- prepare_output_folder
- run_export
- helpers de snapshot/move

## Pontos críticos
- restauração de `FILTER_PLAN` e `TARGET_PAGE` no `finally`;
- encerramento seguro de browser/tab;
- sucesso real depende de `export_ok` + `files_moved`.
