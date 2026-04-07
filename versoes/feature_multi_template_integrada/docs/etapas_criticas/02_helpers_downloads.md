# 02) Funções críticas: helpers de downloads

## Finalidade
Garantir detecção e transferência determinística dos arquivos gerados.

## Funções
- `_get_downloads_folder`
- `_snapshot_downloads`
- `_detect_new_files`
- `_wait_for_downloads_complete`
- `_move_files_to_output`

## Riscos
- concorrência com arquivos externos em Downloads;
- timeout curto para downloads lentos;
- permissões de filesystem.
