# 🧱 Linguagens e bibliotecas utilizadas

## Linguagem principal

## 1) Python
- **Função no script:** linguagem de orquestração da automação (controle de fluxo, integração entre módulos, manipulação de arquivos e execução assíncrona).
- **Onde aparece:** `versao_antiga_com_feature.py`.

---

## Bibliotecas padrão (stdlib)

## 2) `asyncio`
- **Função:** executar o fluxo assíncrono da sessão compartilhada e oferecer wrapper síncrono (`asyncio.run`).
- **Uso no script:** função principal assíncrona e wrapper `_sync`.

## 3) `importlib`
- **Função:** importação dinâmica do módulo principal de automação (`pbi_nico` ou `versao_antiga`).
- **Uso no script:** `_import_pbi_module`.

## 4) `os`
- **Função:** leitura de diretórios/arquivos, montagem de paths e descoberta da pasta de Downloads.
- **Uso no script:** helpers de snapshot e movimentação de arquivos.

## 5) `shutil`
- **Função:** mover arquivos exportados para pasta de saída do template.
- **Uso no script:** `_move_files_to_output`.

## 6) `time`
- **Função:** controle de espera/polling para conclusão de downloads (`.crdownload`, `.tmp`, `.part`).
- **Uso no script:** `_wait_for_downloads_complete`.

## 7) `typing` (`Any`, `Dict`, `List`)
- **Função:** tipagem estática leve para melhorar legibilidade e manutenção.
- **Uso no script:** assinaturas das funções.

---

## Módulos internos do projeto

## 8) `validator`
- **Função:** valida se o template está consistente com o catálogo antes da execução.
- **Uso no script:** `validate_template`.

## 9) `codegen`
- **Função:** gerar identificador de template (`page+filtros`) para nome de pasta.
- **Uso no script:** `generate_template_code`.

## 10) `storage`
- **Função:** criar/limpar pasta de saída de cada template.
- **Uso no script:** `prepare_output_folder`.

---

## Módulo de automação Power BI (dinâmico)

## 11) `pbi_nico` (ou `versao_antiga`)
- **Função:** fornecer o motor de automação (abrir browser, navegar, aplicar foco emulado, exportar).
- **Principais funções consumidas pela versão integrada:**
  - `normalize_browser_path`
  - `start_isolated_browser`
  - `open_report_tab`
  - `ensure_focus_visibility_emulation`
  - `safe_focus_tab`
  - `wait_for_visuals_or_abort`
  - `navigate_to_report_page`
  - `close_open_menus_and_overlays`
  - `run_export`
  - `graceful_browser_shutdown`

