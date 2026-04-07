# 🔗 Dependências entre Arquivos e Funções

## 1) Dependências entre arquivos

```text
executor_session.py
  ├─ imports diretos: validator, codegen, storage
  ├─ imports diretos: executor_real (helpers de download + filter plan)
  └─ import dinâmico: pbi_auto_v06 (contrato de funções do motor PBI)

executor_real.py
  ├─ imports diretos: validator, codegen, storage
  └─ import dinâmico: pbi_auto_v06 (run_export)

scripts de teste (3 arquivos)
  └─ import em runtime: pbi_nico
```

### Dependências centrais
- `pbi_nico.py` (ou módulo equivalente `pbi_auto_v06`) é o **núcleo**.
- `executor_real.py` e `executor_session.py` são **orquestradores de uso** do núcleo.

### Dependências de apoio
- `validator`, `codegen`, `storage` (pré-processamento de template e saída).
- bibliotecas padrão (`asyncio`, `json`, `os`, etc.).

### Acoplamento observado
- **Moderado para alto** entre executores e contrato do módulo PBI.
- `executor_session.py` também acopla em funções auxiliares de `executor_real.py`.

---

## 2) Cadeia de chamadas (quem chama quem)

## 2.1 Fluxo real (template único)

```text
run_template_real_sync
  └─ asyncio.run(run_template_real)
      ├─ validate_template
      ├─ prepare_output_folder
      ├─ build_runtime_filter_plan_from_template
      ├─ _import_pbi_module
      └─ pbi.run_export
```

## 2.2 Fluxo sessão compartilhada (N templates)

```text
run_templates_for_page_in_shared_session
  ├─ _import_pbi
  ├─ pbi.start_isolated_browser / open_report_tab / wait_for_visuals_or_abort
  ├─ loop templates:
  │   ├─ _template_transition_cleanup (a partir do 2º template)
  │   └─ _run_single_template_in_session
  │       ├─ validate_template / prepare_output_folder
  │       ├─ build_runtime_filter_plan_from_template
  │       ├─ pbi.scan_slicers
  │       ├─ pbi.apply_filter_plan
  │       ├─ pbi.export_selected_visuals
  │       └─ helpers de downloads (executor_real)
  └─ pbi.graceful_browser_shutdown
```

## 2.3 Fluxo motor PBI (macro)

```text
run_export
  ├─ start_isolated_browser
  ├─ open_report_tab
  ├─ wait_for_visuals_or_abort
  ├─ navigate_to_report_page
  ├─ scan_slicers
  ├─ apply_filter_plan / apply_filter_safe_control
  ├─ scan_visuals
  ├─ export_selected_visuals
  └─ graceful_browser_shutdown
```

---

## 3) Mapa de funções relevantes (por arquivo)

## 3.1 `executor_real.py`

### `_get_downloads_folder`
- **Propósito:** resolve pasta Downloads do usuário.
- **Chamador:** `run_template_real`.
- **Efeito colateral:** nenhum.
- **Risco:** baixo.

### `_snapshot_downloads` / `_detect_new_files` / `_wait_for_downloads_complete` / `_move_files_to_output` / `_clean_downloads_before_template`
- **Propósito:** ciclo de observação/limpeza/captura de exportações.
- **Chamador:** `run_template_real` e (indiretamente) `executor_session.py` via import.
- **Efeitos colaterais:** leitura/escrita/remoção de arquivos no filesystem.
- **Risco:** médio (pode mover/remover arquivo indevido se regra de extensão mudar).

### `build_runtime_filter_plan_from_template`
- **Entrada:** `template` (dict com filtros declarados).
- **Saída:** `FILTER_PLAN` normalizado para módulo PBI.
- **Chamadores:** `run_template_real`, `_run_single_template_in_session`.
- **Risco:** alto (quebra aplicação de filtros se formato divergir).

### `run_template_real`
- **Tipo:** função principal do arquivo.
- **Responsabilidade:** validar template, preparar saída, configurar runtime no módulo PBI e chamar `run_export`.
- **Retorno:** dict de resultado estruturado.
- **Risco:** alto (ponto de integração principal).

### `run_template_real_sync`
- **Tipo:** wrapper síncrono para uso externo.
- **Risco:** baixo.

---

## 3.2 `executor_session.py`

### `_import_pbi`
- **Propósito:** import dinâmico com fallback de path.
- **Chamador:** `run_templates_for_page_in_shared_session`.
- **Risco:** médio.

### `_run_with_focus_fallback`
- **Propósito:** executar operação sem foco explícito; em falha, aplica `safe_focus_tab` e tenta novamente.
- **Chamadores:** fluxo interno de execução de template.
- **Risco:** médio-alto (impacta robustez quando janela não está ativa).

### `_template_transition_cleanup`
- **Propósito:** limpeza entre templates na mesma sessão (ESC, overlays, popups, estabilidade de visuais).
- **Efeito colateral:** interações de UI e espera.
- **Risco:** alto (se pular etapa, estado residual pode contaminar próximo template).

### `_run_single_template_in_session`
- **Propósito:** executar um template dentro da sessão já aberta.
- **Chamador:** `run_templates_for_page_in_shared_session`.
- **Chamadas principais:** scan/apply/export + captura de downloads.
- **Risco:** alto.

### `run_templates_for_page_in_shared_session`
- **Tipo:** função principal do arquivo.
- **Propósito:** orquestrar sessão compartilhada por página.
- **Risco:** alto (coordena início/fim e repetição por template).

---

## 3.3 `pbi_nico.py` (núcleo)

> Arquivo com alta densidade funcional. Abaixo, mapa das funções mais críticas para manutenção.

### Inicialização/ambiente
- `validate_runtime_config`, `normalize_browser_path`, `build_browser_args`.
- `start_isolated_browser`, `open_report_tab`, `safe_focus_tab`.

### Foco/visibilidade (execução em background)
- `get_page_focus_state`, `install_focus_visibility_emulation`, `ensure_focus_visibility_emulation`.

### Navegação e estabilidade
- `navigate_to_page`, `wait_for_visual_containers`, `wait_for_visuals_or_abort`, `cleanup_residual_ui`, `dismiss_sensitive_data_popup`.

### Slicers e filtros
- `scan_slicers`, `read_current_selection`, `clear_slicer_selection`, `_enumerate_slicer_via_micro_scroll`, `_scroll_to_find_value_in_slicer`, `apply_filter_plan`, `apply_filter_safe_control`.

### Exportação
- `scan_visuals`, `open_visual_more_options`, `click_export_data_menuitem`, `wait_export_dialog`, `confirm_export_dialog`, `export_single_visual`, `export_selected_visuals`.

### Encerramento
- `graceful_browser_shutdown`, `main`.

**Risco de alteração:** muito alto em qualquer função que faça ponte entre seleção de slicer e export.

---

## 3.4 Scripts de teste

### `test_epn_final_emdia_apply.py`
- Foco: cenário único `epn_final=EMDIA` (ativar, limpar, aplicar e validar).
- Entrada: `PBI_REPORT_URL` / `PBI_TAB_NAME`.
- Saída: `.log` + `.json`.

### `test_slicer_micro_scroll.py`
- Foco: enumeração incremental com micro-scroll.
- Entrada: URL e slicer configurável.
- Saída: `.log` + `.json` com diffs de enumeração.

### `test_slicer_registry_hypothesis.py`
- Foco: hipótese de estado lógico persistente vs janela visível.
- Entrada: URL, aba, slicer e targets A/B.
- Saída: `.log` + `.json` com evidências por passo.

