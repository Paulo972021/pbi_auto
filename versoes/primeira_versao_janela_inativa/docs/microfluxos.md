# 🔬 Microfluxos (detalhados)

## Microfluxo A — Inicialização

**Objetivo:** preparar ambiente de execução.

**Gatilho:** chamada de `run_template_real`, `run_templates_for_page_in_shared_session` ou scripts `test_*`.

**Etapas:**
1. validar parâmetros essenciais (URL, browser, template);
2. preparar diretório de saída/log;
3. importar módulo PBI com fallback de path.

**Decisões:**
- Se faltar config crítica, encerra com erro.

**Falhas comuns:**
- caminho do navegador inválido;
- módulo PBI não encontrado.

**Resultado esperado:** ambiente pronto para abrir browser/aba.

---

## Microfluxo B — Navegação + carregamento de página

**Objetivo:** abrir relatório e garantir DOM mínimo antes de agir.

**Gatilho:** início do fluxo no módulo PBI.

**Etapas:**
1. `start_isolated_browser`;
2. `open_report_tab`;
3. `wait_for_visuals_or_abort` / `wait_for_visual_containers`;
4. `navigate_to_page` quando necessário.

**Decisões:**
- se não carregar visuais, abortar/recarregar (dependendo da função).

**Falhas comuns:**
- timeout de renderização;
- redirecionamento para aba errada;
- overlays bloqueando interação.

**Resultado esperado:** página estável e pronta para scan de slicers.

---

## Microfluxo C — Foco/visibilidade (janela inativa)

**Objetivo:** reduzir dependência de janela ativa.

**Gatilho:** antes de operações críticas de filtro/export.

**Etapas:**
1. ler estado de foco (`get_page_focus_state`);
2. instalar emulação (`install_focus_visibility_emulation`);
3. validar (`ensure_focus_visibility_emulation`);
4. fallback para `safe_focus_tab` quando necessário.

**Decisões:**
- usar fallback só quando estratégia DOM não estabilizar.

**Falhas comuns:**
- política de navegador/aba bloqueando eventos;
- estado de visibilidade inconsistente.

**Resultado esperado:** operações de UI mais estáveis em background.

---

## Microfluxo D — Scan e aplicação de filtros

**Objetivo:** localizar slicers e aplicar valores do template.

**Gatilho:** após página carregada.

**Etapas:**
1. `scan_slicers` identifica slicers existentes;
2. `build_runtime_filter_plan_from_template` gera plano de aplicação;
3. `apply_filter_plan` orquestra aplicação por slicer;
4. `apply_filter_safe_control` e helpers fazem clear/apply/validate com fallback e micro-scroll.

**Decisões:**
- se `required=True` e valor não encontrado, falha do template;
- se `clear_first=True`, limpa seleção antes de aplicar.

**Falhas comuns:**
- valor fora da janela visível;
- ruído de enumeração;
- atraso de atualização do slicer após clique.

**Resultado esperado:** filtros efetivamente aplicados e validados.

---

## Microfluxo E — Exportação

**Objetivo:** exportar dados dos visuais selecionados.

**Gatilho:** filtros aplicados e visuais escaneados.

**Etapas:**
1. abrir menu de cada visual (`more options`);
2. identificar opção de exportação;
3. abrir dialog de export;
4. confirmar exportação;
5. aguardar download.

**Decisões:**
- repetir tentativa quando menu/dialog não aparece;
- ignorar visual não exportável.

**Falhas comuns:**
- menu não abre;
- botão export indisponível;
- warning/modal bloqueando ação.

**Resultado esperado:** arquivos de export gerados.

---

## Microfluxo F — Captura e organização de downloads

**Objetivo:** separar arquivos do template corrente e evitar resíduos.

**Gatilho:** antes/depois de export no executor.

**Etapas:**
1. limpar resíduos relevantes em `Downloads`;
2. snapshot pré-export;
3. aguardar fim de `.crdownload/.tmp`;
4. detectar novos arquivos;
5. mover para `output` do template.

**Decisões:**
- se arquivo não existir mais, ignora.

**Falhas comuns:**
- download lento além do timeout;
- concorrência com arquivos externos na mesma pasta.

**Resultado esperado:** saída organizada por template.

---

## Microfluxo G — Encerramento e tratamento de erro

**Objetivo:** finalizar com segurança e preservar rastreabilidade.

**Gatilho:** sucesso, falha ou interrupção.

**Etapas:**
1. registrar erro/estado final;
2. restaurar configurações globais do módulo PBI quando aplicável;
3. fechar aba/browser com `graceful_browser_shutdown`;
4. persistir logs e JSON dos testes.

**Resultado esperado:** ambiente limpo para próxima execução.

