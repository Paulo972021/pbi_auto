# 09) Fluxo detalhado das funções críticas

> Objetivo: visualizar em detalhe como cada função crítica se comporta, quais decisões toma, quais dependências chama e onde pode falhar.

---

## 1. `scan_slicers`

### Fluxo detalhado (passo a passo)

```text
Início
  ↓
Cleanup inicial de UI (overlays/menus)
  ↓
Coleta contexto de DOM (pré-scan)
  ↓
Executa varredura de slicers no relatório
  ↓
Para cada slicer encontrado:
  ├─ normaliza título/campo
  ├─ tenta enumeração via micro-scroll (_enumerate_slicer_via_micro_scroll)
  ├─ verifica se há regra no FILTER_PLAN
  └─ se houver regra: chama apply_filter_plan
  ↓
Atualiza estrutura final de slicers
  ↓
Pressiona ESC para normalizar UI
  ↓
Retorna lista consolidada
```

### Decisões críticas
- se não achar slicers válidos, retorna lista vazia;
- se enumeração falhar, segue com fallback de leitura atual;
- se houver FILTER_PLAN para o slicer, dispara aplicação de filtro.

### Possíveis falhas
- DOM virtualizado sem elementos esperados;
- overlay bloqueando interação;
- atraso de render após abrir slicer.

---

## 2. `read_current_selection`

### Fluxo detalhado

```text
Início
  ↓
Localiza slicer por índice/título
  ↓
Lê selectedValues + availableValues + metadados
  ↓
Normaliza estrutura
  ↓
Retorna dict de estado atual
```

### Decisões críticas
- se o slicer não for localizado, retorna estado vazio/sinal de ausência.

### Possíveis falhas
- seletor de título não refletir o visual correto;
- leitura parcial quando a UI ainda está em transição.

---

## 3. `clear_slicer_selection`

### Fluxo detalhado

```text
Início
  ↓
Recebe estado atual (state)
  ↓
Tenta clear principal (botão/ação explícita)
  ↓
Releitura via read_current_selection
  ↓
[Decisão] Limpou?
  ├─ Sim → retorna True
  └─ Não → tenta fallback (clique em selecionado)
            ↓
        Revalida estado
            ↓
        retorna True/False
```

### Decisões críticas
- clear_first ativo no plano de filtro chama esta função;
- sem confirmação por releitura, não marca clear como efetivo.

### Possíveis falhas
- clique de clear não dispara evento de seleção;
- estado visual muda, mas estado lógico não.

---

## 4. `_enumerate_slicer_via_micro_scroll`

### Fluxo detalhado

```text
Início
  ↓
Ativa cabeçalho do slicer (_click_slicer_header_safe)
  ↓
Detecta container de scroll (_find_scroll_container)
  ↓
Captura snapshot inicial (selecionados + caixa visível)
  ↓
Loop de micro-scroll:
  ├─ executa _micro_scroll_step
  ├─ lê caixa visível (_read_visible_box)
  ├─ acumula valores novos
  ├─ detecta estagnação/limite
  └─ decide continuar/parar
  ↓
Consolida valores + diagnóstico
  ↓
Retorna enum_result
```

### Decisões críticas
- interrompe por estagnação (sem novos itens), limite de passos ou timeout;
- registra first/last scroll positions para suporte ao locator posterior.

### Possíveis falhas
- container incorreto selecionado para scroll;
- ruído textual na caixa visível;
- virtualização agressiva omitindo itens durante leitura.

---

## 5. `_scroll_to_find_value_in_slicer`

### Fluxo detalhado por camadas

```text
Início
  ↓
Camada A: procura no viewport atual (_find_in_viewport)
  ↓
[Decisão] Encontrou?
  ├─ Sim → retorna coordenadas
  └─ Não → tenta reativar cabeçalho e buscar de novo
             ↓
             Camada B: usa enum_hint (posição provável)
             ↓
             [Decisão] Encontrou?
               ├─ Sim → retorna coordenadas
               └─ Não → Camada C (scroll search incremental)
                         ├─ varredura forward
                         ├─ varredura backward
                         └─ retorna found=False se exaurir limites
```

### Decisões críticas
- ordem de camadas reduz custo de busca (rápido -> dirigido -> exaustivo);
- usa scrollTop conhecido da enumeração para pulo eficiente.

### Possíveis falhas
- hint desatualizado após mudança de render;
- valor existe mas não entra na janela por limitação de interação.

---

## 6. `apply_filter_plan`

### Fluxo detalhado

```text
Início
  ↓
Ativa slicer (header/fallback)
  ↓
Lê estado atual (read_current_selection)
  ↓
[Decisão] clear_first?
  ├─ Sim → clear_slicer_selection + releitura
  └─ Não → mantém seleção atual
  ↓
Para cada target_value:
  ├─ localiza via _scroll_to_find_value_in_slicer
  ├─ clica no ponto encontrado (_cdp_click)
  ├─ releitura de seleção
  └─ marca sucesso/falha do item
  ↓
Validação final (validate_filter_final)
  ↓
Retorna relatório completo de aplicação
```

### Decisões críticas
- se `required=True` e valor não aplicado, resultado final deve refletir falha;
- mantém logs diagnósticos por tentativa/alvo.

### Possíveis falhas
- localização encontra item visual, mas click não altera seleção;
- atraso de atualização após click gera falso negativo sem wait suficiente.

---

## 7. `apply_filter_safe_control` (não ativo no fluxo principal)

### Fluxo detalhado

```text
Início
  ↓
Ativa slicer
  ↓
Lê estado com _read_filter_state
  ↓
[Decisão] clear_first?
  ├─ Sim → _safe_clear_filter
  └─ Não → segue
  ↓
Aplica cada target via _apply_single_filter_value
  ↓
Revalida estado final
  ↓
press_escape para fechar UI residual
  ↓
Retorna relatório
```

### Situação de uso no snapshot
- Definida, documentada, porém sem chamada ativa no caminho `run_export -> scan_slicers`.

### Risco funcional
- função longa sem uso efetivo tende a divergir do fluxo real ao longo do tempo.

---

## 8. `run_export`

### Fluxo detalhado ponta a ponta

```text
Início
  ↓
start_isolated_browser
  ↓
open_report_tab
  ↓
ensure_focus_visibility_emulation (fallback safe_focus_tab se necessário)
  ↓
wait_for_visuals_or_abort
  ↓
navigate_to_report_page (quando TARGET_PAGE)
  ↓
cleanup_residual_ui pré-filtros
  ↓
scan_slicers (aplica FILTER_PLAN internamente)
  ↓
post-filter cleanup / validações
  ↓
scan_visuals
  ↓
ask_user_visual_selection
  ↓
export_selected_visuals
  ↓
finally: graceful_browser_shutdown
  ↓
retorna bool de sucesso
```

### Decisões críticas
- aborta quando não há visual-container ou carregamento falha;
- tratamento de foco em modo emulação primeiro e foco explícito só no fallback.

### Possíveis falhas
- bloqueios de popup/dialog;
- falha de leitura do estado de filtro;
- export dialog não abrir para visuais específicos.

---

## 9. Visão de dependência entre as funções críticas

```text
run_export
  └─ scan_slicers
       ├─ _enumerate_slicer_via_micro_scroll
       └─ apply_filter_plan
            ├─ read_current_selection
            ├─ clear_slicer_selection
            │    └─ read_current_selection
            └─ _scroll_to_find_value_in_slicer

apply_filter_safe_control  (isolada no snapshot)
  ├─ _read_filter_state
  ├─ _safe_clear_filter
  └─ _apply_single_filter_value
```

