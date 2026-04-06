# 📚 Documentação Técnica e Funcional — `primeira_versao_janela_inativa`

> Snapshot da primeira versão usada para execução de automação Power BI sem depender de janela ativa.
>
> Commit de referência: `df25d9d`.

---

## ETAPA 1 — Inventário

### Arquivos analisados e papel

1. `pbi_nico.py`  
   - **Papel:** motor principal de automação (navegação, scan de visuais, aplicação de filtros, exportação).  
   - **Tipo:** núcleo/orquestrador técnico.
2. `executor_real.py`  
   - **Papel:** ponte do pipeline de templates para o motor (`run_export`) com preparo de ambiente e captura de downloads.  
   - **Tipo:** adaptador de integração.
3. `executor_session.py`  
   - **Papel:** execução de múltiplos templates em sessão compartilhada (mesma página).  
   - **Tipo:** orquestrador de lote/sessão.
4. `test_epn_final_emdia_apply.py`  
   - **Papel:** teste diagnóstico focado no caso `epn_final=EMDIA`.  
   - **Tipo:** script de validação dirigida.
5. `test_slicer_micro_scroll.py`  
   - **Papel:** teste de enumeração por micro-scroll no slicer `acao_tatica`.  
   - **Tipo:** script de diagnóstico de enumeração.
6. `test_slicer_registry_hypothesis.py`  
   - **Papel:** teste da hipótese de perda de estado lógico após scroll.  
   - **Tipo:** script de diagnóstico comportamental.

---

## ETAPA 2 — Mapa de arquitetura

### Visão geral estrutural

```text
versoes/primeira_versao_janela_inativa/
├── README.md                               # documentação consolidada
├── executor_real.py                        # integração template → motor PBI (template único)
├── executor_session.py                     # execução em sessão compartilhada por página
├── pbi_nico.py                             # motor técnico central de automação
├── test_epn_final_emdia_apply.py           # teste diagnóstico focado em EMDIA
├── test_slicer_micro_scroll.py             # teste diagnóstico de enumeração por micro-scroll
├── test_slicer_registry_hypothesis.py      # teste diagnóstico de hipótese de estado lógico
└── docs/
    ├── dependencias.md                     # mapa de dependências (arquivos/funções)
    ├── fluxo_geral.md                      # macrofluxo para leigos e técnicos
    ├── microfluxos.md                      # subprocessos e falhas típicas
    └── legenda_navegacao.md                # convenções visuais (emojis/comentários)
```

### Relação operacional entre os arquivos

```text
pipeline (templates/catálogo)
   │
   ├── executor_real.py
   │      └── importa módulo pbi (pbi_auto_v06)
   │             └── (neste snapshot: pbi_nico.py)
   │
   ├── executor_session.py
   │      ├── usa executor_real.py (helpers de download + filter plan)
   │      └── importa módulo pbi (pbi_auto_v06)
   │
   └── scripts de teste isolado
          ├── test_epn_final_emdia_apply.py
          ├── test_slicer_micro_scroll.py
          └── test_slicer_registry_hypothesis.py
```

### Dependências entre arquivos (resumo)

- `executor_session.py` **depende de** `executor_real.py` (helpers de download + montagem de filtros).  
- `executor_real.py` **depende de** `validator`, `codegen`, `storage` e do módulo PBI importado dinamicamente (`pbi_auto_v06`).
- `executor_session.py` **depende de** `validator`, `codegen`, `storage` e do mesmo módulo PBI.
- Os três scripts de teste são **independentes entre si** e importam apenas `pbi_nico` em runtime.

> Inferência explicitada: no fluxo operacional real, `pbi_auto_v06` pode apontar para outro arquivo equivalente ao motor deste snapshot. Aqui documentamos com base no conteúdo local e no contrato de função utilizado pelos executores.

Mais detalhes em: `docs/dependencias.md`.

---

## ETAPA 3 — Mapa de funções

A documentação função a função foi organizada em `docs/dependencias.md` (chamadores/chamadas) e complementada por visão de tarefas em `docs/microfluxos.md`.

Funções centrais por arquivo:

- `executor_real.py`:
  - `run_template_real`, `run_template_real_sync`, `build_runtime_filter_plan_from_template`, helpers de download.
- `executor_session.py`:
  - `run_templates_for_page_in_shared_session`, `_run_single_template_in_session`, `_template_transition_cleanup`, `_run_with_focus_fallback`.
- `pbi_nico.py`:
  - `run_export`, `scan_slicers`, `apply_filter_plan`, `apply_filter_safe_control`, `export_selected_visuals`, `wait_for_visuals_or_abort`, `ensure_focus_visibility_emulation`.
- Testes:
  - `run_test` / `run` e helpers de leitura/interação de slicer.

---

## ETAPA 4 — Fluxo macro

Veja `docs/fluxo_geral.md` para:
- fluxo principal completo;
- versão explicada para leigos;
- versão técnica com nomes reais de funções/arquivos;
- diagrama textual macro.

---

## ETAPA 5 — Microfluxos

Veja `docs/microfluxos.md` para:
- inicialização;
- carregamento de configuração;
- navegação e validação de página;
- aplicação de filtros (incluindo fallback);
- exportação e movimentação de downloads;
- tratamento de erro e encerramento.

---

## ETAPA 6 — Convenção visual

Legenda completa em `docs/legenda_navegacao.md`.

Padrão adotado nos arquivos principais deste snapshot:
- `🚀` entrada/orquestração
- `⚙️` configuração e parâmetros
- `🧩` integração entre módulos
- `🔎` busca/localização
- `🧠` regra de decisão
- `🔁` laços/repetições
- `🛡️` validação e fallback
- `📤` saída/exportação
- `❌` falhas/exceções
- `🧪` testes

---

## ETAPA 7 — Proposta de comentários por seção

Implementada diretamente nos arquivos com blocos de navegação visual:
- `executor_real.py`
- `executor_session.py`
- `pbi_nico.py`
- `test_epn_final_emdia_apply.py`
- `test_slicer_micro_scroll.py`
- `test_slicer_registry_hypothesis.py`

Objetivo: facilitar onboarding e manutenção sem alterar comportamento.

---

## ETAPA 8 — Documentação consolidada

- `README.md` (este arquivo): visão executiva + mapa de leitura.
- `docs/dependencias.md`: dependências entre arquivos/funções e cadeia de chamadas.
- `docs/fluxo_geral.md`: macrofluxo para leigos e técnicos.
- `docs/microfluxos.md`: subprocessos detalhados.
- `docs/legenda_navegacao.md`: convenção visual e padrão de comentários.
- `docs/etapas_criticas/`: detalhamento focado das funções críticas de slicers/filtros e `run_export`.
  - inclui `09_fluxos_detalhados_funcoes_criticas.md` com fluxos passo a passo das funções críticas.

---

## ETAPA 9 — Riscos e observações finais

1. **Arquivo crítico único (`pbi_nico.py`)**: concentra muitas responsabilidades; alterações sem testes podem gerar regressão transversal.  
2. **Dependência de DOM do Power BI**: seletores e heurísticas podem quebrar com mudanças de UI.
3. **Dependência de timing**: waits e retentativas são sensíveis a ambiente/rede.
4. **Dependência de filesystem/downloads**: mover/limpar arquivos exige cuidado para não capturar arquivos externos.
5. **Acoplamento funcional**: `executor_session.py` reutiliza funções internas de `executor_real.py`; alterar assinatura impacta os dois.

### O que não deve ser alterado sem cuidado
- Assinaturas usadas pelos executores para chamar o módulo PBI (`run_export`, scan/apply/export helpers).
- Contrato de FILTER_PLAN e modo de restauração de estado global após execução.
- Sequência de limpeza/transição em sessão compartilhada.

---

## Ordem de leitura recomendada

1. `README.md` (este arquivo)  
2. `docs/legenda_navegacao.md`  
3. `docs/fluxo_geral.md`  
4. `docs/dependencias.md`  
5. `docs/microfluxos.md`  
6. `executor_session.py`  
7. `executor_real.py`  
8. `pbi_nico.py`  
9. scripts de teste (`test_*.py`)
