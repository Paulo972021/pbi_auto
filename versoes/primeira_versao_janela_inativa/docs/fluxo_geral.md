# 🌐 Fluxo Macro (Visão Leiga + Técnica)

## 1) Visão leiga

Pense no sistema como um robô que abre um relatório no navegador, escolhe filtros, exporta dados e organiza os arquivos.

1. Ele recebe um "template" com instruções de filtro.
2. Abre o Power BI e entra na aba certa.
3. Confere se a página realmente carregou.
4. Procura os filtros (slicers), aplica os valores desejados.
5. Procura os visuais que podem ser exportados.
6. Exporta e guarda os arquivos no lugar certo.
7. Fecha tudo com segurança.

### Problema que essa versão resolve
Ela melhora a confiabilidade quando a janela do navegador **não está ativa** (em segundo plano), reduzindo falhas de foco/visibilidade.

---

## 2) Visão técnica

### Entradas principais
- URL do relatório (`POWERBI_URL`/`PBI_REPORT_URL`)
- Caminho do navegador (`BROWSER_PATH`)
- Página alvo (`TARGET_PAGE`)
- `FILTER_PLAN` derivado de template

### Núcleo técnico
- `executor_real.py` / `executor_session.py` fazem a orquestração.
- `pbi_nico.py` executa navegação, scan, filtro e export.

### Diagrama textual (macro)

```text
Template/Catálogo
  ↓
executor_real.py ou executor_session.py
  ↓
Importa módulo PBI (pbi_auto_v06 / pbi_nico)
  ↓
Abre browser isolado + abre relatório
  ↓
Valida visuais carregados
  ↓
Navega para página alvo
  ↓
Escaneia slicers
  ↓
Aplica FILTER_PLAN (com fallback e validação)
  ↓
Escaneia visuais exportáveis
  ↓
Executa exportação
  ↓
Captura/organiza downloads
  ↓
Encerra sessão com limpeza
```

---

## 3) Papéis dos arquivos mais importantes

- **Ponto de entrada funcional (batch/sessão):** `executor_session.py`
- **Ponto de entrada funcional (template único):** `executor_real.py`
- **Motor técnico principal:** `pbi_nico.py`
- **Scripts auxiliares de diagnóstico:** `test_*.py`

