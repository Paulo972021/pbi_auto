# 🌐 Fluxo geral da feature

```text
Recebe página + templates + catálogo
  ↓
Abre browser/tab uma vez
  ↓
Valida foco por emulação (fallback de foco se necessário)
  ↓
Navega para página alvo
  ↓
Loop de templates:
  ├─ valida template
  ├─ gera código da pasta (page+filtros)
  ├─ prepara pasta de saída
  ├─ aplica FILTER_PLAN/TARGET_PAGE em runtime
  ├─ executa run_export
  ├─ detecta novos arquivos em Downloads
  └─ move arquivos para pasta do template
  ↓
Consolida resultado
  ↓
Fecha browser/tab com graceful shutdown
```

## Objetivo
Reutilizar sessão, reduzir instabilidade em background e organizar arquivos por template.
