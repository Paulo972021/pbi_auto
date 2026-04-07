# 🧭 Legenda de Navegação (Emojis e Convenções)

## Emojis padronizados

- 🚀 **Ponto de entrada / execução principal**
- ⚙️ **Configuração e parâmetros**
- 📥 **Entrada de dados/configuração externa**
- 📤 **Saída (exportação, logs, arquivos)**
- 🔎 **Busca/localização de elementos no DOM**
- 🧠 **Regra de decisão de negócio/fluxo**
- 🔁 **Loop/retentativa/polling**
- 🛡️ **Validação e fallback defensivo**
- ❌ **Tratamento de erro/falha esperada**
- 📌 **Trecho crítico de manutenção**
- 🧩 **Integração entre módulos**
- 🧪 **Teste/diagnóstico**
- 📝 **Observação operacional importante**
- 🧹 **Limpeza/encerramento de estado**

## Convenções de comentário por seção

Padrão recomendado (e aplicado nos arquivos principais deste snapshot):

```python
# ============================================================
# 🚀 NOME DA SEÇÃO
# Breve descrição funcional do bloco
# ============================================================
```

## Convenções de manutenção

1. **Não mover funções críticas sem atualizar documentação de dependências.**
2. **Não renomear funções públicas usadas pelos executores sem refatoração coordenada.**
3. **Ao adicionar fallback/retry, comentar objetivo e limite para evitar comportamento oculto.**
4. **Quando inferir comportamento de UI, marcar explicitamente no comentário como heurística.**

