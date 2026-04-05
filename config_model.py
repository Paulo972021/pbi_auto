"""
config_model.py — Etapa 1: Modelagem de catálogo e templates

CATALOG: define telas disponíveis, filtros por tela, e valores por filtro.
TEMPLATES: lista de execuções planejadas (tela + combinação de filtros).

Estes dados são a fonte de verdade para validação, geração de código,
e execução (mock ou real).
"""

# ---------------------------------------------------------------------------
# CATÁLOGO — telas, filtros e valores disponíveis
# ---------------------------------------------------------------------------
CATALOG = {
    "COMPARATIVO": {
        "label": "Comparativo",
        "filters": {
            "entrante": {
                "label": "Entrante",
                "values": ["0", "1"],
            },
            "epn_final": {
                "label": "EPN Final",
                "values": ["ALMAVIVA", "BELLINATI", "EMDIA"],
            },
        },
    },
}

# ---------------------------------------------------------------------------
# TEMPLATES — execuções planejadas
# ---------------------------------------------------------------------------
TEMPLATES = [
    {
        "template_id": "tpl_001",
        "page": "COMPARATIVO",
        "filters": {
            "entrante": "1",
            "epn_final": "ALMAVIVA",
        },
    },
    {
        "template_id": "tpl_002",
        "page": "COMPARATIVO",
        "filters": {
            "entrante": "0",
            "epn_final": "EMDIA",
        },
    },
    {
        "template_id": "tpl_003",
        "page": "COMPARATIVO",
        "filters": {
            "entrante": "1",
            "epn_final": "BELLINATI",
        },
    },
]
