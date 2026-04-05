"""
codegen.py — Etapa 3: Geração automática de código do template

Gera um identificador único e legível para cada template,
usado como nome de pasta de saída e referência em logs.

Formato: PAGINA__filtro1-valor1__filtro2-valor2
Exemplo: COMPARATIVO__entrante-1__epn_final-ALMAVIVA
"""

import re


def _sanitize(text: str) -> str:
    """Remove caracteres problemáticos para nomes de pasta."""
    text = text.strip()
    text = re.sub(r'[^\w\-.]', '_', text)
    text = re.sub(r'_+', '_', text)
    return text.strip('_')


def generate_template_code(template: dict) -> str:
    """
    Gera código legível do template.

    Entrada:
      {"page": "COMPARATIVO", "filters": {"entrante": "1", "epn_final": "ALMAVIVA"}}

    Saída:
      "COMPARATIVO__entrante-1__epn_final-ALMAVIVA"
    """
    page = _sanitize(template.get("page", "UNKNOWN"))

    filter_parts = []
    for key in sorted(template.get("filters", {}).keys()):
        value = template["filters"][key]
        filter_parts.append(f"{_sanitize(key)}-{_sanitize(str(value))}")

    if filter_parts:
        return f"{page}__{'__'.join(filter_parts)}"
    return page
