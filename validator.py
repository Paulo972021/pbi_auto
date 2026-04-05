"""
validator.py — Etapa 2: Validação de templates contra o catálogo

Regras:
  - página deve existir no CATALOG
  - cada filtro do template deve existir na página
  - cada valor do filtro deve existir no catálogo
  - template inválido → abortar apenas ele, não o lote
"""


def validate_template(template: dict, catalog: dict) -> dict:
    """
    Valida um template contra o catálogo.

    Retorna:
      {"valid": True/False, "errors": [...]}
    """
    errors = []
    tid = template.get("template_id", "?")
    page = template.get("page", "")
    filters = template.get("filters", {})

    # 1. Página existe?
    if page not in catalog:
        errors.append(f"[{tid}] página '{page}' não existe no catálogo. Disponíveis: {list(catalog.keys())}")
        return {"valid": False, "errors": errors}

    page_def = catalog[page]
    page_filters = page_def.get("filters", {})

    # 2. Filtros existem na página?
    for filter_key, filter_value in filters.items():
        if filter_key not in page_filters:
            errors.append(
                f"[{tid}] filtro '{filter_key}' não existe na página '{page}'. "
                f"Disponíveis: {list(page_filters.keys())}"
            )
            continue

        # 3. Valor existe no filtro?
        allowed_values = page_filters[filter_key].get("values", [])
        if filter_value not in allowed_values:
            errors.append(
                f"[{tid}] valor '{filter_value}' não existe no filtro '{filter_key}'. "
                f"Disponíveis: {allowed_values}"
            )

    return {"valid": len(errors) == 0, "errors": errors}


def validate_all_templates(templates: list, catalog: dict) -> dict:
    """
    Valida todos os templates e retorna resumo.

    Retorna:
      {"results": [{template_id, valid, errors}, ...],
       "total": int, "valid_count": int, "invalid_count": int}
    """
    results = []
    for tpl in templates:
        result = validate_template(tpl, catalog)
        results.append({
            "template_id": tpl.get("template_id", "?"),
            "valid": result["valid"],
            "errors": result["errors"],
        })

    valid_count = sum(1 for r in results if r["valid"])
    invalid_count = len(results) - valid_count

    return {
        "results": results,
        "total": len(results),
        "valid_count": valid_count,
        "invalid_count": invalid_count,
    }
