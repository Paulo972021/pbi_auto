"""
Integração conservadora de feature:
- múltiplos templates na mesma página/sessão
- foco emulado com fallback
- captura e transferência de arquivos para pastas por template (page+filters)

Este arquivo NÃO altera os módulos principais existentes.
"""

import asyncio
import importlib
import os
import shutil
import time
from typing import Any, Dict, List

from validator import validate_template
from codegen import generate_template_code
from storage import prepare_output_folder

# ============================================================
# 🧩 CARREGAMENTO DE MÓDULO PBI
# ============================================================

def _import_pbi_module(module_name: str = "pbi_nico"):
    try:
        return importlib.import_module(module_name)
    except Exception:
        return None


def _get_downloads_folder() -> str:
    return os.path.join(os.path.expanduser("~"), "Downloads")


def _snapshot_downloads(downloads_folder: str) -> set:
    try:
        return set(os.listdir(downloads_folder))
    except Exception:
        return set()


def _detect_new_files(before_files: set, after_files: set) -> list:
    return sorted(list(after_files - before_files))


def _wait_for_downloads_complete(downloads_folder: str, max_wait: int = 20) -> None:
    start = time.time()
    transient_suffixes = (".crdownload", ".part", ".tmp")
    while time.time() - start < max_wait:
        try:
            files = os.listdir(downloads_folder)
        except Exception:
            return
        if not any(name.lower().endswith(transient_suffixes) for name in files):
            return
        time.sleep(0.5)


def _move_files_to_output(new_files: list, downloads_folder: str, output_folder: str) -> list:
    os.makedirs(output_folder, exist_ok=True)
    moved = []
    for name in new_files:
        src = os.path.join(downloads_folder, name)
        if not os.path.isfile(src):
            continue
        dst = os.path.join(output_folder, name)
        try:
            shutil.move(src, dst)
            moved.append(name)
        except Exception:
            continue
    return moved


def build_runtime_filter_plan_from_template(template: dict) -> dict:
    filters = template.get("filters", {})
    plan = {}
    for filter_key, filter_value in filters.items():
        normalized_key = str(filter_key).strip().lower()
        if isinstance(filter_value, list):
            target_values = [str(v) for v in filter_value]
            mode = "multi" if len(target_values) > 1 else "single"
        else:
            target_values = [str(filter_value)]
            mode = "single"
        plan[normalized_key] = {
            "mode": mode,
            "clear_first": True,
            "target_values": target_values,
            "required": True,
        }
    return plan


async def run_templates_for_page_shared_session(
    page: str,
    templates: List[Dict[str, Any]],
    catalog: Dict[str, Any],
    pbi_module_name: str = "pbi_nico",
) -> Dict[str, Any]:
    """
    🚀 Função principal da feature integrada.
    """
    pbi = _import_pbi_module(pbi_module_name)
    if pbi is None:
        return {"page": page, "error": f"module_not_found:{pbi_module_name}", "results": []}

    url = getattr(pbi, "POWERBI_URL", "")
    browser_path = getattr(pbi, "BROWSER_PATH", "")
    if not url or not browser_path:
        return {"page": page, "error": "missing_config", "results": []}

    results = []
    success = failed = skipped = total_files = 0
    browser = None
    tab = None
    owned_tab_refs = set()

    try:
        browser_path_norm = pbi.normalize_browser_path(browser_path)
        browser, _ = await pbi.start_isolated_browser(browser_path_norm)
        tab = await pbi.open_report_tab(browser, url, owned_tab_refs)

        focus_ok = await pbi.ensure_focus_visibility_emulation(tab, stage_label="shared_session_open")
        if not focus_ok:
            await pbi.safe_focus_tab(tab)
            await pbi.ensure_focus_visibility_emulation(tab, stage_label="shared_session_open_after_fallback")

        await pbi.close_extra_tabs_created_by_script(browser, tab, owned_tab_refs)
        await pbi.allow_multiple_downloads(tab)

        loaded = await pbi.wait_for_visuals_or_abort(tab, stage_label="shared_session_init", retries=10, wait_seconds=3)
        if not loaded:
            return {"page": page, "error": "report_not_loaded", "results": []}

        if page:
            nav_ok = await pbi.navigate_to_report_page(tab, page)
            if not nav_ok:
                return {"page": page, "error": f"navigation_failed:{page}", "results": []}

            loaded2 = await pbi.wait_for_visuals_or_abort(tab, stage_label=f"navigate_{page}", retries=8, wait_seconds=3)
            if not loaded2:
                return {"page": page, "error": "visuals_not_loaded_after_navigation", "results": []}

        await pbi.close_open_menus_and_overlays(tab, aggressive=True)

        downloads_folder = _get_downloads_folder()

        for tpl in templates:
            tid = tpl.get("template_id", "?")
            validation = validate_template(tpl, catalog)
            if not validation.get("valid"):
                results.append({
                    "template_id": tid,
                    "success": False,
                    "skipped": True,
                    "error": "; ".join(validation.get("errors", [])),
                })
                skipped += 1
                continue

            template_code = generate_template_code(tpl)
            output_folder = prepare_output_folder(template_code)

            before = _snapshot_downloads(downloads_folder)
            runtime_plan = build_runtime_filter_plan_from_template(tpl)
            original_plan = getattr(pbi, "FILTER_PLAN", {})
            original_page = getattr(pbi, "TARGET_PAGE", "")

            export_ok = False
            moved = []
            error = None

            try:
                pbi.FILTER_PLAN = runtime_plan
                pbi.TARGET_PAGE = page
                export_ok = await pbi.run_export(url, browser_path_norm, page, stop_after_filters=False)

                _wait_for_downloads_complete(downloads_folder, max_wait=20)
                after = _snapshot_downloads(downloads_folder)
                new_files = _detect_new_files(before, after)
                moved = _move_files_to_output(new_files, downloads_folder, output_folder)
            except Exception as e:
                error = str(e)
            finally:
                pbi.FILTER_PLAN = original_plan
                pbi.TARGET_PAGE = original_page

            tpl_success = bool(export_ok and len(moved) > 0 and not error)
            if tpl_success:
                success += 1
            else:
                failed += 1

            total_files += len(moved)
            results.append({
                "template_id": tid,
                "template_code": template_code,
                "output_folder": output_folder,
                "export_ok": bool(export_ok),
                "files_moved": len(moved),
                "files_list": moved,
                "success": tpl_success,
                "skipped": False,
                "error": error,
            })

        return {
            "page": page,
            "results": results,
            "success": success,
            "failed": failed,
            "skipped": skipped,
            "total_files": total_files,
        }
    finally:
        if tab is not None or browser is not None:
            try:
                await pbi.graceful_browser_shutdown(tab, browser)
            except Exception:
                pass


def run_templates_for_page_shared_session_sync(
    page: str,
    templates: List[Dict[str, Any]],
    catalog: Dict[str, Any],
    pbi_module_name: str = "pbi_nico",
) -> Dict[str, Any]:
    return asyncio.run(run_templates_for_page_shared_session(page, templates, catalog, pbi_module_name))
