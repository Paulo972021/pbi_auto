"""
batch_runner_real.py — Execução em lote LIVE com sessão compartilhada por página

Arquitetura:
  1 página = 1 browser/session
  N templates = mesma sessão → cleanup entre eles

NÃO altera: lógica de exportação, lógica de filtros, pbi_auto_v06.py
"""

import asyncio
import logging
import time
from collections import OrderedDict

from validator import validate_template
from codegen import generate_template_code
from storage import prepare_output_folder
from executor_session import run_templates_for_page_in_shared_session

log = logging.getLogger("batch_runner_real")


def group_templates_by_page(templates: list) -> OrderedDict:
    """Agrupa templates por página, mantendo ordem de inserção."""
    groups = OrderedDict()
    for tpl in templates:
        page = tpl.get("page", "UNKNOWN")
        if page not in groups:
            groups[page] = []
        groups[page].append(tpl)
    return groups


def run_all_templates_live(
    templates: list,
    catalog: dict,
    pbi_module_name: str = "pbi_auto_v06",
) -> dict:
    """
    Executa todos os templates em lote.

    Agrupa por página → 1 sessão por página → N templates na mesma sessão.
    """
    groups = group_templates_by_page(templates)
    all_results = []
    total_success = 0
    total_failed = 0
    total_skipped = 0
    total_files = 0
    batch_start = time.time()

    by_page = OrderedDict()

    log.info(f"[BATCH_LIVE_START]")
    log.info(f"  total_templates={len(templates)}")
    log.info(f"  pages={list(groups.keys())}")
    log.info(f"  session_mode=shared_per_page")

    print(f"\n{'=' * 70}")
    print(f"📦 EXECUÇÃO LIVE EM LOTE (sessão por página)")
    print(f"{'=' * 70}")
    print(f"  Templates: {len(templates)}")
    print(f"  Páginas: {list(groups.keys())}")
    for page, tpls in groups.items():
        tids = [t.get("template_id", "?") for t in tpls]
        print(f"    📄 {page}: {tids} ({len(tids)} templates, 1 sessão)")
    print()

    for page, page_templates in groups.items():
        log.info(f"[BATCH_LIVE_PAGE] page={page} templates_count={len(page_templates)} session_mode=shared")

        print(f"\n{'─' * 60}")
        print(f"📄 PÁGINA: {page} ({len(page_templates)} templates — sessão compartilhada)")
        print(f"{'─' * 60}")

        page_start = time.time()

        # Executar todos os templates desta página em sessão compartilhada
        page_result = asyncio.run(
            run_templates_for_page_in_shared_session(
                page, page_templates, catalog, pbi_module_name
            )
        )

        page_duration = round(time.time() - page_start, 1)

        page_success = page_result.get("success", 0)
        page_failed = page_result.get("failed", 0)
        page_skipped = page_result.get("skipped", 0)
        page_files = page_result.get("total_files", 0)

        total_success += page_success
        total_failed += page_failed
        total_skipped += page_skipped
        total_files += page_files

        by_page[page] = {
            "total": len(page_templates),
            "success": page_success,
            "failed": page_failed,
            "skipped": page_skipped,
            "files": page_files,
            "duration_s": page_duration,
            "sessions": 1,
        }

        # Acumular resultados individuais
        for r in page_result.get("results", []):
            all_results.append(r)

            tid = r.get("template_id", "?")
            is_ok = r.get("success", False)
            fm = r.get("files_moved", 0)
            icon = "✅" if is_ok else ("⏭️" if r.get("skipped") else "❌")
            err = r.get("error", "")

            log.info(f"[BATCH_LIVE_TEMPLATE_END] template_id={tid} success={is_ok} files_moved={fm}")

            line = f"  {icon} {tid}"
            tc = r.get("template_code")
            if tc:
                line += f" → {tc}"
            if fm:
                line += f" ({fm} arquivos)"
            if err:
                line += f" [{err[:50]}]"
            print(line)

        if page_result.get("error"):
            print(f"  ⚠️ Erro de sessão: {page_result['error']}")

        print(f"  ⏱️ Página {page}: {page_duration}s ({page_success}✅ {page_failed}❌ {page_skipped}⏭️)")

    batch_duration = round(time.time() - batch_start, 1)

    # ── BATCH_LIVE_SUMMARY ──
    total_sessions = len(groups)
    log.info(f"[BATCH_LIVE_SUMMARY]")
    log.info(f"  total={len(templates)}")
    log.info(f"  success={total_success}")
    log.info(f"  failed={total_failed}")
    log.info(f"  skipped={total_skipped}")
    log.info(f"  total_files_moved={total_files}")
    log.info(f"  total_sessions={total_sessions}")
    log.info(f"  duration_s={batch_duration}")
    for page, stats in by_page.items():
        log.info(f"  page={page} {stats}")

    print(f"\n{'=' * 70}")
    print(f"📊 RESUMO DO LOTE LIVE")
    print(f"{'=' * 70}")
    print(f"  Total:        {len(templates)}")
    print(f"  Sucesso:      {total_success}")
    print(f"  Falha:        {total_failed}")
    print(f"  Pulados:      {total_skipped}")
    print(f"  Arq. movidos: {total_files}")
    print(f"  Sessões:      {total_sessions} (1 por página)")
    print(f"  Duração:      {batch_duration}s")

    return {
        "results": all_results,
        "total": len(templates),
        "success": total_success,
        "failed": total_failed,
        "skipped": total_skipped,
        "total_files_moved": total_files,
        "total_sessions": total_sessions,
        "duration_s": batch_duration,
        "by_page": dict(by_page),
    }
