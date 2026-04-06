#!/usr/bin/env python3
"""Teste isolado da hipótese de perda de estado lógico após scroll em slicer.

Cenário A: alvo no topo (primeiros itens visíveis)
Cenário B: alvo que exige scroll
"""
# ============================================================
# 🧪 TESTE DIAGNÓSTICO: HIPÓTESE DE REGISTRY VS JANELA VISÍVEL
# ============================================================

import asyncio
import json
import os
import sys
import time
import traceback
from datetime import datetime, timezone
from typing import Any, Dict, List

REPORT_URL = os.environ.get("PBI_REPORT_URL", "")
TAB_NAME = os.environ.get("PBI_TAB_NAME", "COMPARATIVO")
SLICER_NAME = os.environ.get("PBI_SLICER_NAME", "acao_tatica")
SCENARIO_A_TARGET = os.environ.get("SCENARIO_A_TARGET", "NAO")
SCENARIO_B_TARGET = os.environ.get("SCENARIO_B_TARGET", "SIM")
SCROLL_DELTA = int(os.environ.get("SCROLL_DELTA", "50"))
MAX_STEPS = int(os.environ.get("MAX_STEPS", "100"))
STALL_THRESHOLD = int(os.environ.get("STALL_THRESHOLD", "5"))
MAX_DURATION_SEC = int(os.environ.get("MAX_DURATION_SEC", "120"))
LOG_DIR = "./logs"

# ============================================================
# ⚙️ LOGGING E INSTRUMENTAÇÃO DE CHAMADAS
# ============================================================

class L:
    def __init__(self, path: str):
        self.path = path
        self.lines: List[str] = []

    @staticmethod
    def t() -> str:
        return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")

    def info(self, msg: str):
        line = f"{self.t()} {msg}"
        print(line)
        self.lines.append(line)

    def flush(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            f.write("\n".join(self.lines) + "\n")


def _type_map(payload: Any) -> Any:
    if isinstance(payload, dict):
        return {k: type(v).__name__ for k, v in payload.items()}
    if isinstance(payload, list):
        return [type(v).__name__ for v in payload]
    return type(payload).__name__


async def _evaluate_logged(tab, log: L, fn_name: str, script: str, payload: Any) -> Any:
    log.info(f"[CALL] fn={fn_name}")
    log.info(f"[CALL] payload={payload}")
    log.info(f"[CALL] payload_types={_type_map(payload)}")
    log.info(f"[CDP] method=Runtime.callFunctionOn params={payload}")
    try:
        return await tab.evaluate(script, payload)
    except Exception as e:
        log.info(f"[EXCEPTION] fn={fn_name} error={e!r}")
        log.info(f"[EXCEPTION] traceback={traceback.format_exc().strip()}")
        raise


async def _activate_tab(tab, log: L, tab_name: str) -> bool:
    return bool(await _evaluate_logged(
        tab,
        log,
        "_activate_tab",
        """
        (name) => {
          const norm = (s) => (s || '').trim().toLowerCase();
          const tabs = Array.from(document.querySelectorAll('[role="tab"], .tab, .pivotItem, .page-tab'));
          const t = tabs.find(el => norm(el.textContent).includes(norm(name)));
          if (!t) return false;
          t.click();
          return true;
        }
        """,
        tab_name,
    ))


async def _read_window(tab, slicer_name: str) -> Dict[str, Any]:
    payload = await tab.evaluate(
        """
        (slicerName) => {
          const norm = (s) => (s || '').trim().toLowerCase();
          const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
          const target = visuals.find(v => {
            const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
            const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
            return norm(title).includes(norm(slicerName));
          });
          if (!target) return {ok:false, values:[], scrollTop:0, maxScroll:0};

          const items = Array.from(target.querySelectorAll('[role="option"], .slicerItemContainer, .slicer-item, .item, li, .row'));
          const values = items.map(el => (el.textContent || '').trim()).filter(Boolean);

          const candidates = Array.from(target.querySelectorAll('*')).filter(el => el.scrollHeight > el.clientHeight);
          const chosen = candidates.sort((a, b) => (b.scrollHeight - b.clientHeight) - (a.scrollHeight - a.clientHeight))[0] || null;

          return {
            ok:true,
            values,
            scrollTop: chosen ? chosen.scrollTop : 0,
            maxScroll: chosen ? Math.max(0, chosen.scrollHeight - chosen.clientHeight) : 0,
          };
        }
        """,
        slicer_name,
    )
    return payload if isinstance(payload, dict) else {"ok": False, "values": [], "scrollTop": 0, "maxScroll": 0}


async def _scroll_once(tab, slicer_name: str, delta: int) -> bool:
    return bool(await tab.evaluate(
        """
        (slicerName, delta) => {
          const norm = (s) => (s || '').trim().toLowerCase();
          const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
          const target = visuals.find(v => {
            const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
            const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
            return norm(title).includes(norm(slicerName));
          });
          if (!target) return false;
          const candidates = Array.from(target.querySelectorAll('*')).filter(el => el.scrollHeight > el.clientHeight);
          const chosen = candidates.sort((a, b) => (b.scrollHeight - b.clientHeight) - (a.scrollHeight - a.clientHeight))[0] || null;
          if (!chosen) return false;
          chosen.scrollTop = chosen.scrollTop + delta;
          return true;
        }
        """,
        slicer_name,
        delta,
    ))


def _norm(v: str) -> str:
    return (v or "").strip().lower()

# ============================================================
# 🔁 CENÁRIOS DE EXECUÇÃO E CONTROLE DO TESTE
# ============================================================

async def _run_scenario(tab, log: L, scenario_name: str, target: str) -> Dict[str, Any]:
    found_registry = set()
    first_seen_step = None
    first_seen_scroll_top = None
    first_seen_text = None
    target_registry_lost = False
    target_seen_previously = False
    target_confirmation_mode = None

    no_new = 0
    t0 = time.time()
    step_rows = []

    while len(step_rows) < MAX_STEPS:
        if time.time() - t0 > MAX_DURATION_SEC:
            stop = "STOP_TIMEOUT"
            break

        step = len(step_rows) + 1
        snapshot = await _read_window(tab, SLICER_NAME)
        if not snapshot.get("ok"):
            stop = "STOP_SLICER_NOT_FOUND"
            break

        values = snapshot.get("values") or []
        scroll_top = int(snapshot.get("scrollTop") or 0)
        max_scroll = int(snapshot.get("maxScroll") or 0)

        found_registry_before = sorted(found_registry)
        target_seen_now = any(_norm(v) == _norm(target) for v in values)

        # atualiza registro lógico persistente
        for v in values:
            found_registry.add(v)

        if target_seen_now and first_seen_step is None:
            first_seen_step = step
            first_seen_scroll_top = scroll_top
            first_seen_text = target
            target_confirmation_mode = "visual"

        target_seen_previously = any(_norm(v) == _norm(target) for v in found_registry)

        found_registry_after = sorted(found_registry)

        if target_seen_previously and not any(_norm(v) == _norm(target) for v in found_registry_after):
            target_registry_lost = True

        # confirmação do alvo por memória
        if target_seen_previously and not target_seen_now:
            target_confirmation_mode = "memory"

        row = {
            "step": step,
            "scrollTop": scroll_top,
            "values_now": values,
            "found_registry_before_step": found_registry_before,
            "found_registry_after_step": found_registry_after,
            "target_seen_now": target_seen_now,
            "target_seen_previously": target_seen_previously,
            "target_confirmation_mode": target_confirmation_mode,
            "target_registry_lost": target_registry_lost,
        }
        step_rows.append(row)

        log.info(f"[{scenario_name}] step={step} first_seen_step={first_seen_step} first_seen_scrollTop={first_seen_scroll_top} first_seen_text={first_seen_text}")
        log.info(f"[{scenario_name}] found_registry_before_step={found_registry_before}")
        log.info(f"[{scenario_name}] found_registry_after_step={found_registry_after}")
        log.info(f"[{scenario_name}] target_seen_now={target_seen_now} target_seen_previously={target_seen_previously}")
        log.info(f"[{scenario_name}] target_confirmation_mode={target_confirmation_mode} target_registry_lost={target_registry_lost}")

        prev_size = len(found_registry_before)
        curr_size = len(found_registry_after)
        if curr_size == prev_size:
            no_new += 1
        else:
            no_new = 0

        if scroll_top >= max_scroll and no_new >= STALL_THRESHOLD:
            stop = "STOP_SCROLL_LIMIT_AND_NO_NEW"
            break

        ok_scroll = await _scroll_once(tab, SLICER_NAME, SCROLL_DELTA)
        await asyncio.sleep(0.35)
        if not ok_scroll:
            stop = "STOP_SCROLL_FAILED"
            break

    else:
        stop = "STOP_MAX_STEPS"

    # validação final por readback
    final_snapshot = await _read_window(tab, SLICER_NAME)
    final_values = final_snapshot.get("values") or []
    target_in_final_readback = any(_norm(v) == _norm(target) for v in final_values)
    if target_in_final_readback:
        target_confirmation_mode = "final_readback"

    return {
        "scenario": scenario_name,
        "target": target,
        "first_seen_step": first_seen_step,
        "first_seen_scrollTop": first_seen_scroll_top,
        "first_seen_text": first_seen_text,
        "target_seen_previously": target_seen_previously,
        "target_registry_lost": target_registry_lost,
        "target_confirmation_mode": target_confirmation_mode,
        "target_in_final_readback": target_in_final_readback,
        "found_registry_final": sorted(found_registry),
        "steps": step_rows,
        "stop_condition": stop,
    }


async def run() -> int:
    try:
        import pbi_nico as pbi
    except Exception as e:
        print(f"Falha ao importar pbi_nico: {e}")
        return 2

    report_url = REPORT_URL or getattr(pbi, "POWERBI_URL", "")
    browser_path = pbi.normalize_browser_path(getattr(pbi, "BROWSER_PATH", ""))
    if not report_url or not browser_path:
        print("REPORT_URL/BROWSER_PATH ausentes")
        return 2

    ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.abspath(os.path.join(LOG_DIR, f"test_slicer_registry_hypothesis_{ts}.log"))
    json_path = os.path.abspath(os.path.join(LOG_DIR, f"test_slicer_registry_hypothesis_{ts}.json"))
    log = L(log_path)

    summary = {
        "timestamp_start": L.t(),
        "timestamp_end": None,
        "report_url": report_url,
        "tab_name": TAB_NAME,
        "slicer_name": SLICER_NAME,
        "scenario_a_target": SCENARIO_A_TARGET,
        "scenario_b_target": SCENARIO_B_TARGET,
        "scenario_a": None,
        "scenario_b": None,
        "log_file": log_path,
        "json_file": json_path,
    }

    browser = None
    tab = None
    try:
        log.info("[FASE] start")
        browser, _ = await pbi.start_isolated_browser(browser_path)
        tab = await pbi.open_report_tab(browser, report_url, set())
        loaded = await pbi.wait_for_visuals_or_abort(tab, stage_label="registry_hypothesis", retries=10, wait_seconds=3)
        if not loaded:
            raise RuntimeError("report_not_loaded")

        tab_ok = await _activate_tab(tab, log, TAB_NAME)
        log.info(f"[FASE] tab_activated={tab_ok}")
        await asyncio.sleep(1)

        initial = await _read_window(tab, SLICER_NAME)
        log.info(f"[INITIAL] slicer_name={SLICER_NAME}")
        log.info(f"[INITIAL] snapshot_ok={initial.get('ok')} scrollTop={initial.get('scrollTop')} maxScroll={initial.get('maxScroll')}")
        log.info(f"[INITIAL] values_now={initial.get('values')}")
        summary["scenario_a"] = {
            "status": "SKIPPED_TEMPORARILY_FOR_STRUCTURAL_FIX",
            "reason": "stopped after initial slicer state capture",
            "initial_snapshot": initial,
        }
        summary["scenario_b"] = {
            "status": "SKIPPED_TEMPORARILY_FOR_STRUCTURAL_FIX",
            "reason": "stopped after initial slicer state capture",
        }

    except Exception as e:
        log.info(f"[ERROR] {e!r}")
        log.info(f"[ERROR] traceback={traceback.format_exc().strip()}")
    finally:
        summary["timestamp_end"] = L.t()
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        log.info(f"[OUTPUT] log_file={log_path}")
        log.info(f"[OUTPUT] json_file={json_path}")
        if tab is not None or browser is not None:
            try:
                await pbi.graceful_browser_shutdown(tab, browser)
            except Exception:
                pass
        log.flush()

    return 0


def main() -> int:
    return asyncio.run(run())


if __name__ == "__main__":
    sys.exit(main())
