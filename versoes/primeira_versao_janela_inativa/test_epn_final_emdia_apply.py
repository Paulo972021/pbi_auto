#!/usr/bin/env python3
"""Teste isolado do slicer epn_final para caso EMDIA.

Objetivo: detectar seleção acidental na ativação, validar clear e apply com re-read em múltiplos timings.
"""
# ============================================================
# 🧪 TESTE DIAGNÓSTICO: epn_final = EMDIA
# ============================================================

import asyncio
import json
import os
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

REPORT_URL = os.environ.get("PBI_REPORT_URL", "")
TAB_NAME = os.environ.get("PBI_TAB_NAME", "COMPARATIVO")
SLICER_NAME = "epn_final"
TARGET_VALUE = "EMDIA"
LOG_DIR = "./logs"
WAIT_SHORT = 0.4
WAIT_STABILIZE = 1.0
HEADLESS = False

# ============================================================
# ⚙️ LOGGING E ESTADO DE LEITURA DO SLICER
# ============================================================

class TLogger:
    def __init__(self, path: str):
        self.path = path
        self.lines: List[str] = []

    @staticmethod
    def ts() -> str:
        return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")

    def log(self, phase: str, msg: str, level: str = "INFO"):
        line = f"{self.ts()} [{phase}] [{level}] {msg}"
        print(line)
        self.lines.append(line)

    def dump(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            f.write("\n".join(self.lines) + "\n")


@dataclass
class SlicerState:
    found: bool
    title: str
    selected: List[str]
    available: List[str]
    is_open: bool
    bbox: Dict[str, Any]


def _norm(s: Optional[str]) -> str:
    return (s or "").strip().lower()

# ============================================================
# 🔎 OPERAÇÕES DE UI (ABA/SLICER/CLEAR/APPLY)
# ============================================================

async def _activate_tab(tab, tab_name: str) -> bool:
    return bool(await tab.evaluate(
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


async def _slicer_state(tab, slicer_name: str) -> SlicerState:
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
          if (!target) return {found:false, title:'', selected:[], available:[], is_open:false, bbox:{}};

          const r = target.getBoundingClientRect();
          const titleEl = target.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
          const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();

          const items = Array.from(target.querySelectorAll('[role="option"], .slicerItemContainer, .slicer-item, .item, li, .row'));
          const available = items.map(el => (el.textContent || '').trim()).filter(Boolean);

          const selected = items
            .filter(el =>
              el.getAttribute('aria-selected') === 'true' ||
              el.classList.contains('selected') ||
              el.classList.contains('isSelected') ||
              !!el.querySelector('input[type="checkbox"]:checked')
            )
            .map(el => (el.textContent || '').trim())
            .filter(Boolean);

          const popupLike = !!target.querySelector('[role="listbox"], [role="menu"], .slicer-dropdown-popup, .dropdown-list');
          const isOpen = popupLike || target.matches(':focus-within');

          return {
            found: true,
            title,
            selected,
            available,
            is_open: isOpen,
            bbox: {x:r.x, y:r.y, width:r.width, height:r.height}
          };
        }
        """,
        slicer_name,
    )
    if not isinstance(payload, dict):
        return SlicerState(False, "", [], [], False, {})
    return SlicerState(
        found=bool(payload.get("found")),
        title=str(payload.get("title") or ""),
        selected=list(payload.get("selected") or []),
        available=list(payload.get("available") or []),
        is_open=bool(payload.get("is_open")),
        bbox=dict(payload.get("bbox") or {}),
    )


async def _activate_slicer_safe(tab, slicer_name: str) -> Dict[str, Any]:
    return await tab.evaluate(
        """
        (slicerName) => {
          const norm = (s) => (s || '').trim().toLowerCase();
          const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
          const target = visuals.find(v => {
            const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
            const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
            return norm(title).includes(norm(slicerName));
          });
          if (!target) return {ok:false, method:'none', x:null, y:null};

          const header = target.querySelector('.visualTitle, .slicer-header, .header, [title]');
          const clickEl = header || target;
          const r = clickEl.getBoundingClientRect();
          const x = Math.floor(r.left + Math.max(8, Math.min(r.width - 8, 16)));
          const y = Math.floor(r.top + Math.max(8, Math.min(r.height - 8, 16)));
          clickEl.dispatchEvent(new MouseEvent('click', {bubbles:true, cancelable:true, clientX:x, clientY:y}));

          target.dataset._emdia_old_outline = target.style.outline || '';
          target.style.outline = '3px solid #FF4444';
          target.style.outlineOffset = '2px';

          return {ok:true, method: header ? 'header_safe_click' : 'container_safe_click', x, y};
        }
        """,
        slicer_name,
    )


async def _clear_by_button(tab, slicer_name: str) -> Dict[str, Any]:
    return await tab.evaluate(
        """
        (slicerName) => {
          const norm = (s) => (s || '').trim().toLowerCase();
          const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
          const target = visuals.find(v => {
            const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
            const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
            return norm(title).includes(norm(slicerName));
          });
          if (!target) return {ok:false, clicked:false, method:'clear_button'};

          const btn = target.querySelector('[aria-label*="Limpar" i], [title*="Limpar" i], [aria-label*="Clear" i], [title*="Clear" i]');
          if (!btn) return {ok:false, clicked:false, method:'clear_button'};
          btn.click();
          return {ok:true, clicked:true, method:'clear_button'};
        }
        """,
        slicer_name,
    )


async def _clear_by_click_selected_row(tab, slicer_name: str, selected_value: str) -> Dict[str, Any]:
    return await tab.evaluate(
        """
        (slicerName, selectedValue) => {
          const norm = (s) => (s || '').trim().toLowerCase();
          const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
          const target = visuals.find(v => {
            const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
            const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
            return norm(title).includes(norm(slicerName));
          });
          if (!target) return {ok:false, clicked:false, method:'row_click_clear'};

          const items = Array.from(target.querySelectorAll('[role="option"], .slicerItemContainer, .slicer-item, .item, li, .row'));
          const row = items.find(el => norm(el.textContent).includes(norm(selectedValue)));
          if (!row) return {ok:false, clicked:false, method:'row_click_clear'};
          row.click();
          return {ok:true, clicked:true, method:'row_click_clear'};
        }
        """,
        slicer_name,
        selected_value,
    )


async def _click_target_row(tab, slicer_name: str, target_value: str) -> Dict[str, Any]:
    return await tab.evaluate(
        """
        (slicerName, targetValue) => {
          const norm = (s) => (s || '').trim().toLowerCase();
          const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
          const target = visuals.find(v => {
            const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
            const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
            return norm(title).includes(norm(slicerName));
          });
          if (!target) return {found:false, clicked:false, method:'row_container_click'};

          const items = Array.from(target.querySelectorAll('[role="option"], .slicerItemContainer, .slicer-item, .item, li, .row'));
          const row = items.find(el => norm(el.textContent).includes(norm(targetValue)));
          if (!row) return {found:false, clicked:false, method:'row_container_click'};

          const rr = row.getBoundingClientRect();
          row.dispatchEvent(new MouseEvent('click', {
            bubbles:true,
            cancelable:true,
            clientX: rr.left + rr.width / 2,
            clientY: rr.top + rr.height / 2,
          }));

          return {
            found:true,
            clicked:true,
            method:'row_container_click',
            bbox:{x:rr.x, y:rr.y, width:rr.width, height:rr.height},
            text:(row.textContent || '').trim(),
          };
        }
        """,
        slicer_name,
        target_value,
    )


async def run_test() -> int:
    try:
        import pbi_nico as pbi
    except Exception as e:
        print(f"Falha ao importar pbi_nico: {e}")
        return 2

    report_url = REPORT_URL or getattr(pbi, "POWERBI_URL", "")
    browser_path = pbi.normalize_browser_path(getattr(pbi, "BROWSER_PATH", ""))
    if not report_url or not browser_path:
        print("REPORT_URL/BROWSER_PATH não configurados.")
        return 2

    run_ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.abspath(os.path.join(LOG_DIR, f"test_epn_final_emdia_apply_{run_ts}.log"))
    json_path = os.path.abspath(os.path.join(LOG_DIR, f"test_epn_final_emdia_apply_{run_ts}.json"))

    L = TLogger(log_path)
    summary: Dict[str, Any] = {
        "timestamp_start": TLogger.ts(),
        "timestamp_end": None,
        "report_url": report_url,
        "tab_name": TAB_NAME,
        "slicer_name": SLICER_NAME,
        "target_value": TARGET_VALUE,
        "pre_interaction": {},
        "activation_immediate": {},
        "activation_stabilized": {},
        "clear_result": {},
        "target_click_immediate": {},
        "target_click_stabilized": {},
        "target_click_confirmed": {},
        "final_validate": {},
        "status": "FAIL",
        "failure_reason": None,
        "log_file": log_path,
        "json_file": json_path,
    }

    browser = None
    tab = None
    t0 = time.time()

    try:
        # Etapa 1
        L.log("ETAPA 1", "abrindo navegador")
        browser, _ = await pbi.start_isolated_browser(browser_path)
        tab = await pbi.open_report_tab(browser, report_url, set())

        loaded = await pbi.wait_for_visuals_or_abort(tab, stage_label="emdia_test_load", retries=10, wait_seconds=3)
        if not loaded:
            raise RuntimeError("timeout_no_visuals")
        tab_ok = await _activate_tab(tab, TAB_NAME)
        await asyncio.sleep(WAIT_STABILIZE)
        L.log("ETAPA 1", f"aba_ativada={tab_ok}")

        # Etapa 2 — PRE_INTERACTION_STATE
        pre = await _slicer_state(tab, SLICER_NAME)
        summary["pre_interaction"] = {
            "selected_before": pre.selected,
            "available_values": pre.available,
            "is_open": pre.is_open,
            "bbox": pre.bbox,
            "found": pre.found,
            "title": pre.title,
        }
        L.log("PRE_INTERACTION_STATE", f"selected_before={pre.selected}")
        L.log("PRE_INTERACTION_STATE", f"available_values={pre.available}")
        L.log("PRE_INTERACTION_STATE", f"is_open={pre.is_open} bbox={pre.bbox}")

        # Etapa 3 — ativação controlada
        act_meta = await _activate_slicer_safe(tab, SLICER_NAME)
        im = await _slicer_state(tab, SLICER_NAME)
        await asyncio.sleep(WAIT_STABILIZE)
        st = await _slicer_state(tab, SLICER_NAME)

        acc_sel = [v for v in st.selected if v not in pre.selected]
        summary["activation_immediate"] = {
            "method": act_meta.get("method"),
            "coords": {"x": act_meta.get("x"), "y": act_meta.get("y")},
            "selected": im.selected,
            "is_open": im.is_open,
        }
        summary["activation_stabilized"] = {
            "selected": st.selected,
            "is_open": st.is_open,
            "accidental_selection_detected": len(acc_sel) > 0,
            "accidental_values": acc_sel,
        }
        L.log("ACTIVATION_IMMEDIATE_STATE", f"method={act_meta.get('method')} coords=({act_meta.get('x')},{act_meta.get('y')}) selected={im.selected}")
        L.log("ACTIVATION_STABILIZED_STATE", f"selected={st.selected} accidental_selection_detected={len(acc_sel) > 0} accidental_values={acc_sel}")

        # Etapa 4 — clear controlado
        clear_a = {"used": False, "ok": False}
        clear_b = {"used": False, "ok": False}
        clear_final = await _slicer_state(tab, SLICER_NAME)

        if clear_final.selected:
            clear_a_raw = await _clear_by_button(tab, SLICER_NAME)
            clear_a = {"used": True, **(clear_a_raw if isinstance(clear_a_raw, dict) else {})}
            await asyncio.sleep(WAIT_STABILIZE)
            clear_after_a = await _slicer_state(tab, SLICER_NAME)
            clear_a["selected_after"] = clear_after_a.selected
            L.log("CLEAR_ATTEMPT_A", f"payload={clear_a}")

            if clear_after_a.selected:
                sel0 = clear_after_a.selected[0]
                clear_b_raw = await _clear_by_click_selected_row(tab, SLICER_NAME, sel0)
                clear_b = {"used": True, **(clear_b_raw if isinstance(clear_b_raw, dict) else {})}
                await asyncio.sleep(WAIT_STABILIZE)
                clear_after_b = await _slicer_state(tab, SLICER_NAME)
                clear_b["selected_after"] = clear_after_b.selected
                clear_final = clear_after_b
                L.log("CLEAR_ATTEMPT_B", f"payload={clear_b}")
            else:
                clear_final = clear_after_a
        
        summary["clear_result"] = {
            "selected_after_clear": clear_final.selected,
            "clear_ok": len(clear_final.selected) == 0,
            "method_used": "A" if clear_a.get("used") and not clear_b.get("used") else ("B" if clear_b.get("used") else "none"),
            "clear_attempt_a": clear_a,
            "clear_attempt_b": clear_b,
        }
        L.log("CLEAR_RESULT", f"selected_after_clear={clear_final.selected} clear_ok={len(clear_final.selected) == 0}")

        # Etapa 5 — apply EMDIA
        target_loc = await _click_target_row(tab, SLICER_NAME, TARGET_VALUE)
        if not isinstance(target_loc, dict):
            target_loc = {"found": False, "clicked": False, "method": "row_container_click"}
        L.log("TARGET_LOCATOR", f"payload={target_loc}")

        imm = await _slicer_state(tab, SLICER_NAME)
        await asyncio.sleep(WAIT_SHORT)
        stz = await _slicer_state(tab, SLICER_NAME)
        await asyncio.sleep(WAIT_STABILIZE)
        cnf = await _slicer_state(tab, SLICER_NAME)

        summary["target_click_immediate"] = {"selected": imm.selected, "emdia_selected": any(_norm(v) == _norm(TARGET_VALUE) for v in imm.selected)}
        summary["target_click_stabilized"] = {"selected": stz.selected, "emdia_selected": any(_norm(v) == _norm(TARGET_VALUE) for v in stz.selected)}
        summary["target_click_confirmed"] = {"selected": cnf.selected, "emdia_selected": any(_norm(v) == _norm(TARGET_VALUE) for v in cnf.selected)}

        L.log("TARGET_CLICK_RESULT_IMMEDIATE", f"selected={imm.selected} emdia_selected={summary['target_click_immediate']['emdia_selected']} wait={WAIT_SHORT}")
        L.log("TARGET_CLICK_RESULT_STABILIZED", f"selected={stz.selected} emdia_selected={summary['target_click_stabilized']['emdia_selected']} wait={WAIT_STABILIZE}")
        L.log("TARGET_CLICK_RESULT_CONFIRMED", f"selected={cnf.selected} emdia_selected={summary['target_click_confirmed']['emdia_selected']}")

        # Etapa 6 — validação final
        # fecha/reabre via clique seguro
        await _activate_slicer_safe(tab, SLICER_NAME)
        await asyncio.sleep(WAIT_SHORT)
        await _activate_slicer_safe(tab, SLICER_NAME)
        await asyncio.sleep(WAIT_STABILIZE)
        final = await _slicer_state(tab, SLICER_NAME)

        summary["final_validate"] = {
            "initial_selected": pre.selected,
            "after_activation_selected": st.selected,
            "after_clear_selected": clear_final.selected,
            "after_apply_selected": cnf.selected,
            "final_selected": final.selected,
            "emdia_final": any(_norm(v) == _norm(TARGET_VALUE) for v in final.selected),
            "ui_vs_readback_match": any(_norm(v) == _norm(TARGET_VALUE) for v in cnf.selected) == any(_norm(v) == _norm(TARGET_VALUE) for v in final.selected),
        }
        L.log("FINAL_VALIDATE", f"payload={summary['final_validate']}")

        summary["status"] = "SUCCESS"
        summary["failure_reason"] = "none"

    except Exception as e:
        summary["status"] = "FAIL"
        summary["failure_reason"] = str(e)
        L.log("ERROR", f"exception={e}", level="ERROR")
    finally:
        summary["timestamp_end"] = TLogger.ts()
        summary["duration_seconds"] = round(time.time() - t0, 3)

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)

        L.log("OUTPUT", f"log_path={log_path}")
        L.log("OUTPUT", f"json_path={json_path}")

        if tab is not None:
            try:
                await tab.evaluate(
                    """
                    (slicerName) => {
                      const norm = (s) => (s || '').trim().toLowerCase();
                      const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
                      const target = visuals.find(v => {
                        const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
                        const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
                        return norm(title).includes(norm(slicerName));
                      });
                      if (target) {
                        target.style.outline = target.dataset._emdia_old_outline || '';
                      }
                    }
                    """,
                    SLICER_NAME,
                )
            except Exception:
                pass
        if tab is not None or browser is not None:
            try:
                await pbi.graceful_browser_shutdown(tab, browser)
            except Exception:
                pass

        L.dump()

    return 0 if summary["status"] == "SUCCESS" else 1


def main() -> int:
    return asyncio.run(run_test())


if __name__ == "__main__":
    sys.exit(main())
