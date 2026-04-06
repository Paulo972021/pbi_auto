#!/usr/bin/env python3
"""
Teste isolado para enumeração por micro-scroll do slicer `acao_tatica`.
Não altera fluxo principal do projeto.
"""

import asyncio
import json
import os
import sys
import time
import uuid
from datetime import datetime, timezone
from typing import Any, Dict, List, Tuple

# =============================
# Parâmetros de configuração
# =============================
REPORT_URL = os.environ.get("PBI_REPORT_URL", "")
TAB_NAME = "COMPARATIVO"
SLICER_NAME = "acao_tatica"
SCROLL_DELTA = 50
SCROLL_METHOD = "wheel"  # wheel | scrollTop | key
MAX_STEPS = 100
MAX_DURATION_SEC = 120
STALL_THRESHOLD = 5
BLANK_FILTER = True
LOG_DIR = "./logs"
HEADLESS = False
BROWSER_SLOW_MO = 200


class PhaseLogger:
    def __init__(self, log_path: str):
        self.log_path = log_path
        self._lines: List[str] = []

    @staticmethod
    def _hhmmss_mmm() -> str:
        return datetime.now().strftime("%H:%M:%S.%f")[:-3]

    @staticmethod
    def iso_now() -> str:
        return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")

    def log(self, phase: str, level: str, message: str):
        line = f"[{self._hhmmss_mmm()}] [{phase}] [{level}] {message}"
        print(line)
        self._lines.append(line)

    def flush(self):
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)
        with open(self.log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(self._lines) + "\n")


def normalize_value(text: str) -> str:
    return (text or "").strip()


def should_discard(text: str) -> Tuple[bool, str]:
    v = normalize_value(text)
    if not v:
        return True, "empty"
    if BLANK_FILTER and v.lower() == "(em branco)":
        return True, "blank_filter"
    return False, "accepted"


async def wait_report_loaded(tab, logger: PhaseLogger, timeout_s: int = 45) -> bool:
    t0 = time.time()
    while time.time() - t0 < timeout_s:
        try:
            visual_count = await tab.evaluate("document.querySelectorAll('visual-container').length")
            if int(visual_count or 0) > 0:
                return True
        except Exception:
            pass
        await asyncio.sleep(1)
    return False


async def activate_tab(tab, tab_name: str) -> bool:
    script = """
    (name) => {
      const norm = (s) => (s || '').trim().toLowerCase();
      const targets = Array.from(document.querySelectorAll('[role="tab"], .tab, .pivotItem, .page-tab'));
      const t = targets.find(el => norm(el.textContent).includes(norm(name)));
      if (!t) return false;
      t.click();
      return true;
    }
    """
    return bool(await tab.evaluate(script, tab_name))


async def find_slicer(tab, slicer_name: str) -> Dict[str, Any]:
    script = """
    (slicerName) => {
      const norm = (s) => (s || '').trim().toLowerCase();
      const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
      const candidates = [];
      for (const v of visuals) {
        const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
        const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
        if (norm(title).includes(norm(slicerName))) {
          const r = v.getBoundingClientRect();
          candidates.push({
            title,
            x: r.x, y: r.y, width: r.width, height: r.height,
            selectorHint: 'visual-container'
          });
        }
      }
      if (candidates.length === 0) return {found:false, ambiguous:false, candidate:null, count:0};
      return {
        found: true,
        ambiguous: candidates.length > 1,
        candidate: candidates[0],
        count: candidates.length,
      };
    }
    """
    result = await tab.evaluate(script, slicer_name)
    return result if isinstance(result, dict) else {"found": False, "ambiguous": False, "candidate": None, "count": 0}


async def highlight_slicer(tab, slicer_name: str) -> bool:
    script = """
    (slicerName) => {
      const norm = (s) => (s || '').trim().toLowerCase();
      const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
      const target = visuals.find(v => {
        const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
        const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
        return norm(title).includes(norm(slicerName));
      });
      if (!target) return false;
      target.dataset._ms_test_old_outline = target.style.outline || '';
      target.dataset._ms_test_old_outline_offset = target.style.outlineOffset || '';
      target.style.outline = '3px solid #FF4444';
      target.style.outlineOffset = '2px';
      return true;
    }
    """
    return bool(await tab.evaluate(script, slicer_name))


async def remove_highlight(tab, slicer_name: str):
    script = """
    (slicerName) => {
      const norm = (s) => (s || '').trim().toLowerCase();
      const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
      const target = visuals.find(v => {
        const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
        const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
        return norm(title).includes(norm(slicerName));
      });
      if (!target) return false;
      target.style.outline = target.dataset._ms_test_old_outline || '';
      target.style.outlineOffset = target.dataset._ms_test_old_outline_offset || '';
      return true;
    }
    """
    try:
        await tab.evaluate(script, slicer_name)
    except Exception:
        pass


async def activate_slicer(tab, slicer_name: str) -> Dict[str, Any]:
    script = """
    (slicerName) => {
      const norm = (s) => (s || '').trim().toLowerCase();
      const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
      const target = visuals.find(v => {
        const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
        const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
        return norm(title).includes(norm(slicerName));
      });
      if (!target) return {ok:false, beforeActive:false, afterActive:false};
      const before = target.classList.contains('isFocused') || target.classList.contains('selected') || target.matches(':focus-within');
      target.click();
      const after = target.classList.contains('isFocused') || target.classList.contains('selected') || target.matches(':focus-within');
      return {ok:true, beforeActive:before, afterActive:after};
    }
    """
    result = await tab.evaluate(script, slicer_name)
    if not isinstance(result, dict):
        return {"ok": False, "beforeActive": False, "afterActive": False}
    return result


async def read_visible_values(tab, slicer_name: str) -> Dict[str, Any]:
    script = """
    (slicerName) => {
      const norm = (s) => (s || '').trim().toLowerCase();
      const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
      const target = visuals.find(v => {
        const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
        const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
        return norm(title).includes(norm(slicerName));
      });
      if (!target) return {ok:false, values:[], nodeIds:[]};
      const items = Array.from(target.querySelectorAll('[role="option"], .slicerItemContainer, .slicer-item, .item, li, .row'));
      const values = items
        .map(el => (el.textContent || '').trim())
        .filter(v => v.length > 0);
      const nodeIds = items.map(el => {
        if (!el.dataset.msNodeId) {
          el.dataset.msNodeId = `${Date.now()}_${Math.random().toString(16).slice(2)}`;
        }
        return el.dataset.msNodeId;
      });
      return {ok:true, values, nodeIds};
    }
    """
    payload = await tab.evaluate(script, slicer_name)
    if not isinstance(payload, dict):
        return {"ok": False, "values": [], "nodeIds": []}
    return payload


async def find_scroll_container(tab, slicer_name: str) -> Dict[str, Any]:
    script = """
    (slicerName) => {
      const norm = (s) => (s || '').trim().toLowerCase();
      const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
      const target = visuals.find(v => {
        const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
        const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
        return norm(title).includes(norm(slicerName));
      });
      if (!target) return {ok:false, candidates:[], chosen:null};

      const nodes = Array.from(target.querySelectorAll('*'));
      const candidates = nodes
        .filter(el => el.scrollHeight > el.clientHeight)
        .map((el, idx) => {
          if (!el.dataset.msScrollId) el.dataset.msScrollId = `ms_scroll_${idx}_${Math.random().toString(16).slice(2)}`;
          return {
            scrollId: el.dataset.msScrollId,
            tag: el.tagName,
            className: el.className,
            scrollHeight: el.scrollHeight,
            clientHeight: el.clientHeight,
            scrollTop: el.scrollTop,
          };
        });

      if (!candidates.length) return {ok:false, candidates:[], chosen:null};
      const chosen = candidates.sort((a, b) => (b.scrollHeight - b.clientHeight) - (a.scrollHeight - a.clientHeight))[0];

      const chosenEl = target.querySelector(`[data-ms-scroll-id="${chosen.scrollId}"]`);
      if (chosenEl) {
        chosenEl.dataset._ms_test_old_bg = chosenEl.style.backgroundColor || '';
        chosenEl.style.backgroundColor = 'rgba(255,200,0,0.15)';
        setTimeout(() => {
          chosenEl.style.backgroundColor = chosenEl.dataset._ms_test_old_bg || '';
        }, 1000);
      }

      return {
        ok:true,
        candidates,
        chosen,
        selectorHint: `[data-ms-scroll-id="${chosen.scrollId}"]`,
      };
    }
    """
    result = await tab.evaluate(script, slicer_name)
    if not isinstance(result, dict):
        return {"ok": False, "candidates": [], "chosen": None, "selectorHint": ""}
    return result


async def get_scroll_top(tab, scroll_id: str, slicer_name: str) -> int:
    script = """
    (slicerName, sid) => {
      const norm = (s) => (s || '').trim().toLowerCase();
      const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
      const target = visuals.find(v => {
        const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
        const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
        return norm(title).includes(norm(slicerName));
      });
      if (!target) return null;
      const el = target.querySelector(`[data-ms-scroll-id="${sid}"]`);
      if (!el) return null;
      return el.scrollTop;
    }
    """
    v = await tab.evaluate(script, slicer_name, scroll_id)
    return int(v or 0)


async def do_micro_scroll(tab, scroll_id: str, slicer_name: str, method: str, delta: int) -> bool:
    script = """
    (slicerName, sid, method, delta) => {
      const norm = (s) => (s || '').trim().toLowerCase();
      const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
      const target = visuals.find(v => {
        const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
        const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
        return norm(title).includes(norm(slicerName));
      });
      if (!target) return false;
      const el = target.querySelector(`[data-ms-scroll-id="${sid}"]`);
      if (!el) return false;

      if (method === 'wheel') {
        const evt = new WheelEvent('wheel', {deltaY: delta, bubbles: true, cancelable: true});
        el.dispatchEvent(evt);
        el.scrollTop = el.scrollTop + delta;
      } else if (method === 'scrollTop') {
        el.scrollTop = el.scrollTop + delta;
      } else {
        el.scrollTop = el.scrollTop + delta;
      }
      console.log(`[MICRO-SCROLL TEST] step=unknown scrollTop=${el.scrollTop}`);
      return true;
    }
    """
    return bool(await tab.evaluate(script, slicer_name, scroll_id, method, delta))


def build_summary_schema(**kwargs) -> Dict[str, Any]:
    schema_keys = [
        "test_id", "timestamp_start", "timestamp_end", "duration_seconds", "slicer_name", "report_url",
        "tab_name", "slicer_found", "slicer_activated", "container_found", "container_selector",
        "passive_values", "enumerated_values", "final_consolidated", "diff_passive_vs_enum",
        "total_steps", "scrollTop_initial", "scrollTop_final", "stop_condition", "stop_detail",
        "enumeration_complete", "log_file", "json_file",
    ]
    return {k: kwargs[k] for k in schema_keys}


async def run() -> int:
    try:
        import pbi_nico as pbi
    except Exception as e:
        print(f"Falha ao importar pbi_nico: {e}")
        return 2

    report_url = REPORT_URL or getattr(pbi, "POWERBI_URL", "")
    start_iso = PhaseLogger.iso_now()
    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
    os.makedirs(LOG_DIR, exist_ok=True)
    log_file = os.path.abspath(os.path.join(LOG_DIR, f"test_slicer_micro_scroll_{run_id}.log"))
    json_file = os.path.abspath(os.path.join(LOG_DIR, f"test_slicer_micro_scroll_{run_id}.json"))
    logger = PhaseLogger(log_file)

    t0 = time.time()
    browser = None
    tab = None
    stop_condition = None
    stop_detail = ""

    summary: Dict[str, Any] = {
        "test_id": str(uuid.uuid4()),
        "timestamp_start": start_iso,
        "timestamp_end": None,
        "duration_seconds": 0.0,
        "slicer_name": SLICER_NAME,
        "report_url": report_url,
        "tab_name": TAB_NAME,
        "slicer_found": False,
        "slicer_activated": False,
        "container_found": False,
        "container_selector": "",
        "passive_values": [],
        "enumerated_values": [],
        "final_consolidated": [],
        "diff_passive_vs_enum": [],
        "total_steps": 0,
        "scrollTop_initial": 0,
        "scrollTop_final": 0,
        "stop_condition": None,
        "stop_detail": "",
        "enumeration_complete": False,
        "log_file": log_file,
        "json_file": json_file,
    }

    logger.log("FASE 0", "INFO", f"timestamp_start={start_iso}")
    logger.log("FASE 0", "INFO", f"params url={report_url} tab={TAB_NAME} slicer={SLICER_NAME}")
    logger.log("FASE 0", "INFO", f"params method={SCROLL_METHOD} delta={SCROLL_DELTA} max_steps={MAX_STEPS}")

    try:
        # FASE 1
        logger.log("FASE 1", "INFO", "abrindo navegador")
        browser_path = pbi.normalize_browser_path(getattr(pbi, "BROWSER_PATH", ""))
        browser, _profile = await pbi.start_isolated_browser(browser_path)
        owned_tab_refs = set()
        tab = await pbi.open_report_tab(browser, report_url, owned_tab_refs)
        loaded = await wait_report_loaded(tab, logger)
        if not loaded:
            logger.log("FASE 1", "ERROR", "relatório não carregou no timeout")
            stop_condition = "STOP_TIMEOUT"
            stop_detail = "report_load_timeout"
            return 1
        logger.log("FASE 1", "INFO", "relatório carregado")

        # FASE 2
        logger.log("FASE 2", "INFO", f"ativando aba={TAB_NAME}")
        tab_ok = await activate_tab(tab, TAB_NAME)
        await asyncio.sleep(2)
        if not tab_ok:
            logger.log("FASE 2", "WARN", "aba não encontrada (seguindo na aba atual)")
        else:
            logger.log("FASE 2", "INFO", "aba ativada")

        # FASE 3
        slicer_info = await find_slicer(tab, SLICER_NAME)
        summary["slicer_found"] = bool(slicer_info.get("found"))
        if not summary["slicer_found"]:
            logger.log("FASE 3", "ERROR", "slicer não encontrado")
            stop_condition = "STOP_ACTIVATION_FAIL"
            stop_detail = "slicer_not_found"
            return 1
        logger.log("FASE 3", "INFO", f"slicer encontrado ambiguous={slicer_info.get('ambiguous')} count={slicer_info.get('count')}")
        logger.log("FASE 3", "INFO", f"slicer_bbox={slicer_info.get('candidate')}")

        # FASE 4
        await highlight_slicer(tab, SLICER_NAME)
        act = await activate_slicer(tab, SLICER_NAME)
        summary["slicer_activated"] = bool(act.get("ok") and (act.get("afterActive") or not act.get("beforeActive")))
        logger.log("FASE 4", "INFO", f"activation before={act.get('beforeActive')} after={act.get('afterActive')} ok={act.get('ok')}")

        # FASE 5
        passive_payload = await read_visible_values(tab, SLICER_NAME)
        passive_values = []
        for raw in passive_payload.get("values", []):
            discarded, reason = should_discard(raw)
            if discarded:
                logger.log("FASE 5", "RESULT", f"raw='{raw}' accepted=False reason={reason}")
            else:
                logger.log("FASE 5", "RESULT", f"raw='{raw}' accepted=True")
                passive_values.append(normalize_value(raw))
        summary["passive_values"] = sorted(list(dict.fromkeys(passive_values)))
        logger.log("FASE 5", "RESULT", f"passive_values={summary['passive_values']}")

        # FASE 6
        container = await find_scroll_container(tab, SLICER_NAME)
        for c in container.get("candidates", []):
            logger.log("FASE 6", "RESULT", f"candidate tag={c['tag']} class={c['className']} scrollHeight={c['scrollHeight']} clientHeight={c['clientHeight']} scrollTop={c['scrollTop']}")
        if not container.get("ok"):
            logger.log("FASE 6", "ERROR", "nenhum container de scroll adequado")
            stop_condition = "STOP_CONTAINER_LOST"
            stop_detail = "scroll_container_not_found"
            return 1
        chosen = container.get("chosen")
        summary["container_found"] = True
        summary["container_selector"] = container.get("selectorHint") or ""
        scroll_id = chosen.get("scrollId")
        logger.log("FASE 6", "INFO", f"container escolhido={chosen} motivo=max_scroll_range")

        # FASE 7
        all_unique_values: List[str] = []
        prev_texts: List[str] = []
        prev_nodes: List[str] = []
        no_new_values_streak = 0
        start_scroll_top = await get_scroll_top(tab, scroll_id, SLICER_NAME)
        summary["scrollTop_initial"] = start_scroll_top

        for step in range(1, MAX_STEPS + 1):
            if (time.time() - t0) > MAX_DURATION_SEC:
                stop_condition = "STOP_TIMEOUT"
                stop_detail = "max_duration_exceeded"
                logger.log("FASE 7", "STOP", f"condition={stop_condition}")
                break

            logger.log("FASE 7", "STEP", f"step={step} iniciando scroll_method={SCROLL_METHOD} delta={SCROLL_DELTA}")

            before_top = await get_scroll_top(tab, scroll_id, SLICER_NAME)
            payload_before = await read_visible_values(tab, SLICER_NAME)
            raw_before = payload_before.get("values", [])
            node_before = payload_before.get("nodeIds", [])

            scrolled = await do_micro_scroll(tab, scroll_id, SLICER_NAME, SCROLL_METHOD, SCROLL_DELTA)
            await asyncio.sleep(max(BROWSER_SLOW_MO / 1000.0, 0.2))
            after_top = await get_scroll_top(tab, scroll_id, SLICER_NAME)
            payload_after = await read_visible_values(tab, SLICER_NAME)
            raw_texts = payload_after.get("values", [])
            node_after = payload_after.get("nodeIds", [])

            discarded_texts = []
            discard_reasons: Dict[str, str] = {}
            accepted_values = []
            for t in raw_texts:
                discarded, reason = should_discard(t)
                if discarded:
                    discarded_texts.append(t)
                    discard_reasons[t] = reason
                else:
                    accepted_values.append(normalize_value(t))

            new_values = [v for v in accepted_values if v not in all_unique_values]
            for v in new_values:
                all_unique_values.append(v)

            scroll_changed = (after_top != before_top)
            text_changed = (raw_texts != prev_texts)
            node_identity_changed = (node_after != prev_nodes)
            selection_side_effect_detected = False
            virtualization_evidence = node_identity_changed
            window_changed = text_changed or node_identity_changed

            if new_values:
                no_new_values_streak = 0
            else:
                no_new_values_streak += 1

            logger.log("FASE 7", "RESULT", f"step={step} scroll_method={SCROLL_METHOD} delta={SCROLL_DELTA}")
            logger.log("FASE 7", "RESULT", f"scrollTop_before={before_top} scrollTop_after={after_top} scroll_changed={scroll_changed}")
            logger.log("FASE 7", "RESULT", f"raw_texts={raw_texts}")
            logger.log("FASE 7", "RESULT", f"discarded_texts={discarded_texts}")
            logger.log("FASE 7", "RESULT", f"discard_reasons={discard_reasons}")
            logger.log("FASE 7", "RESULT", f"accepted_values={accepted_values}")
            logger.log("FASE 7", "RESULT", f"new_values_in_this_step={new_values}")
            logger.log("FASE 7", "RESULT", f"all_unique_values_so_far={all_unique_values}")
            logger.log("FASE 7", "RESULT", f"window_changed={window_changed} text_changed={text_changed} node_identity_changed={node_identity_changed}")
            logger.log("FASE 7", "RESULT", f"selection_side_effect_detected={selection_side_effect_detected}")
            logger.log("FASE 7", "RESULT", f"virtualization_evidence={virtualization_evidence}")

            prev_texts = raw_texts
            prev_nodes = node_after
            summary["total_steps"] = step
            summary["scrollTop_final"] = after_top

            if not scrolled:
                stop_condition = "STOP_CONTAINER_LOST"
                stop_detail = "scroll_action_failed"
                logger.log("FASE 7", "STOP", f"condition={stop_condition}")
                break

            max_scroll_reached = False
            try:
                max_scroll_reached = bool(await tab.evaluate(
                    """
                    (slicerName, sid) => {
                      const norm = (s) => (s || '').trim().toLowerCase();
                      const visuals = Array.from(document.querySelectorAll('visual-container, .visualContainerHost, .visual'));
                      const target = visuals.find(v => {
                        const titleEl = v.querySelector('[title], .visualTitleText, .slicer-header-text, .headerText');
                        const title = (titleEl?.getAttribute('title') || titleEl?.textContent || '').trim();
                        return norm(title).includes(norm(slicerName));
                      });
                      if (!target) return false;
                      const el = target.querySelector(`[data-ms-scroll-id="${sid}"]`);
                      if (!el) return false;
                      return (el.scrollTop + el.clientHeight) >= (el.scrollHeight - 1);
                    }
                    """,
                    SLICER_NAME,
                    scroll_id,
                ))
            except Exception:
                max_scroll_reached = False

            if not scroll_changed and max_scroll_reached:
                stop_condition = "STOP_SCROLL_LIMIT"
                stop_detail = "scrollTop não mudou e limite máximo atingido"
                logger.log("FASE 7", "STOP", f"condition={stop_condition} scrollTop_final={after_top}")
                break
            if no_new_values_streak >= STALL_THRESHOLD:
                stop_condition = "STOP_NO_NEW_VALUES"
                stop_detail = f"{STALL_THRESHOLD} passos sem novos valores"
                logger.log("FASE 7", "STOP", f"condition={stop_condition}")
                break
            if no_new_values_streak >= STALL_THRESHOLD and not window_changed:
                stop_condition = "STOP_WINDOW_FROZEN"
                stop_detail = "janela congelada"
                logger.log("FASE 7", "STOP", f"condition={stop_condition}")
                break
            logger.log("FASE 7", "RESULT", "stop_condition=None")

        if stop_condition is None:
            stop_condition = "STOP_TIMEOUT"
            stop_detail = "loop finalizado sem condição explícita"

        # FASE 8
        summary["enumerated_values"] = list(all_unique_values)
        summary["final_consolidated"] = sorted(list(dict.fromkeys(all_unique_values)))
        summary["diff_passive_vs_enum"] = sorted(list(set(summary["passive_values"]) ^ set(summary["enumerated_values"])))
        summary["stop_condition"] = stop_condition
        summary["stop_detail"] = stop_detail
        summary["enumeration_complete"] = stop_condition in {"STOP_SCROLL_LIMIT", "STOP_NO_NEW_VALUES"}

        logger.log("FASE 8", "RESULT", f"passive_values={summary['passive_values']}")
        logger.log("FASE 8", "RESULT", f"enumerated_values={summary['enumerated_values']}")
        logger.log("FASE 8", "RESULT", f"final_consolidated={summary['final_consolidated']}")
        logger.log("FASE 8", "RESULT", f"diff_passive_vs_enum={summary['diff_passive_vs_enum']}")

    except Exception as e:
        stop_condition = stop_condition or "STOP_TIMEOUT"
        stop_detail = f"exception:{e}"
        summary["stop_condition"] = stop_condition
        summary["stop_detail"] = stop_detail
        logger.log("FASE 10", "ERROR", f"falha geral: {e}")
    finally:
        # FASE 9
        end_iso = PhaseLogger.iso_now()
        summary["timestamp_end"] = end_iso
        summary["duration_seconds"] = round(time.time() - t0, 3)
        summary["log_file"] = log_file
        summary["json_file"] = json_file

        summary_obj = build_summary_schema(**summary)
        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(summary_obj, f, ensure_ascii=False, indent=2)
        logger.log("FASE 9", "INFO", f"log salvo em {log_file}")
        logger.log("FASE 9", "INFO", f"json salvo em {json_file}")

        # FASE 10
        if tab is not None:
            await remove_highlight(tab, SLICER_NAME)
        if tab is not None or browser is not None:
            try:
                await pbi.graceful_browser_shutdown(tab, browser)
            except Exception:
                pass

        logger.log("FASE 10", "INFO", f"timestamp_end={end_iso} duration_seconds={summary['duration_seconds']}")
        if summary.get("slicer_found") and summary.get("container_found"):
            logger.log("FASE 10", "RESULT", f"SUCESSO stop_condition={summary.get('stop_condition')} detail={summary.get('stop_detail')}")
        else:
            logger.log("FASE 10", "RESULT", f"FALHA stop_condition={summary.get('stop_condition')} detail={summary.get('stop_detail')}")

        logger.flush()

    return 0 if summary.get("slicer_found") and summary.get("container_found") else 1


def main() -> int:
    return asyncio.run(run())


if __name__ == "__main__":
    sys.exit(main())
