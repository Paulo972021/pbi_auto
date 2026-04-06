#pbi_auto_v06.py
"""
Automação Power BI - Scan + Exportação Seletiva de Visuais
Usa nodriver para:
  1. Abrir o Power BI e navegar para a aba desejada
  2. Escanear TODOS os visuais que possuem "Mais opções"
  3. Listar ordenados no terminal para você escolher
  4. Exportar os dados dos visuais selecionados

Uso:
    1. Cole seu link do Power BI na variável POWERBI_URL
    2. Cole o caminho do navegador na variável BROWSER_PATH
    3. (Opcional) Altere TARGET_PAGE se quiser outra aba
    4. Execute: python powerbi_export.py
"""

import asyncio
import logging
import sys
import os
import json
import re
import contextlib
import tempfile

try:
    import configpbi
except ModuleNotFoundError:
    print(
        "Arquivo 'configpbi.py' não encontrado. "
        "Crie esse arquivo com as variáveis 'url' e 'browser'."
    )
    sys.exit(1)

try:
    import nodriver as uc
except ImportError:
    print("nodriver não encontrado. Instale com: pip install nodriver")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Configuração de logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("powerbi_export")

# ---------------------------------------------------------------------------
# Constantes de tempo (segundos)
# ---------------------------------------------------------------------------
# Estas variáveis controlam o ritmo do script.
# Ajuste conforme a velocidade da máquina, internet e comportamento do relatório.

# Tempo de espera após abrir inicialmente o Power BI.
# AUMENTE quando o relatório abre em branco ou ainda sem visuais.
# REDUZA se a sua máquina abre o relatório rapidamente e você quer agilizar.
PAGE_LOAD_WAIT = 20

# Espera curta usada entre ações pequenas:
# ex.: fechar menu, pressionar ESC, aguardar animação curta.
# AUMENTE se menus ficarem "presos".
# REDUZA para acelerar o fluxo quando a interface responde instantaneamente.
SHORT_WAIT = 2

# Espera média usada entre etapas intermediárias:
# ex.: após hover, abertura de menu, troca de foco.
# AUMENTE se ações em sequência falharem por "timing".
# REDUZA se o processo estiver estável e você quiser ganhar velocidade.
MEDIUM_WAIT = 5

# Espera longa usada quando o Power BI precisa renderizar algo mais pesado:
# ex.: após navegar para outra aba/página do relatório.
# AUMENTE quando trocar de página não mostra visuais no primeiro ciclo.
# REDUZA caso a navegação entre páginas esteja rápida no seu ambiente.
LONG_WAIT = 10

# Tempo de espera após clicar em "Exportar" para o download começar.
# Se os downloads demorarem a iniciar, aumente.
# Se os arquivos já baixam imediatamente, pode reduzir.
DOWNLOAD_WAIT = 15

# Tempo entre tentativas de reabrir "Mais opções".
# Se o botão estiver aparecendo devagar, aumente.
# Se o botão aparece rápido, reduzir deixa as tentativas mais ágeis.
RETRY_MENU_WAIT = 2

# Tempo para aguardar overlays/popups desaparecerem.
# Se muitos popups ficarem "grudados", aumente um pouco.
# Se quase nunca há popup, pode reduzir para agilizar cada tentativa.
OVERLAY_SETTLE_WAIT = 2

# Tempo máximo para o usuário responder no terminal.
# Se passar disso, o script poderá seguir com a ação padrão.
# AUMENTE para dar mais tempo de escolha manual.
# REDUZA para execução mais automática/rápida.
USER_INPUT_TIMEOUT = 5

# Número de tentativas para abrir o botão "Mais opções".
MORE_OPTIONS_RETRIES = 5

# Número de tentativas para confirmar que um visual suporta exportação.
EXPORT_PROBE_RETRIES = 3


    # ╔══════════════════════════════════════════════════════════════╗
    # ║  COLE O LINK DO POWER BI AQUI EMBAIXO (entre as aspas):    ║
    # ╚══════════════════════════════════════════════════════════════╝
POWERBI_URL = getattr(configpbi, "url", "")
    

    # ╔══════════════════════════════════════════════════════════════╗
    # ║  COLE O CAMINHO DO EXECUTÁVEL DO NAVEGADOR (entre aspas):  ║
    # ║                                                              ║
    # ║  Edge:                                                       ║
    # ║  C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe║
    # ║                                                              ║
    # ║  Chrome:                                                     ║
    # ║  C:\Program Files\Google\Chrome\Application\chrome.exe      ║
    # ╚══════════════════════════════════════════════════════════════╝
BROWSER_PATH = getattr(configpbi, "browser", "")


    # ╔══════════════════════════════════════════════════════════════╗
    # ║  NOME DA ABA/PÁGINA PARA NAVEGAR (entre aspas):            ║
    # ║  Deixe vazio "" para ficar na página inicial                ║
    # ╚══════════════════════════════════════════════════════════════╝
TARGET_PAGE = "COMPARATIVO"


# ---------------------------------------------------------------------------
# FILTER_PLAN — Plano declarativo de filtros a aplicar após enumeração
# ---------------------------------------------------------------------------
# Chaves: título normalizado do slicer (lowercase, sem "(Ainda não aplicado)")
# mode: "single" (exatamente 1 valor) | "multi" (um ou mais valores)
# clear_first: True → limpa seleção antes de aplicar
# target_values: lista de valores exatos a selecionar
# ---------------------------------------------------------------------------
FILTER_PLAN: dict = {
    "entrante": {
        "mode": "single",
        "clear_first": True,
        "target_values": ["1"],
        "required": True,
    },
        "epn_final": {
        "mode": "single",
        "clear_first": True,
        "target_values": ["ALMAVIVA"],
        "required": True,
    },
}

# ---------------------------------------------------------------------------
# Funções auxiliares
# ---------------------------------------------------------------------------

def validate_runtime_config():
    """Valida as configurações vindas do configpbi.py."""
    normalized_url = str(POWERBI_URL or "").strip()
    normalized_browser = str(BROWSER_PATH or "").strip()

    if not normalized_url:
        log.error("❌ configpbi.url está vazio ou ausente.")
        return False

    if not normalized_browser:
        log.error("❌ configpbi.browser está vazio ou ausente.")
        return False

    if not os.path.exists(normalized_browser):
        log.error(f"❌ Caminho do navegador não existe: {normalized_browser}")
        return False

    return True


def normalize_browser_path(path: str) -> str:
    """Normaliza o caminho do executável do navegador."""
    return os.path.abspath(os.path.expanduser(path.strip()))

def build_browser_args():
    """
    Argumentos do navegador.
    Mantidos em função separada para facilitar ajustes futuros.
    """
    base_args = [
        "--no-sandbox",
        "--disable-blink-features=AutomationControlled",
        "--disable-infobars",
        "--start-maximized",
        "--lang=pt-BR",
        "--disable-popup-blocking",
    ]
    background_safe_args = [
        "--disable-background-timer-throttling",
        "--disable-renderer-backgrounding",
        "--disable-backgrounding-occluded-windows",
        "--disable-features=CalculateNativeWinOcclusion",
    ]

    merged_args = list(base_args)
    for flag in background_safe_args:
        if flag not in merged_args:
            merged_args.append(flag)

    kept_flags = [f for f in base_args if f in merged_args]
    added_flags = [f for f in merged_args if f not in base_args]
    log.info("[BACKGROUND_MODE_FLAGS]")
    log.info(f"  flags_added={added_flags}")
    log.info(f"  flags_kept={kept_flags}")
    return merged_args


def _tab_ref(tab):
    """
    Retorna uma referência estável da tab para rastrear apenas abas da automação.
    """
    target = getattr(tab, "target", None)
    target_id = getattr(target, "target_id", None)
    if target_id:
        return str(target_id)
    return str(id(tab))


async def start_isolated_browser(browser_path: str):
    """
    Inicia uma instância isolada do navegador SEM matar processos existentes.

    A ideia aqui é:
    - não encerrar outras janelas/abas já abertas pelo usuário
    - usar um perfil temporário separado para não acoplar em sessão existente
    - controlar somente as abas abertas pela automação
    """
    profile_dir = tempfile.mkdtemp(prefix="pbi_auto_profile_")
    browser = await uc.start(
        headless=False,
        browser_executable_path=browser_path,
        user_data_dir=profile_dir,
        browser_args=build_browser_args(),
    )
    return browser, profile_dir


async def open_report_tab(browser, url: str, owned_tab_refs: set[str]):
    """
    Abre a aba do relatório e devolve a referência dela.
    Essa aba será a única aba que o script deve manipular diretamente.
    """
    tab = await browser.get(url, new_tab=True)
    owned_tab_refs.add(_tab_ref(tab))
    return tab


async def safe_focus_tab(tab):
    """
    Tenta trazer a aba controlada para frente sem interferir nas demais.
    """
    try:
        await tab.activate()
    except Exception:
        with contextlib.suppress(Exception):
            await tab.evaluate("window.focus()")


async def get_page_focus_state(tab, stage_label: str = "") -> dict:
    """Coleta o estado atual de foco/visibilidade da página."""
    script = """
        (() => {
            const state = {
                hidden: null,
                visibilityState: null,
                hasFocus: null,
                webkitHidden: null,
                onblur_bound: false,
                onfocus_bound: false,
            };
            try { state.hidden = document.hidden; } catch (_) {}
            try { state.visibilityState = document.visibilityState; } catch (_) {}
            try { state.hasFocus = document.hasFocus ? document.hasFocus() : null; } catch (_) {}
            try { state.webkitHidden = document.webkitHidden; } catch (_) {}
            try { state.onblur_bound = typeof window.onblur === "function"; } catch (_) {}
            try { state.onfocus_bound = typeof window.onfocus === "function"; } catch (_) {}
            return state;
        })()
    """
    try:
        raw_state = await tab.evaluate(script)
    except Exception as e:
        raw_state = {
            "ok": False,
            "hidden": None,
            "visibilityState": None,
            "hasFocus": None,
            "stage": stage_label or "unknown",
            "error": str(e),
        }

    if not isinstance(raw_state, dict):
        log.warning("[FOCUS_STATE_CONTRACT_ERROR]")
        log.warning(f"  stage={stage_label or 'unknown'}")
        log.warning(f"  returned_type={type(raw_state).__name__}")
        log.warning("  fallback_dict_created=True")
        raw_state = {
            "ok": False,
            "hidden": None,
            "visibilityState": None,
            "hasFocus": None,
            "stage": stage_label or "unknown",
            "error": f"invalid_focus_state_type:{type(raw_state).__name__}",
        }

    state = {
        "ok": bool(
            raw_state.get("hidden") is False
            and str(raw_state.get("visibilityState") or "").lower() == "visible"
            and bool(raw_state.get("hasFocus"))
        ),
        "hidden": raw_state.get("hidden"),
        "visibilityState": raw_state.get("visibilityState"),
        "hasFocus": raw_state.get("hasFocus"),
        "stage": stage_label or "unknown",
        "error": raw_state.get("error"),
    }

    log.info("[PAGE_FOCUS_STATE]")
    log.info(f"  stage={state['stage']}")
    log.info(f"  hidden={state.get('hidden')}")
    log.info(f"  visibilityState={state.get('visibilityState')}")
    log.info(f"  hasFocus={state.get('hasFocus')}")
    log.info(f"  ok={state.get('ok')}")
    return state


async def install_focus_visibility_emulation(tab) -> dict:
    """
    Injeta emulação defensiva/idempotente de foco e visibilidade.
    """
    script = """
        (() => {
            const result = {
                ok: false,
                patched_hidden: false,
                patched_visibility_state: false,
                patched_has_focus: false,
                patched_webkit_hidden: false,
            };
            try {
                const defineSafe = (obj, prop, getterFn) => {
                    try {
                        Object.defineProperty(obj, prop, {
                            get: getterFn,
                            configurable: true,
                        });
                        return true;
                    } catch (_) {
                        return false;
                    }
                };

                const dProto = Object.getPrototypeOf(document);
                result.patched_hidden = defineSafe(dProto, "hidden", () => false);
                result.patched_visibility_state = defineSafe(dProto, "visibilityState", () => "visible");
                if ("webkitHidden" in document) {
                    result.patched_webkit_hidden = defineSafe(dProto, "webkitHidden", () => false);
                } else {
                    result.patched_webkit_hidden = true;
                }
                result.patched_has_focus = defineSafe(dProto, "hasFocus", () => (() => true));

                try { window.dispatchEvent(new Event("focus")); } catch (_) {}
                try { document.dispatchEvent(new Event("visibilitychange")); } catch (_) {}
                try { document.dispatchEvent(new Event("focus")); } catch (_) {}

                result.ok = Boolean(
                    result.patched_hidden &&
                    result.patched_visibility_state &&
                    result.patched_has_focus
                );
                return result;
            } catch (err) {
                result.error = String(err);
                return result;
            }
        })()
    """
    try:
        patched_raw = await tab.evaluate(script)
    except Exception as e:
        patched_raw = {"ok": False, "error": str(e)}

    if not isinstance(patched_raw, dict):
        patched = {
            "ok": False,
            "hidden": None,
            "visibilityState": None,
            "hasFocus": None,
            "stage": "install_focus_visibility_emulation",
            "error": f"invalid_focus_install_type:{type(patched_raw).__name__}",
            "patched_hidden": False,
            "patched_visibility_state": False,
            "patched_has_focus": False,
        }
        log.warning("[FOCUS_STATE_CONTRACT_ERROR]")
        log.warning("  stage=install_focus_visibility_emulation")
        log.warning(f"  returned_type={type(patched_raw).__name__}")
        log.warning("  fallback_dict_created=True")
    else:
        patched = {
            "ok": bool(patched_raw.get("ok", False)),
            "hidden": None,
            "visibilityState": None,
            "hasFocus": None,
            "stage": "install_focus_visibility_emulation",
            "error": patched_raw.get("error"),
            "patched_hidden": bool(patched_raw.get("patched_hidden", False)),
            "patched_visibility_state": bool(patched_raw.get("patched_visibility_state", False)),
            "patched_has_focus": bool(patched_raw.get("patched_has_focus", False)),
        }

    log.info("[FOCUS_EMULATION_INSTALL]")
    log.info(f"  ok={patched.get('ok', False)}")
    log.info(f"  patched_hidden={patched.get('patched_hidden', False)}")
    log.info(f"  patched_visibility_state={patched.get('patched_visibility_state', False)}")
    log.info(f"  patched_has_focus={patched.get('patched_has_focus', False)}")
    return patched


async def ensure_focus_visibility_emulation(tab, stage_label: str) -> bool:
    """
    Verifica estado de foco/visibilidade e reaplica emulação quando necessário.
    """
    state_before = await get_page_focus_state(tab, stage_label=f"{stage_label}:before")
    if not isinstance(state_before, dict):
        log.warning("[FOCUS_STATE_CONTRACT_ERROR]")
        log.warning(f"  stage={stage_label}:before")
        log.warning(f"  returned_type={type(state_before).__name__}")
        log.warning("  fallback_dict_created=True")
        state_before = {
            "ok": False,
            "hidden": None,
            "visibilityState": None,
            "hasFocus": None,
            "stage": f"{stage_label}:before",
            "error": f"invalid_focus_state_type:{type(state_before).__name__}",
        }

    hidden = state_before.get("hidden")
    visibility_state = str(state_before.get("visibilityState") or "").lower()
    has_focus = bool(state_before.get("hasFocus"))
    ok_before = (hidden is False) and (visibility_state == "visible") and has_focus

    if not ok_before:
        await install_focus_visibility_emulation(tab)
        state_after = await get_page_focus_state(tab, stage_label=f"{stage_label}:after_install")
    else:
        state_after = state_before

    hidden_after = state_after.get("hidden")
    visibility_after = str(state_after.get("visibilityState") or "").lower()
    has_focus_after = bool(state_after.get("hasFocus"))
    ok = (hidden_after is False) and (visibility_after == "visible") and has_focus_after

    log.info("[FOCUS_EMULATION_CHECK]")
    log.info(f"  stage={stage_label}")
    log.info(f"  document_hidden={hidden_after}")
    log.info(f"  visibility_state={visibility_after}")
    log.info(f"  has_focus={has_focus_after}")
    log.info(f"  ok={ok}")
    return ok


async def get_tab_url(tab) -> str:
    """Lê a URL atual da aba com tolerância a falhas."""
    try:
        url = await tab.evaluate("window.location.href")
        return str(url or "")
    except Exception:
        return ""


async def close_extra_tabs_created_by_script(browser, keep_tab, owned_tab_refs: set[str]):
    """
    Fecha apenas abas extras abertas PELO PRÓPRIO SCRIPT, se necessário.

    Importante:
    - não tenta fechar janelas do navegador já existentes do usuário
    - só atua nas tabs conhecidas pela instância controlada
    """
    try:
        tabs = list(getattr(browser, "tabs", []) or [])
    except Exception:
        tabs = []

    keep_ref = _tab_ref(keep_tab)
    for tab in tabs:
        tab_ref = _tab_ref(tab)
        if tab_ref == keep_ref:
            continue
        if tab_ref not in owned_tab_refs:
            # Segurança extra: nunca fecha aba que não foi aberta pela automação.
            continue
        try:
            await tab.close()
            owned_tab_refs.discard(tab_ref)
            await asyncio.sleep(0.5)
        except Exception:
            pass

async def js_click(tab, js_expression: str, description: str = "") -> bool:
    """Click via JavaScript — contorna overlays e pointer-events:none."""
    try:
        result = await tab.evaluate(f"""
            (() => {{
                const el = {js_expression};
                if (el) {{
                    el.scrollIntoView({{block: 'center'}});
                    el.dispatchEvent(new MouseEvent('mouseover', {{bubbles: true}}));
                    el.dispatchEvent(new MouseEvent('mouseenter', {{bubbles: true}}));
                    el.click();
                    return true;
                }}
                return false;
            }})()
        """)
        if result:
            log.info(f"  ✅ {description}")
            return True
        return False
    except Exception:
        return False


async def js_click_xpath(tab, xpath: str, description: str = "") -> bool:
    """Click via JS usando XPath."""
    escaped = xpath.replace('"', '\\"')
    return await js_click(
        tab,
        f'document.evaluate("{escaped}", document, null, '
        f'XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue',
        description or xpath,
    )


async def allow_multiple_downloads(tab):
    """
    Autoriza múltiplos downloads via CDP Page.setDownloadBehavior.
    Também tenta aceitar qualquer diálogo de permissão do browser.
    """
    log.info("📥 Configurando permissão para múltiplos downloads...")
    try:
        # CDP: permite downloads automáticos
        await tab.send(uc.cdp.page.set_download_behavior(
            behavior="allow",
            download_path=None,  # Usa pasta Downloads padrão
        ))
        log.info("  ✅ Download behavior configurado via CDP (Page)")
    except Exception:
        try:
            # Alternativa: Browser.setDownloadBehavior
            await tab.send(uc.cdp.browser.set_download_behavior(
                behavior="allow",
                download_path=None,
            ))
            log.info("  ✅ Download behavior configurado via CDP (Browser)")
        except Exception as e:
            log.debug(f"  CDP download behavior não disponível: {e}")

    # Aceita automaticamente o diálogo de permissão de downloads via JS
    # (o Chrome/Edge mostra "Este site quer fazer download de vários arquivos")
    try:
        await tab.evaluate("""
            // Intercepta diálogos de permissão
            if (window.Notification && Notification.permission !== 'granted') {
                Notification.requestPermission();
            }
        """)
    except Exception:
        pass


async def accept_download_permission(tab):
    """
    Aceita o diálogo do browser que pede permissão para múltiplos downloads.
    Esse diálogo aparece como uma barra no topo ou um popup.
    """
    try:
        # Tenta clicar em botões de permissão comuns do browser
        await tab.evaluate("""
            (() => {
                // Chrome/Edge download permission bar
                const permissionSelectors = [
                    // Botão "Permitir" / "Allow" em português e inglês
                    'button[id*="allow"]',
                    'button[id*="permit"]',
                    '#infobar-allow-button',
                    '#download-permission-allow',
                    'button.permission-allow',
                ];
                for (const sel of permissionSelectors) {
                    const btn = document.querySelector(sel);
                    if (btn) { btn.click(); return 'permission button'; }
                }
                return null;
            })()
        """)
    except Exception:
        pass


async def close_any_dialog(tab):
    """Fecha qualquer diálogo/modal aberto (botão cancelar, X, escape)."""
    await tab.evaluate("""
        (() => {
            // Tenta botão Cancelar
            for (const btn of document.querySelectorAll('button')) {
                const t = btn.textContent?.trim()?.toLowerCase();
                if (t === 'cancelar' || t === 'cancel' || t === 'fechar' || t === 'close') {
                    const r = btn.getBoundingClientRect();
                    if (r.width > 0 && r.height > 0) { btn.click(); return; }
                }
            }
            // Tenta botão X de fechar modal
            const closeBtn = document.querySelector(
                'mat-dialog-container button[aria-label*="close"], ' +
                'mat-dialog-container button[aria-label*="fechar"], ' +
                '.cdk-overlay-pane button.close'
            );
            if (closeBtn) closeBtn.click();
        })()
    """)
    await asyncio.sleep(SHORT_WAIT)


async def close_menu(tab):
    """Fecha menu de contexto aberto clicando fora dele."""
    await tab.evaluate("""
        (() => {
            // Clica no body para fechar menus
            document.body.click();
            // Também tenta ESC
            document.dispatchEvent(new KeyboardEvent('keydown', {key: 'Escape', bubbles: true}));
        })()
    """)
    await asyncio.sleep(SHORT_WAIT)


# ---------------------------------------------------------------------------
# Navegação para aba/página do Power BI
# ---------------------------------------------------------------------------

async def navigate_to_page(tab, page_name: str) -> bool:
    """Navega para uma aba/página específica do Power BI."""
    if not page_name:
        log.info("📌 Nenhuma página alvo definida, permanecendo na página atual")
        return True

    log.info(f"🔍 Navegando para aba '{page_name}'...")

    result = await tab.evaluate(f"""
        (() => {{
            const name = "{page_name}".toUpperCase();
            
            // Busca nas abas de navegação
            const navSelectors = [
                'li.section', 'li[class*="section"]',
                'ul.pane li', '[role="tab"]', '[role="listitem"]',
            ];
            for (const sel of navSelectors) {{
                for (const item of document.querySelectorAll(sel)) {{
                    const text = item.textContent?.trim();
                    if (text && text.toUpperCase().includes(name)) {{
                        const target = item.querySelector('span, div, a') || item;
                        target.scrollIntoView({{block: 'center'}});
                        target.click();
                        return `Aba: "${{text.substring(0, 50)}}"`;
                    }}
                }}
            }}
            
            // Busca em elementos pequenos com o texto
            for (const el of document.querySelectorAll('span, a, button, label, h3, h4')) {{
                const text = el.textContent?.trim();
                if (text && text.toUpperCase().includes(name) && text.length < 40) {{
                    const rect = el.getBoundingClientRect();
                    if (rect.width > 5 && rect.height > 5) {{
                        el.scrollIntoView({{block: 'center'}});
                        el.click();
                        return `Texto: "${{text.substring(0, 40)}}"`;
                    }}
                }}
            }}
            
            // Busca por aria-label/title
            const ariaEl = document.querySelector(`[aria-label*="${{name}}"]`)
                        || document.querySelector(`[title*="${{name}}"]`);
            if (ariaEl) {{ ariaEl.click(); return "aria-label/title"; }}
            
            return null;
        }})()
    """)

    if result:
        log.info(f"  ✅ Navegou: {result}")
        return True

    log.error(f"❌ Aba '{page_name}' não encontrada")
    return False


async def press_escape(tab, times: int = 1, wait_each: float = 0.4):
    """
    Pressiona ESC algumas vezes para fechar menus, tooltips, dialogs e overlays.
    Isso ajuda muito no Power BI, que costuma deixar camadas abertas após hover/click.
    """
    for _ in range(times):
        try:
            await tab.send(uc.cdp.input_.dispatch_key_event(
                type_="keyDown",
                windows_virtual_key_code=27,
                key="Escape",
                code="Escape",
            ))
            await tab.send(uc.cdp.input_.dispatch_key_event(
                type_="keyUp",
                windows_virtual_key_code=27,
                key="Escape",
                code="Escape",
            ))
        except Exception:
            pass
        await asyncio.sleep(wait_each)


async def close_open_menus_and_overlays(tab, aggressive: bool = False):
    """
    Fecha menus, popups e overlays residuais do Power BI sem sair clicando em links externos.
    """
    await press_escape(tab, times=2 if aggressive else 1, wait_each=0.3)

    try:
        await tab.evaluate("""
            (() => {
                const textOf = (el) => (el?.innerText || el?.textContent || '').trim().toLowerCase();

                const isLearnLink = (el) => {
                    const txt = textOf(el);
                    const href = (el?.href || '').toLowerCase();
                    return (
                        txt.includes('saiba mais sobre como exportar dados') ||
                        txt.includes('learn more about exporting data') ||
                        href.includes('learn.microsoft.com') ||
                        href.includes('microsoft.com')
                    );
                };

                const candidates = Array.from(document.querySelectorAll(
                    'button, [role="button"], .close, .close-btn, .close-button, ' +
                    '.dialog-close, .modal-close, [data-testid*="close"], [class*="close"]'
                ));

                for (const el of candidates) {
                    const txt = textOf(el);
                    const aria = (el.getAttribute('aria-label') || '').trim().toLowerCase();
                    const title = (el.getAttribute('title') || '').trim().toLowerCase();

                    if (isLearnLink(el)) continue;

                    const rect = el.getBoundingClientRect();
                    if (rect.width <= 0 || rect.height <= 0) continue;

                    const shouldClose =
                        txt === 'x' || aria === 'x' || title === 'x' ||
                        txt.includes('fechar') || txt.includes('close') ||
                        aria.includes('fechar') || aria.includes('close') ||
                        title.includes('fechar') || title.includes('close') ||
                        txt.includes('cancelar') || txt.includes('cancel') ||
                        aria.includes('cancelar') || aria.includes('cancel');

                    if (shouldClose) {
                        try { el.click(); } catch (e) {}
                    }
                }
            })()
        """)
    except Exception:
        pass

    await asyncio.sleep(OVERLAY_SETTLE_WAIT if aggressive else 0.5)

async def cleanup_residual_ui(tab, stage_label: str, aggressive: bool = False):
    """
    Rotina centralizada para limpar a interface entre etapas.

    Essa limpeza evita herdar menu/dialog/overlay antigo no passo seguinte.
    É usada antes de scans, antes de abrir "Mais opções" e após exportação.
    """
    try:
        before_count = await tab.evaluate("""
            (() => {
                const visible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };
                const selectors = [
                    '[role="menu"]',
                    '[role="dialog"]',
                    '[role="listbox"]',
                    '[aria-modal="true"]',
                    '.cdk-overlay-pane',
                    '.contextMenu',
                    '.dropdown-menu',
                    '.popup',
                    '.modal'
                ].join(',');
                return Array.from(document.querySelectorAll(selectors)).filter(visible).length;
            })()
        """)
        before_count = int(before_count or 0)
    except Exception:
        before_count = -1

    log.info(f"🧹 Limpando interface ({stage_label})...")
    await close_open_menus_and_overlays(tab, aggressive=aggressive)

    try:
        after_count = await tab.evaluate("""
            (() => {
                const visible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };
                const selectors = [
                    '[role="menu"]',
                    '[role="dialog"]',
                    '[role="listbox"]',
                    '[aria-modal="true"]',
                    '.cdk-overlay-pane',
                    '.contextMenu',
                    '.dropdown-menu',
                    '.popup',
                    '.modal'
                ].join(',');
                return Array.from(document.querySelectorAll(selectors)).filter(visible).length;
            })()
        """)
        after_count = int(after_count or 0)
    except Exception:
        after_count = -1

    if before_count >= 0 and after_count >= 0:
        log.info(f"  ✅ Overlays visíveis: antes={before_count} | depois={after_count}")
    else:
        log.info("  ✅ Limpeza aplicada (contagem de overlays indisponível)")


async def dismiss_sensitive_data_popup(tab, max_rounds: int = 6, preserve_export_dialog: bool = False) -> bool:
    """
    Fecha popup de confidencialidade sem fechar o diálogo de exportação quando preserve_export_dialog=True.
    """
    handled_any = False
    preserve_flag = "true" if preserve_export_dialog else "false"

    for _ in range(max_rounds):
        try:
            result = await tab.evaluate(f"""
                (() => {{
                    const PRESERVE_EXPORT_DIALOG = {preserve_flag};
                    const textOf = (el) => (el?.innerText || el?.textContent || '').trim();
                    const lower = (s) => (s || '').toLowerCase();

                    const dialogs = Array.from(document.querySelectorAll([
                        '[role="dialog"]',
                        '[aria-modal="true"]',
                        '.modal',
                        '.popup',
                        '.dialog',
                    ].join(',')));

                    const visibleDialogs = dialogs.filter(d => {{
                        const r = d.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    }});

                    const allCandidates = [];

                    for (const dlg of visibleDialogs) {{
                        const dlgText = lower(textOf(dlg));
                        const isSensitive =
                            dlgText.includes('confidencial') ||
                            dlgText.includes('confidential') ||
                            dlgText.includes('copiando dados') ||
                            dlgText.includes('copying data') ||
                            dlgText.includes('exportar dados') ||
                            dlgText.includes('exporting data');
                        if (!isSensitive) continue;

                        const looksLikeExportDialog =
                            dlgText.includes('.xlsx') ||
                            dlgText.includes('dados resumidos') ||
                            dlgText.includes('dados subjacentes') ||
                            dlgText.includes('data with current layout');
                        if (looksLikeExportDialog && PRESERVE_EXPORT_DIALOG) continue;

                        const els = dlg.querySelectorAll('button, [role="button"], [aria-label], [title]');
                        for (const el of els) {{
                            const txt = lower(textOf(el));
                            const aria = lower(el.getAttribute('aria-label') || '');
                            const title = lower(el.getAttribute('title') || '');
                            const rect = el.getBoundingClientRect();
                            if (!(rect.width > 0 && rect.height > 0)) continue;

                            const isLearn =
                                txt.includes('saiba mais sobre como exportar dados') ||
                                txt.includes('learn more about exporting data') ||
                                txt.includes('saiba mais') ||
                                txt.includes('learn more');
                            if (isLearn) continue;

                            allCandidates.push({{ el, txt, aria, title }});
                        }}
                    }}

                    for (const item of allCandidates) {{
                        const {{ el, txt, aria, title }} = item;
                        const shouldClose =
                            txt === 'x' || aria === 'x' || title === 'x' ||
                            txt.includes('fechar') || txt.includes('close') ||
                            txt.includes('cancelar') || txt.includes('cancel') ||
                            aria.includes('fechar') || aria.includes('close') ||
                            aria.includes('cancelar') || aria.includes('cancel') ||
                            title.includes('fechar') || title.includes('close');
                        if (shouldClose) {{
                            try {{ el.click(); }} catch (e) {{}}
                            return "closed";
                        }}
                    }}

                    return "";
                }})()
            """)
        except Exception:
            result = ""

        if result:
            handled_any = True
            log.info(f"  ⚠️🔒 Popup confidencial tratado com ação segura: {result}")
            await asyncio.sleep(1.2)
            await press_escape(tab, times=1, wait_each=0.3)
            continue

        await close_open_menus_and_overlays(tab, aggressive=True)
        await asyncio.sleep(0.8)
        break

    return handled_any


async def block_microsoft_learn_and_external_links(tab):
    """
    Injeta um bloqueio no DOM para evitar que links do Microsoft Learn ou outros links externos
    sejam abertos por engano durante a automação.
    """
    try:
        await tab.evaluate("""
            (() => {
                if (window.__pbi_external_link_blocked__) return true;
                window.__pbi_external_link_blocked__ = true;

                const shouldBlock = (el) => {
                    const href = (el?.getAttribute?.('href') || '').toLowerCase();
                    const txt = (el?.innerText || el?.textContent || '').toLowerCase();
                    return (
                        href.includes('learn.microsoft.com') ||
                        href.includes('microsoft.com') ||
                        txt.includes('saiba mais sobre como exportar dados') ||
                        txt.includes('learn more about exporting data') ||
                        txt.includes('saiba mais') ||
                        txt.includes('learn more')
                    );
                };

                document.addEventListener('click', function(ev) {
                    let el = ev.target;
                    while (el) {
                        if (el.tagName === 'A') {
                            if (shouldBlock(el)) {
                                ev.preventDefault();
                                ev.stopPropagation();
                                ev.stopImmediatePropagation();
                                return false;
                            }
                            break;
                        }
                        // Em alguns casos o clique cai em span/div dentro do link.
                        if (el.closest && el.closest('a')) {
                            const link = el.closest('a');
                            if (shouldBlock(link)) {
                                ev.preventDefault();
                                ev.stopPropagation();
                                ev.stopImmediatePropagation();
                                return false;
                            }
                        }
                        el = el.parentElement;
                    }
                }, true);

                return true;
            })()
        """)
        log.info("🛡️ Bloqueio de links externos/Microsoft Learn ativado")
    except Exception:
        pass


async def ensure_report_tab_still_valid(tab, expected_url: str):
    """
    Garante que a aba continua no relatório, e não foi desviada para página externa.
    """
    try:
        current_url = await get_tab_url(tab)
    except Exception:
        current_url = ""

    current = (current_url or "").lower()
    expected = (expected_url or "").lower()

    if not current:
        return tab

    if "learn.microsoft.com" in current or (
        "microsoft.com" in current and "powerbi.com" not in current
    ):
        log.warning("⚠️ Aba foi desviada para página externa. Tentando voltar ao relatório...")
        try:
            await tab.get(expected_url)
            await asyncio.sleep(LONG_WAIT)
        except Exception:
            pass

    return tab

# ---------------------------------------------------------------------------
# Scan: descobrir todos os visuais com "Mais opções"
# ---------------------------------------------------------------------------

async def hover_all_visuals(tab):
    """Faz hover em TODOS os visual-containers para revelar botões de opções."""
    log.info("🖱️  Fazendo hover em todos os visual-containers...")
    await tab.evaluate("""
        (() => {
            const containers = document.querySelectorAll('visual-container');
            containers.forEach((vc, i) => {
                const el = vc.querySelector('transform') || vc;
                const rect = el.getBoundingClientRect();
                if (rect.width > 30 && rect.height > 30) {
                    ['pointerenter','pointerover','mouseenter','mouseover',
                     'mousemove','pointermove'
                    ].forEach(type => {
                        el.dispatchEvent(new PointerEvent(type, {
                            bubbles: true, composed: true, view: window,
                            clientX: rect.x + rect.width / 2,
                            clientY: rect.y + rect.height / 2
                        }));
                    });
                }
            });
        })()
    """)
    await asyncio.sleep(MEDIUM_WAIT)


# ---------------------------------------------------------------------------
# Helpers extras para visuais / JSON
# ---------------------------------------------------------------------------

async def eval_json(tab, script: str, default=None):
    """Executa JS que retorna JSON.stringify(...)."""
    try:
        raw = await tab.evaluate(script)
        if raw is None:
            return default
        return json.loads(str(raw))
    except Exception:
        return default


async def scroll_visual_into_view_by_index(tab, idx: int) -> bool:
    """Rola o visual para o centro da viewport para reduzir falhas por virtualização."""
    result = await tab.evaluate(f"""
        (() => {{
            const vc = document.querySelectorAll('visual-container')[{idx}];
            if (!vc) return false;
            const el = vc.querySelector('transform') || vc;
            el.scrollIntoView({{block: 'center', inline: 'center', behavior: 'instant'}});
            return true;
        }})()
    """)
    await asyncio.sleep(1.2)
    return bool(result)


async def get_visual_center(tab, idx: int):
    """Obtém centro e tamanho do visual em JSON normalizado."""
    return await eval_json(tab, f"""
        (() => {{
            const vc = document.querySelectorAll('visual-container')[{idx}];
            if (!vc) return JSON.stringify(null);
            const el = vc.querySelector('transform') || vc;
            const rect = el.getBoundingClientRect();
            return JSON.stringify({{
                cx: Math.round(rect.left + rect.width / 2),
                cy: Math.round(rect.top + rect.height / 2),
                w: Math.round(rect.width),
                h: Math.round(rect.height),
                top: Math.round(rect.top),
                left: Math.round(rect.left)
            }});
        }})()
    """, default=None)


async def hover_visual(tab, idx: int) -> bool:
    """Faz hover real no visual após scroll até a área visível."""
    await scroll_visual_into_view_by_index(tab, idx)
    coords = await get_visual_center(tab, idx)
    if not coords:
        return False

    cx = int(coords.get('cx', 0))
    cy = int(coords.get('cy', 0))
    if cx <= 0 or cy <= 0:
        return False

    try:
        await tab.send(uc.cdp.input_.dispatch_mouse_event(type_='mouseMoved', x=cx, y=cy))
        await asyncio.sleep(0.25)
        await tab.send(uc.cdp.input_.dispatch_mouse_event(type_='mouseMoved', x=cx + 2, y=cy + 2))
    except Exception:
        await tab.evaluate(f"""
            (() => {{
                const vc = document.querySelectorAll('visual-container')[{idx}];
                if (!vc) return false;
                const el = vc.querySelector('transform') || vc;
                const rect = el.getBoundingClientRect();
                ['pointerenter','pointerover','mouseenter','mouseover','mousemove','pointermove'].forEach(type => {{
                    el.dispatchEvent(new PointerEvent(type, {{
                        bubbles: true,
                        composed: true,
                        view: window,
                        clientX: rect.left + rect.width / 2,
                        clientY: rect.top + rect.height / 2
                    }}));
                }});
                return true;
            }})()
        """)
    await asyncio.sleep(1.0)
    return True


async def open_visual_more_options(tab, idx: int, attempts: int = 5) -> bool:
    """Abre o menu de Mais opções do visual com retries e re-scroll."""
    for attempt in range(1, attempts + 1):
        await cleanup_residual_ui(
            tab,
            stage_label=f"antes de abrir 'Mais opções' (visual-container #{idx})",
            aggressive=True,
        )
        await hover_visual(tab, idx)
        opened = await tab.evaluate(f"""
            (() => {{
                const vc = document.querySelectorAll('visual-container')[{idx}];
                if (!vc) return false;

                const selectors = [
                    'button[aria-label*="Mais opções"]',
                    'button[aria-label*="More options"]',
                    'button[aria-label*="opções"]',
                    'button[aria-label*="options"]',
                    'button[title*="Mais opções"]',
                    'button[title*="More options"]',
                    'button[title*="opções"]',
                    'button[title*="options"]',
                    'button[class*="more-options"]',
                    'button[class*="moreOptions"]',
                    'visual-header-item-container button',
                    'visual-container-options-menu button',
                    'visual-container-header button'
                ];

                for (const sel of selectors) {{
                    const btns = vc.querySelectorAll(sel);
                    for (const btn of btns) {{
                        const rect = btn.getBoundingClientRect();
                        if (rect.width > 0 && rect.height > 0 && rect.width < 140) {{
                            btn.scrollIntoView({{block: 'center', inline: 'center', behavior: 'instant'}});
                            btn.dispatchEvent(new MouseEvent('mouseover', {{ bubbles: true }}));
                            btn.click();
                            return true;
                        }}
                    }}
                }}

                const header = vc.querySelector('visual-container-header');
                if (header) {{
                    const anyBtn = header.querySelector('button');
                    if (anyBtn) {{
                        anyBtn.click();
                        return true;
                    }}
                }}

                return false;
            }})()
        """)
        if opened:
            await asyncio.sleep(1.2)
            return True
        log.info(f"    Tentativa {attempt}/{attempts}: botão 'Mais opções' ainda não abriu")
        await close_menu(tab)
        await asyncio.sleep(0.8)
    return False


async def menu_contains_export(tab) -> bool:
    result = await tab.evaluate("""
        (() => {
            const nodes = document.querySelectorAll('[role="menuitem"], [role="menuitemcheckbox"], [role="option"], button, li, a, div[tabindex], span[tabindex]');
            for (const item of nodes) {
                const text = item.textContent?.trim()?.toLowerCase() || '';
                const rect = item.getBoundingClientRect();
                if (rect.width > 10 && rect.height > 10 && (
                    text === 'exportar dados' ||
                    text === 'export data' ||
                    text.includes('exportar dados') ||
                    text.includes('export data')
                )) {
                    return true;
                }
            }
            return false;
        })()
    """)
    return bool(result)


async def click_export_data_menu(tab):
    return await tab.evaluate("""
        (() => {
            const candidates = document.querySelectorAll('[role="menuitem"], [role="menuitemcheckbox"], [role="option"], button, li, a, div[tabindex], span[tabindex]');
            for (const item of candidates) {
                const text = item.textContent?.trim() || '';
                const lower = text.toLowerCase();
                const rect = item.getBoundingClientRect();
                if (rect.width > 10 && rect.height > 10 && (
                    lower === 'exportar dados' || lower === 'export data' ||
                    lower.includes('exportar dados') || lower.includes('export data')
                )) {
                    item.click();
                    return text.substring(0, 60);
                }
            }
            return null;
        })()
    """)


async def probe_visual_export(tab, idx: int) -> bool:
    """Confirma se o visual suporta Exportar dados, inclusive fora da dobra."""
    if not await open_visual_more_options(tab, idx, attempts=3):
        return False
    has_export = await menu_contains_export(tab)
    await close_menu(tab)
    return has_export

async def scroll_visual_into_view(tab, visual):
    """
    Rola o visual para o centro da viewport e dispara pointer events para
    forçar o PBI a renderizar o visual-container-header.
    """
    idx = visual.get("index")
    try:
        await tab.evaluate(f"""
            (() => {{
                const vcs = Array.from(document.querySelectorAll('visual-container'));
                const vc = vcs[{idx}];
                if (!vc) return;
                vc.scrollIntoView({{ behavior: 'auto', block: 'center', inline: 'center' }});
                const r = vc.getBoundingClientRect();
                const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
                ['pointerenter','pointerover','mouseenter','mouseover','mousemove','pointermove'].forEach(t => {{
                    vc.dispatchEvent(new PointerEvent(t, {{
                        bubbles: true, composed: true, view: window,
                        clientX: cx, clientY: cy
                    }}));
                }});
            }})()
        """)
    except Exception:
        pass
    await asyncio.sleep(1.2)


async def hover_visual_center(tab, visual):
    """
    Faz hover no centro do visual com base nas coordenadas atuais do DOM.
    """
    idx = visual.get("index")
    try:
        coords = await tab.evaluate(f"""
            (() => {{
                const vcs = Array.from(document.querySelectorAll('visual-container'));
                const vc = vcs[{idx}];
                if (!vc) return null;
                const r = vc.getBoundingClientRect();
                return JSON.stringify({{
                    x: Math.round(r.left + (r.width / 2)),
                    y: Math.round(r.top + (r.height / 2)),
                    width: Math.round(r.width),
                    height: Math.round(r.height)
                }});
            }})()
        """)
        data = json.loads(str(coords)) if coords else None
    except Exception:
        data = None

    if not data:
        return False

    try:
        await tab.mouse.move(data["x"], data["y"])
        await asyncio.sleep(1.2)
        return True
    except Exception:
        return False


async def get_clean_visual_menu_items(tab):
    """
    Lê apenas itens de menu realmente relacionados ao visual,
    ignorando overlays residuais de filtro, tooltip e links do Learn.
    """
    try:
        raw = await tab.evaluate("""
            (() => {
                const textOf = (el) => (el?.innerText || el?.textContent || '').trim();
                const lower = (s) => (s || '').toLowerCase();

                const isVisible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };

                const overlayRoots = Array.from(document.querySelectorAll([
                    '[role="menu"]',
                    '[role="listbox"]',
                    '[role="dialog"]',
                    '.contextMenu',
                    '.menu',
                    '.dropdown-menu',
                    '.qtip',
                ].join(',')))
                .filter(isVisible);

                const allItems = [];

                for (const root of overlayRoots) {
                    const rootText = lower(textOf(root));

                    const rootLooksLikeVisualMenu =
                        root.querySelector('[role="menuitem"]') ||
                        rootText.includes('exportar dados') ||
                        rootText.includes('mostrar como uma tabela') ||
                        rootText.includes('obter insights') ||
                        rootText.includes('sort axis') ||
                        rootText.includes('classificar eixo');

                    const rootLooksLikeFilterOverlay =
                        rootText.includes('não está em branco') ||
                        rootText.includes('is not blank') ||
                        rootText.includes('é 20') ||
                        rootText.includes('filtro') ||
                        rootText.includes('filter');

                    if (!rootLooksLikeVisualMenu || rootLooksLikeFilterOverlay) {
                        continue;
                    }

                    const items = root.querySelectorAll('button, [role="menuitem"], [role="menuitemcheckbox"], li, a');
                    for (const el of items) {
                        if (!isVisible(el)) continue;

                        const txt = textOf(el);
                        const txtLow = lower(txt);
                        const role = (el.getAttribute('role') || '').trim();
                        const tag = (el.tagName || '').trim();
                        const href = (el.getAttribute('href') || '').trim().toLowerCase();

                        const shouldIgnore =
                            !txt ||
                            txtLow.includes('saiba mais sobre como exportar dados') ||
                            txtLow.includes('learn more about exporting data') ||
                            href.includes('learn.microsoft.com') ||
                            href.includes('microsoft.com') ||
                            txtLow === 'é 202602' ||
                            txtLow === 'não está em branco';

                        if (shouldIgnore) continue;

                        allItems.push({
                            text: txt,
                            role,
                            tag
                        });
                    }
                }

                return JSON.stringify(allItems);
            })()
        """)
        items = json.loads(str(raw)) if raw else []
    except Exception:
        items = []

    # remove duplicados mantendo ordem
    seen = set()
    clean = []
    for item in items:
        key = (item.get("text", ""), item.get("role", ""), item.get("tag", ""))
        if key in seen:
            continue
        seen.add(key)
        clean.append(item)

    return clean


async def visual_menu_has_export_option(tab) -> bool:
    """
    Confirma se o menu atualmente aberto é de visual e possui 'Exportar dados'.
    """
    items = await get_clean_visual_menu_items(tab)

    if not items:
        return False

    for item in items:
        txt = (item.get("text") or "").strip().lower()
        if "exportar dados" in txt:
            return True

    return False

async def locate_more_options_button(tab, visual):
    """
    Localiza o ÚLTIMO botão do header (rightmost = Mais opções).
    Usa JSON.stringify para garantir retorno string ao nodriver.
    """
    idx = visual.get("index")
    try:
        raw = await tab.evaluate(f"""
            (() => {{
                const vcs = Array.from(document.querySelectorAll('visual-container'));
                const vc = vcs[{idx}];
                if (!vc) return JSON.stringify(null);
                const vcRect = vc.getBoundingClientRect();
                let btns = Array.from(vc.querySelectorAll('button')).filter(b => {{
                    const r = b.getBoundingClientRect();
                    return r.width > 1 && r.height > 1 &&
                           r.top >= vcRect.top && r.top <= vcRect.top + 50;
                }});
                if (btns.length === 0) {{
                    btns = Array.from(vc.querySelectorAll('button')).filter(b => {{
                        const r = b.getBoundingClientRect();
                        return r.width > 1 && r.height > 1;
                    }});
                }}
                if (btns.length === 0) return JSON.stringify(null);
                btns.sort((a, b) => b.getBoundingClientRect().left - a.getBoundingClientRect().left);
                const btn = btns[0];
                const r = btn.getBoundingClientRect();
                return JSON.stringify({{
                    x: Math.round(r.left + r.width / 2),
                    y: Math.round(r.top + r.height / 2),
                    width: Math.round(r.width),
                    height: Math.round(r.height),
                    total: btns.length
                }});
            }})()
        """)
        import json as _j
        if isinstance(raw, str):
            return _j.loads(raw)
        return raw if isinstance(raw, dict) else None
    except Exception:
        return None

async def click_more_options_button(tab, visual):
    """
    Tenta clicar no botão 'Mais opções' usando localização atual.
    """
    coords = await locate_more_options_button(tab, visual)
    if not coords:
        return False

    try:
        await tab.mouse.move(coords["x"], coords["y"])
        await asyncio.sleep(0.8)
        await tab.mouse.click(coords["x"], coords["y"])
        await asyncio.sleep(MEDIUM_WAIT)
        return True
    except Exception:
        return False


async def is_visual_menu_open(tab) -> bool:
    """
    Verifica se um menu de visual do Power BI está aberto.
    Cobre as variações de renderização do PBI (role=menu, cdk-overlay, etc).
    """
    try:
        result = await tab.evaluate("""
            (() => {
                const isVisible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };
                const textOf = (el) => (el?.innerText || el?.textContent || '').toLowerCase();
                const roots = Array.from(document.querySelectorAll(
                    '[role="menu"], [role="listbox"], .contextMenu, .dropdown-menu, ' +
                    '.menu, [class*="contextMenu"], [class*="ContextMenu"], ' +
                    '[class*="dropdown"], [class*="Dropdown"], ' +
                    '.cdk-overlay-pane, [class*="overlayPane"]'
                )).filter(isVisible);
                for (const root of roots) {
                    const txt = textOf(root);
                    const looksLikeVisualMenu = (
                        txt.includes('exportar') || txt.includes('export') ||
                        txt.includes('foco') || txt.includes('focus') ||
                        txt.includes('tabela') || txt.includes('table') ||
                        txt.includes('insights') || txt.includes('classificar') ||
                        txt.includes('sort') || txt.includes('fixar') ||
                        txt.includes('pin') ||
                        root.querySelector('[role="menuitem"], [role="menuitemcheckbox"]')
                    );
                    if (looksLikeVisualMenu) return true;
                }
                return false;
            })()
        """)
        return bool(result)
    except Exception:
        return False


async def open_more_options_robust(tab, visual, attempt: int, retries: int) -> bool:
    """
    Abre 'Mais opções' (botão mais à direita do header) via JS com JSON.stringify.
    """
    idx   = visual.get("index")
    title = visual.get("title", f"Visual #{idx}")

    log.info(f"    🎯 [{attempt}/{retries}] Abrindo 'Mais opções' em {title} (container #{idx})")

    await scroll_visual_into_view(tab, visual)
    await hover_visual_center(tab, visual)
    await asyncio.sleep(0.8)
    await force_hover_visual_header(tab, visual)
    await asyncio.sleep(0.8)

    try:
        result_raw = await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return JSON.stringify({{ok: false, reason: 'no_vc'}});

                const vcRect = vc.getBoundingClientRect();
                const hx = vcRect.left + vcRect.width - 20;
                const hy = vcRect.top + 16;
                ['pointerenter','pointerover','mouseenter','mouseover','mousemove','pointermove'].forEach(t => {{
                    vc.dispatchEvent(new PointerEvent(t, {{
                        bubbles: true, composed: true, view: window,
                        clientX: hx, clientY: hy
                    }}));
                }});

                // Botões do header (primeiros 50px verticais)
                let btns = Array.from(vc.querySelectorAll('button')).filter(b => {{
                    const r = b.getBoundingClientRect();
                    return r.width > 1 && r.height > 1 &&
                           r.top >= vcRect.top && r.top <= vcRect.top + 50;
                }});

                // Sem filtro de altura se vazio
                if (btns.length === 0) {{
                    btns = Array.from(vc.querySelectorAll('button')).filter(b => {{
                        const r = b.getBoundingClientRect();
                        return r.width > 1 && r.height > 1;
                    }});
                }}
                if (btns.length === 0) return JSON.stringify({{ok: false, reason: 'no_buttons'}});

                // Mais opções = botão mais à direita
                btns.sort((a, b) => b.getBoundingClientRect().left - a.getBoundingClientRect().left);
                const btn = btns[0];
                const r = btn.getBoundingClientRect();
                btn.dispatchEvent(new MouseEvent('mouseover', {{bubbles: true}}));
                btn.click();
                return JSON.stringify({{
                    ok: true,
                    reason: 'rightmost_btn',
                    x: Math.round(r.left + r.width / 2),
                    y: Math.round(r.top + r.height / 2),
                    w: Math.round(r.width),
                    h: Math.round(r.height),
                    total: btns.length
                }});
            }})()
        """)
    except Exception as e:
        result_raw = None
        log.info(f"      ⚠️ JS exception: {e}")

    # Normaliza resultado — nodriver pode retornar str, dict, list ou None
    import json as _json
    result = {"ok": False, "reason": "not_set"}
    if result_raw is None:
        result = {"ok": False, "reason": "js_none"}
    elif isinstance(result_raw, str):
        try:
            result = _json.loads(result_raw)
        except Exception:
            result = {"ok": False, "reason": f"parse_err:{result_raw[:60]}"}
    elif isinstance(result_raw, dict):
        result = result_raw
    else:
        result = {"ok": False, "reason": f"type:{type(result_raw).__name__}:{str(result_raw)[:80]}"}

    if not result.get("ok"):
        log.info(f"      ⚠️ JS não clicou: {result.get('reason')}")
        log.info(f"      ❌ tentativa {attempt}/{retries} sem sucesso para abrir 'Mais opções'")
        return False

    log.info(
        f"      • clicado: x={result.get('x')} y={result.get('y')} "
        f"(w={result.get('w')} h={result.get('h')}) "
        f"[{result.get('reason')}] total_btns={result.get('total')}"
    )

    # Verifica múltiplas vezes se o menu apareceu
    for check in range(6):
        await asyncio.sleep(0.4)
        if await is_visual_menu_open(tab):
            log.info(f"      ✅ menu aberto (check {check + 1})")
            return True

    log.info("      ⚠️ clique JS executou mas menu não ficou visível")
    log.info(f"      ❌ tentativa {attempt}/{retries} sem sucesso para abrir 'Mais opções'")
    return False


async def force_hover_visual_header(tab, visual):
    """
    Faz hover na região superior direita do visual, onde normalmente mora o header.
    """
    idx = visual.get("index")

    try:
        raw = await tab.evaluate(f"""
            (() => {{
                const vcs = Array.from(document.querySelectorAll('visual-container'));
                const vc = vcs[{idx}];
                if (!vc) return null;
                const r = vc.getBoundingClientRect();
                return JSON.stringify({{
                    x: Math.round(r.left + r.width - 18),
                    y: Math.round(r.top + 16)
                }});
            }})()
        """)
        pos = json.loads(str(raw)) if raw else None
    except Exception:
        pos = None

    if not pos:
        return False

    try:
        await tab.mouse.move(pos["x"], pos["y"])
        await asyncio.sleep(1.0)
        return True
    except Exception:
        return False


async def try_open_visual_menu_and_confirm_export(tab, visual, retries=5):
    """
    Tenta abrir o menu real do visual e confirmar se ele contém 'Exportar dados'.

    Fluxo:
    1. limpa overlays
    2. centraliza visual
    3. hover no centro
    4. hover no header
    5. tenta clicar em 'Mais opções'
    6. valida se o menu aberto é mesmo de visual
    """
    idx = visual.get("index")
    title = visual.get("title", f"Visual #{idx}")

    for attempt in range(1, retries + 1):
        await cleanup_residual_ui(
            tab,
            stage_label=f"antes de abrir 'Mais opções' (visual #{idx})",
            aggressive=True,
        )
        clicked = await open_more_options_robust(tab, visual, attempt, retries)
        if not clicked:
            await asyncio.sleep(RETRY_MENU_WAIT)
            continue

        # pequena espera para menu montar
        await asyncio.sleep(1.5)

        if await visual_menu_has_export_option(tab):
            return True

        # debug opcional do menu lido
        items = await get_clean_visual_menu_items(tab)
        if items:
            log.info(f"  📋 Menu aberto com {len(items)} itens")
            for item in items[:12]:
                log.info(f"    tag={item.get('tag','')} role='{item.get('role','')}' text='{item.get('text','')}'")

        log.info(f"    Tentativa {attempt}/{retries}: menu abriu, mas sem 'Exportar dados' válido para {title}")
        await cleanup_residual_ui(
            tab,
            stage_label=f"após falha de menu do visual #{idx}",
            aggressive=True,
        )
        await asyncio.sleep(RETRY_MENU_WAIT)

    return False


async def probe_visual_export_status(tab, visual, retries=3):
    """
    Avalia estado de exportação do visual separando os cenários:
    - menu não abriu
    - menu abriu sem exportação
    - menu abriu com exportação
    """
    idx = visual.get("index")

    for attempt in range(1, retries + 1):
        await cleanup_residual_ui(
            tab,
            stage_label=f"probe de exportação - visual #{idx} tentativa {attempt}",
            aggressive=True,
        )
        clicked = await open_more_options_robust(tab, visual, attempt, retries)
        if not clicked:
            await asyncio.sleep(RETRY_MENU_WAIT)
            continue

        await asyncio.sleep(1.2)
        items = await get_clean_visual_menu_items(tab)
        if not items:
            await cleanup_residual_ui(
                tab,
                stage_label=f"menu inválido/residual - visual #{idx}",
                aggressive=True,
            )
            await asyncio.sleep(RETRY_MENU_WAIT)
            continue

        has_export = any(
            "exportar dados" in (item.get("text", "").strip().lower()) or
            "export data" in (item.get("text", "").strip().lower())
            for item in items
        )

        await cleanup_residual_ui(
            tab,
            stage_label=f"fechando menu após probe - visual #{idx}",
            aggressive=True,
        )
        if has_export:
            return {
                "menu_opened": True,
                "has_export": True,
                "reason": "menu_visual_com_exportacao",
            }
        return {
            "menu_opened": True,
            "has_export": False,
            "reason": "menu_visual_sem_exportacao",
        }

    return {
        "menu_opened": False,
        "has_export": False,
        "reason": "menu_nao_abriu",
    }


async def scan_visuals(tab, slicer_indexes: set = None):
    """
    Escaneia visuais disponíveis na página e tenta confirmar quais suportam exportação.

    slicer_indexes: conjunto de índices de visual-container já identificados como slicers.
    Esses containers são pulados no probe de exportação.
    """
    if slicer_indexes is None:
        slicer_indexes = set()

    log.info("🔎 Escaneando visuais na página (base v5)...")
    await get_report_dom_context_info(tab, caller="scan_visuals (pré-scan)")
    await close_open_menus_and_overlays(tab, aggressive=True)

    minimal_raw = await tab.evaluate("""
        (() => {
            try {
                return JSON.stringify({
                    ok: true,
                    count: document.querySelectorAll('visual-container').length
                });
            } catch (err) {
                return JSON.stringify({
                    ok: false,
                    error: String(err && err.message ? err.message : err)
                });
            }
        })()
    """)
    log.info(f"🧪 [scan_visuals] evaluate mínimo bruto: {minimal_raw}")

    payload_raw = await tab.evaluate("""
        (() => {
            try {
                const results = [];
                const discardReasons = {};
                const incDiscard = (reason) => {
                    discardReasons[reason] = (discardReasons[reason] || 0) + 1;
                };

                const resolveReportDocument = () => {
                    const mainCount = document.querySelectorAll('visual-container').length;
                    if (mainCount > 0) return { doc: document, source: 'document' };
                    const iframes = Array.from(document.querySelectorAll('iframe'));
                    for (let i = 0; i < iframes.length; i++) {
                        try {
                            const d = iframes[i].contentDocument;
                            if (!d) continue;
                            if (d.querySelectorAll('visual-container').length > 0) {
                                return { doc: d, source: `iframe[${i}]` };
                            }
                        } catch (e) {}
                    }
                    return { doc: document, source: 'document' };
                };

                const reportCtx = resolveReportDocument();
                const containers = Array.from(reportCtx.doc.querySelectorAll('visual-container'));
                const rawCount = containers.length;

                containers.forEach((vc, index) => {
                    const el = vc.querySelector('transform') || vc;
                    if (!el) { incDiscard('no_host_element'); return; }

                    const rect = el.getBoundingClientRect();
                    const w = Math.round(rect.width);
                    const h = Math.round(rect.height);

                    if (w < 50 || h < 50) {
                        incDiscard('tiny_rect(<50x50)');
                        return;
                    }
                    if (w > h * 20 || h > w * 10) {
                        incDiscard('extreme_aspect_ratio');
                        return;
                    }
                    if (w * h < 5000) {
                        incDiscard('area_too_small(<5000)');
                        return;
                    }

                    let title = '';
                    let type = 'desconhecido';

                    const headerText = vc.querySelector(
                        '.slicer-header-text, .visual-title, .visualTitle, [class*="title"], h2, h3, h4'
                    );
                    if (headerText) {
                        title = headerText.textContent?.trim()?.substring(0, 80) || '';
                    }

                    const ariaLabel = el.getAttribute('aria-label') || vc.getAttribute('aria-label') || '';
                    if (!title && ariaLabel) {
                        title = ariaLabel.substring(0, 80);
                    }

                    const allClasses = [
                        vc.className || '',
                        el.className || '',
                        ...Array.from(vc.querySelectorAll('[class]')).slice(0, 30).map(e => String(e.className || ''))
                    ].join(' ').toLowerCase();

                    if (allClasses.includes('tablix') || allClasses.includes('table') || allClasses.includes('pivot')) type = 'Tabela';
                    else if (allClasses.includes('slicer')) type = 'Slicer';
                    else if (allClasses.includes('card'))   type = 'Card';
                    else if (allClasses.includes('kpi'))    type = 'KPI';
                    else if (allClasses.includes('map'))    type = 'Mapa';
                    else if (allClasses.includes('chart') || allClasses.includes('bar') || allClasses.includes('line')) type = 'Gráfico';

                    if (!title) {
                        const textContent = el.textContent?.replace(/\\s+/g, ' ').trim()?.substring(0, 100) || '';
                        if (textContent.length > 5 && textContent.length < 80) title = textContent;
                    }

                    const optionsBtn = vc.querySelector(
                        'button[class*="more-options"], button[class*="moreOptions"], ' +
                        'visual-header-item-container button, visual-container-options-menu button, ' +
                        'button[aria-label*="opções"], button[aria-label*="options"], ' +
                        'button[aria-label*="Mais"], button[aria-label*="More"]'
                    );
                    const hasHeader = !!vc.querySelector(
                        'visual-container-header, visual-container-options-menu, visual-header-item-container'
                    );

                    results.push({
                        index: index,
                        title: title || `Visual #${index + 1}`,
                        type: type,
                        width:  w,
                        height: h,
                        x: Math.round(rect.left),
                        y: Math.round(rect.top),
                        hasOptionsButton: !!optionsBtn,
                        hasHeader: hasHeader,
                        menuOpened:    false,
                        hasExportData: false,
                        exportReason:  'nao_verificado',
                        rawText: (el.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 250)
                    });
                });

                const withHeaderOrBtn = results.filter(v => v.hasHeader || v.hasOptionsButton).length;
                return JSON.stringify({
                    ok: true,
                    payload: {
                        visuals: results,
                        diagnostics: {
                            contextSource: reportCtx.source,
                            rawCount,
                            keptCount: results.length,
                            withHeaderOrBtn,
                            discardedCount: rawCount - results.length,
                            discardReasons
                        }
                    }
                });
            } catch (err) {
                return JSON.stringify({
                    ok: false,
                    error: String(err && err.message ? err.message : err),
                    stack: String(err && err.stack ? err.stack : ''),
                    phase: 'scan_visuals_complex_evaluate'
                });
            }
        })()
    """)
    log.info(f"🧪 [scan_visuals] evaluate complexo bruto: {str(payload_raw)[:500]}")

    try:
        payload_wrapper = json.loads(str(payload_raw)) if payload_raw else {}
    except Exception as exc:
        log.error(f"❌ [scan_visuals] Falha ao parsear retorno bruto do evaluate: {exc}")
        return []

    if not payload_wrapper.get("ok"):
        log.error(
            f"❌ [scan_visuals] erro JS no evaluate ({payload_wrapper.get('phase','?')}): "
            f"{payload_wrapper.get('error','erro desconhecido')}"
        )
        return []

    payload    = payload_wrapper.get("payload") or {}
    visuals    = list(payload.get("visuals") or [])
    diagnostics = payload.get("diagnostics") or {}
    log.info(f"🧪 [scan_visuals] total no payload bruto: {len(payload.get('visuals') or [])}")
    log.info(f"🧪 [scan_visuals] total após parse JSON: {len(visuals)}")

    dedup_map = {}
    for v in visuals:
        idx = v.get("index")
        if idx not in dedup_map:
            dedup_map[idx] = v
    visuals = list(dedup_map.values())
    log.info(f"🧪 [scan_visuals] total após deduplicação (index): {len(visuals)}")

    raw_count       = int(diagnostics.get("rawCount", 0))
    kept_count      = int(diagnostics.get("keptCount", len(visuals)))
    discarded_count = int(diagnostics.get("discardedCount", max(0, raw_count - kept_count)))
    discard_reasons = diagnostics.get("discardReasons", {}) or {}
    context_source  = str(diagnostics.get("contextSource") or "unknown")

    log.info(f"🧪 [diagnóstico] visual-container bruto no DOM: {raw_count} (contexto={context_source})")
    log.info(f"🧪 [diagnóstico] mantidos no scan: {kept_count} | descartados: {discarded_count}")
    if discard_reasons:
        for reason, qty in discard_reasons.items():
            log.info(f"🧪 [diagnóstico] descarte visual: {reason} = {qty}")

    exportable = list(visuals)
    exportable.sort(key=lambda v: (0 if v.get('hasOptionsButton') else 1, v.get('y', 0), v.get('x', 0)))

    log.info(f"📊 Total de visual-containers válidos: {len(exportable)}")
    log.info(f"📋 Visuais com header de opções: {sum(1 for v in exportable if v.get('hasOptionsButton'))}")
    log.info("🔬 Verificando quais visuais suportam 'Exportar dados'...")

    NAV_TITLE_FRAGMENTS = (
        "pressionar enter para explorar",
        "navegação na página",
        "navigation",
    )

    for visual in exportable:
        container_idx = visual.get("index")
        title_lower   = (visual.get("title") or "").lower()

        if container_idx in slicer_indexes:
            visual["menuOpened"]    = False
            visual["hasExportData"] = False
            visual["exportReason"]  = "slicer_ignorado"
            log.info(f"  ⭐  Container #{container_idx} é slicer — pulando probe de exportação.")
            continue

        if any(frag in title_lower for frag in NAV_TITLE_FRAGMENTS):
            visual["menuOpened"]    = False
            visual["hasExportData"] = False
            visual["exportReason"]  = "elemento_navegacao_ignorado"
            log.info(f"  ⭐  Container #{container_idx} parece elemento de navegação — pulando.")
            continue

        await close_open_menus_and_overlays(tab, aggressive=True)
        await scroll_visual_into_view(tab, visual)
        await hover_visual_center(tab, visual)

        probe = await probe_visual_export_status(tab, visual, retries=EXPORT_PROBE_RETRIES)
        visual["menuOpened"]    = bool(probe.get("menu_opened"))
        visual["hasExportData"] = bool(probe.get("has_export"))
        visual["exportReason"]  = str(probe.get("reason") or "indefinido")

        await close_open_menus_and_overlays(tab, aggressive=True)
        await asyncio.sleep(0.5)

    log.info(f"✅ {sum(1 for v in exportable if v.get('hasExportData'))} visuais com 'Exportar dados' confirmado")
    return exportable

def display_visuals(visuals: list):
    """Exibe os visuais encontrados na página."""
    print("\n" + "=" * 90)
    print("📋 VISUAIS ENCONTRADOS NA PÁGINA")
    print("=" * 90)
    print(f"{'#':<4} {'Tipo':<10} {'Título':<34} {'Tam.':<12} {'Btn?':<5} {'Menu?':<6} {'Export?':<8}")
    print("-" * 100)

    for i, v in enumerate(visuals):
        tipo = (v.get("type") or "Tabela")[:10]
        titulo = (v.get("title") or f"Visual #{i+1}")[:34]
        tam = f"{v.get('width', '?')}x{v.get('height', '?')}"
        exporta = "✅" if v.get("hasExportData") else "❌"
        botao = "✅" if v.get("hasOptionsButton") else "❌"
        menu = "✅" if v.get("menuOpened") else "❌"
        print(f"{i:<4} {tipo:<10} {titulo:<34} {tam:<12} {botao:<5} {menu:<6} {exporta:<8}")

        if not v.get("hasExportData"):
            reason = v.get("exportReason", "")
            if reason == "menu_nao_abriu":
                print("     ↳ motivo: menu de visual não abriu com validação")
            elif reason == "menu_visual_sem_exportacao":
                print("     ↳ motivo: menu abriu, mas sem 'Exportar dados'")

    print("-" * 100)
    print("Btn? = botão detectado | Menu? = menu do visual validado | Export? = suporta exportação")

async def async_input_with_timeout(prompt_text: str, timeout_seconds: int = 5, default_value: str = "todos") -> str:
    """
    Lê input do usuário com timeout.
    Se passar do tempo, devolve default_value.
    """
    print(prompt_text, end="", flush=True)

    loop = asyncio.get_running_loop()
    future = loop.run_in_executor(None, sys.stdin.readline)

    try:
        value = await asyncio.wait_for(future, timeout=timeout_seconds)
        value = (value or "").strip()
        if not value:
            log.info(f"⌛ Entrada vazia. Assumindo '{default_value}'.")
            return default_value
        return value
    except asyncio.TimeoutError:
        future.cancel()
        print("")  # quebra linha no terminal
        log.info(f"⌛ Sem resposta em {timeout_seconds}s. Assumindo '{default_value}'.")
        return default_value
    except Exception:
        future.cancel()
        log.warning(f"⚠️ Falha ao ler input. Assumindo '{default_value}'.")
        return default_value

async def ask_user_visual_selection(visuals):
    """
    Pergunta quais visuais exportar.
    Se o usuário demorar mais de USER_INPUT_TIMEOUT, exporta todos.
    """
    exportable = [i for i, v in enumerate(visuals) if v.get("hasExportData")]

    print("\n📥 Quais visuais deseja exportar?")
    print("   Opções:")
    print("   • Digite os números separados por vírgula: 0,2,5")
    print("   • Digite 'todos' para exportar todos")
    print("   • Digite 'sair' para cancelar")
    print(f"   • Se não responder em {USER_INPUT_TIMEOUT}s, o script exporta todos")

    choice = await async_input_with_timeout("   Sua escolha: ", USER_INPUT_TIMEOUT, "todos")
    choice_lower = choice.strip().lower()

    if choice_lower == "sair":
        log.info("🛑 Usuário escolheu 'sair'.")
        return []

    if choice_lower == "todos":
        log.info("✅ Seleção: exportar todos os visuais exportáveis.")
        return exportable

    selected = []
    for part in choice.split(","):
        part = part.strip()
        if not part:
            continue
        if part.isdigit():
            idx = int(part)
            if idx in exportable:
                selected.append(idx)

    # remove duplicados mantendo ordem
    final_selected = []
    seen = set()
    for idx in selected:
        if idx not in seen:
            seen.add(idx)
            final_selected.append(idx)

    if not final_selected:
        log.warning("⚠️ Nenhum índice válido informado. Exportando todos os visuais exportáveis.")
        return exportable

    log.info(f"✅ Seleção manual recebida: {final_selected}")
    return final_selected


# ---------------------------------------------------------------------------
# Exportação de um visual específico
# ---------------------------------------------------------------------------

async def click_export_data_menuitem(tab) -> bool:
    """
    Clica no item 'Exportar dados' do menu atualmente aberto.
    Ignora links do Microsoft Learn e elementos residuais.
    """
    try:
        result = await tab.evaluate("""
            (() => {
                const textOf = (el) => (el?.innerText || el?.textContent || '').trim();
                const lower = (s) => (s || '').toLowerCase();
                const isVisible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };

                const roots = Array.from(document.querySelectorAll([
                    '[role="menu"]',
                    '.contextMenu',
                    '.menu',
                    '.dropdown-menu'
                ].join(','))).filter(isVisible);

                for (const root of roots) {
                    const items = root.querySelectorAll('button, [role="menuitem"], a');
                    for (const el of items) {
                        if (!isVisible(el)) continue;

                        const txt = lower(textOf(el));
                        const href = lower(el.getAttribute('href') || '');

                        if (
                            txt.includes('saiba mais sobre como exportar dados') ||
                            txt.includes('learn more about exporting data') ||
                            href.includes('learn.microsoft.com') ||
                            href.includes('microsoft.com')
                        ) {
                            continue;
                        }

                        if (txt === 'exportar dados' || txt.includes('exportar dados')) {
                            try { el.click(); } catch (e) {}
                            return true;
                        }
                    }
                }
                return false;
            })()
        """)
        return bool(result)
    except Exception:
        return False


async def wait_export_dialog(tab, retries: int = 8) -> bool:
    """
    Aguarda o diálogo de exportação aparecer.
    """
    for _ in range(retries):
        try:
            exists = await tab.evaluate("""
                (() => {
                    const isVisible = (el) => {
                        const r = el.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    };

                    const candidates = Array.from(document.querySelectorAll([
                        '[role="dialog"]',
                        '[aria-modal="true"]',
                        '.modal',
                        '.popup',
                        '.dialog'
                    ].join(','))).filter(isVisible);

                    for (const dlg of candidates) {
                        const txt = (dlg.innerText || dlg.textContent || '').toLowerCase();
                        if (
                            txt.includes('exportar') ||
                            txt.includes('.xlsx') ||
                            txt.includes('dados resumidos') ||
                            txt.includes('dados subjacentes') ||
                            txt.includes('data with current layout')
                        ) {
                            return true;
                        }
                    }
                    return false;
                })()
            """)
            if exists:
                return True
        except Exception:
            pass
        await asyncio.sleep(1)
    return False


async def get_visible_dialog_snapshot(tab) -> list[str]:
    """Diagnóstico: lista textos resumidos dos dialogs visíveis."""
    try:
        raw = await tab.evaluate("""
            (() => {
                const isVisible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };
                const dialogs = Array.from(document.querySelectorAll(
                    '[role="dialog"], [aria-modal="true"], .modal, .popup, .dialog'
                )).filter(isVisible);
                const items = dialogs.map(d => (d.innerText || d.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 140));
                return JSON.stringify(items);
            })()
        """)
        parsed = json.loads(str(raw)) if raw else []
        return [str(x) for x in parsed]
    except Exception:
        return []

async def select_export_type(tab) -> bool:
    """
    Seleciona "Dados Resumidos" no diálogo de exportação do Power BI.
    Cobre mat-radio-button, input[type=radio] e label.
    """
    log.info("  🔘 Selecionando tipo de exportação (Dados Resumidos)...")

    try:
        result = await tab.evaluate("""
            (() => {
                const textOf = (el) => (el?.innerText || el?.textContent || '').trim().toLowerCase();
                const isVisible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };
                const DIALOG_SEL = [
                    'mat-dialog-container', 'export-data-dialog',
                    '[role="dialog"]', '[aria-modal="true"]',
                    '.cdk-overlay-pane', '.modal', '.popup', '.dialog'
                ].join(',');
                const dialogs = Array.from(document.querySelectorAll(DIALOG_SEL)).filter(isVisible);
                const RADIO_SEL = [
                    'mat-radio-button', '[role="radio"]',
                    'input[type="radio"]', 'label', '[role="option"]', 'mat-list-option'
                ].join(',');
                for (const dlg of dialogs) {
                    const txt_dlg = textOf(dlg);
                    if (!txt_dlg.includes('exportar') && !txt_dlg.includes('export') &&
                        !txt_dlg.includes('quais dados') && !txt_dlg.includes('which data')) continue;
                    const candidates = Array.from(dlg.querySelectorAll(RADIO_SEL)).filter(isVisible);
                    for (const el of candidates) {
                        const t = textOf(el);
                        if (t.includes('dados resumidos') || t.includes('summarized')) {
                            try { el.click(); } catch(e) {}
                            return 'dados_resumidos';
                        }
                    }
                    for (const el of candidates) {
                        const t = textOf(el);
                        if (t.includes('.xlsx') || t.includes('excel')) {
                            try { el.click(); } catch(e) {}
                            return 'xlsx_fallback';
                        }
                    }
                    const anyRadio = Array.from(dlg.querySelectorAll(
                        'mat-radio-button, [role="radio"], input[type="radio"]'
                    )).filter(isVisible);
                    if (anyRadio.length > 0) {
                        try { anyRadio[0].click(); } catch(e) {}
                        return 'primeiro_radio_mat';
                    }
                }
                return '';
            })()
        """)
    except Exception:
        result = ''

    if result:
        log.info(f"  ✅ Tipo de exportação selecionado: {result}")
        await asyncio.sleep(SHORT_WAIT)
        return True

    for js_expr in [
        'document.querySelector("mat-radio-button")',
        'document.querySelector("export-data-dialog mat-radio-button")',
        'document.querySelector("#pbi-radio-button-1 > label > section > div")',
        'document.querySelector("#pbi-radio-button-1 label")',
        'document.querySelector("#pbi-radio-button-1")',
        'document.querySelector("input[type=radio]:not([disabled])")',
    ]:
        if await js_click(tab, js_expr, "Radio button (fallback)"):
            await asyncio.sleep(SHORT_WAIT)
            return True

    log.warning("  ⚠️ Não foi possível selecionar o tipo de exportação")
    return False

async def confirm_export_dialog(tab) -> bool:
    """
    Clica no botão final 'Exportar' do diálogo de exportação do Power BI.
    Cobre mat-dialog-container, export-data-dialog e role=dialog.
    """
    log.info("  📤 Confirmando exportação...")

    try:
        result = await tab.evaluate("""
            (() => {
                const textOf = (el) => (el?.innerText || el?.textContent || '').trim().toLowerCase();
                const isVisible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };
                const DIALOG_SEL = [
                    'mat-dialog-container', 'export-data-dialog',
                    '[role="dialog"]', '[aria-modal="true"]',
                    '.cdk-overlay-pane', '.modal', '.popup', '.dialog'
                ].join(',');
                const dialogs = Array.from(document.querySelectorAll(DIALOG_SEL)).filter(isVisible);
                for (const dlg of dialogs) {
                    const dlgTxt = textOf(dlg);
                    if (!dlgTxt.includes('exportar') && !dlgTxt.includes('export') &&
                        !dlgTxt.includes('quais dados') && !dlgTxt.includes('which data')) continue;
                    const buttons = Array.from(dlg.querySelectorAll(
                        'button, [role="button"], mat-button, .mat-button'
                    )).filter(isVisible);
                    for (const el of buttons) {
                        const txt = textOf(el);
                        if (txt.includes('saiba mais') || txt.includes('learn more') ||
                            txt === 'cancelar' || txt === 'cancel') continue;
                        if (txt === 'exportar' || txt.includes('exportar') ||
                            txt === 'export' || txt.includes('export')) {
                            try { el.click(); } catch(e) {}
                            return true;
                        }
                    }
                    const actions = dlg.querySelector('mat-dialog-actions, .mat-dialog-actions');
                    if (actions) {
                        const btns = Array.from(actions.querySelectorAll('button')).filter(isVisible);
                        if (btns.length > 0) {
                            try { btns[0].click(); } catch(e) {}
                            return true;
                        }
                    }
                }
                return false;
            })()
        """)
        ok = bool(result)
    except Exception:
        ok = False

    if ok:
        log.info("  ✅ Botão Exportar clicado")
        await asyncio.sleep(1)
        return True

    for js_expr in [
        'document.querySelector("export-data-dialog button.exportButton")',
        'document.querySelector("export-data-dialog mat-dialog-actions button")',
        'document.querySelector("button.exportButton")',
        'document.querySelector("button.primaryBtn.exportButton")',
        'document.querySelector("mat-dialog-actions button.primaryBtn")',
        'document.querySelector("mat-dialog-actions button:first-child")',
    ]:
        if await js_click(tab, js_expr, "Botão Exportar (fallback)"):
            await asyncio.sleep(1)
            return True

    if await js_click_xpath(
        tab,
        '//*[@id="mat-mdc-dialog-0"]/div/div/export-data-dialog/mat-dialog-actions/button[1]',
        "Exportar (XPath fallback)"
    ):
        await asyncio.sleep(1)
        return True

    log.warning("  ⚠️ Botão final 'Exportar' não encontrado")
    return False

async def export_single_visual(tab, visual) -> bool:
    """
    Exporta um único visual.
    """
    idx = visual.get("index")
    title = visual.get("title", f"Visual #{idx}")

    await block_microsoft_learn_and_external_links(tab)
    await cleanup_residual_ui(tab, stage_label=f"início da exportação do visual #{idx}", aggressive=True)
    await dismiss_sensitive_data_popup(tab)

    log.info(f"  🖱️  Preparando visual #{idx}...")
    await scroll_visual_into_view(tab, visual)
    await hover_visual_center(tab, visual)
    await force_hover_visual_header(tab, visual)

    opened = await try_open_visual_menu_and_confirm_export(tab, visual, retries=MORE_OPTIONS_RETRIES)
    if not opened:
        log.warning(f"  ⚠️ Botão 'Mais opções' não encontrado para visual #{idx}")
        await cleanup_residual_ui(tab, stage_label=f"falha ao abrir menu do visual #{idx}", aggressive=True)
        return False

    clicked_export = False
    for export_click_attempt in range(1, 3):
        clicked_export = await click_export_data_menuitem(tab)
        if clicked_export:
            break
        log.info(f"  ⚠️ Tentativa {export_click_attempt}/2 sem clicar em 'Exportar dados'. Reabrindo menu...")
        await cleanup_residual_ui(tab, stage_label=f"retry menu export visual #{idx}", aggressive=True)
        reopened = await open_more_options_robust(tab, visual, export_click_attempt, 2)
        if not reopened:
            continue

    if not clicked_export:
        log.warning(f"  ⚠️ 'Exportar dados' não encontrada para visual #{idx}")
        await cleanup_residual_ui(tab, stage_label=f"falha em 'Exportar dados' do visual #{idx}", aggressive=True)
        return False

    log.info("  ✅ Menu acionado: Exportar dados")
    await asyncio.sleep(MEDIUM_WAIT)
    tab = await ensure_report_tab_still_valid(tab, POWERBI_URL)
    dialogs_after_export_click = await get_visible_dialog_snapshot(tab)
    if dialogs_after_export_click:
        log.info(f"  🧪 Diálogos visíveis após 'Exportar dados': {dialogs_after_export_click[:3]}")

    confirmed = False
    for dialog_attempt in range(1, 3):
        dialog_ready = await wait_export_dialog(tab, retries=6 if dialog_attempt == 1 else 4)
        if not dialog_ready:
            log.info(f"  ⚠️ Tentativa {dialog_attempt}/2 sem diálogo de exportação pronto.")
            snapshot = await get_visible_dialog_snapshot(tab)
            if snapshot:
                log.info(f"  🧪 Diálogo(s) ainda visíveis na tentativa {dialog_attempt}: {snapshot[:3]}")
            continue

        try:
            diag_html = await tab.evaluate("""
                (() => {
                    const isVisible = (el) => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
                    const DIALOG_SEL = ['mat-dialog-container','export-data-dialog','[role="dialog"]','[aria-modal="true"]','.cdk-overlay-pane'].join(',');
                    const dlgs = Array.from(document.querySelectorAll(DIALOG_SEL)).filter(isVisible);
                    for (const d of dlgs) {
                        const txt = (d.innerText||d.textContent||'').toLowerCase();
                        if (txt.includes('exportar') || txt.includes('quais dados')) {
                            return d.innerHTML.substring(0, 1500);
                        }
                    }
                    return 'no_export_dialog_found';
                })()
            """)
            log.info(f"  🧪 [DOM diálogo exportação]: {str(diag_html)[:800]}")
        except Exception as diag_e:
            log.info(f"  🧪 [DOM diagnóstico falhou]: {diag_e}")

        selected = await select_export_type(tab)
        log.info(f"  🔘 select_export_type retornou: {selected}")

        confirmed = await confirm_export_dialog(tab)
        if not confirmed:
            await asyncio.sleep(1)
            confirmed = await confirm_export_dialog(tab)

        if confirmed:
            break

    if not confirmed:
        await cleanup_residual_ui(tab, stage_label=f"falha ao confirmar exportação do visual #{idx}", aggressive=True)
        return False

    log.info(f"  ⏳ Aguardando download ({DOWNLOAD_WAIT}s)...")
    await asyncio.sleep(3)
    await accept_download_permission(tab)
    await asyncio.sleep(max(1, DOWNLOAD_WAIT - 3))

    await dismiss_sensitive_data_popup(tab)
    await cleanup_residual_ui(tab, stage_label=f"após exportar visual #{idx}", aggressive=True)

    log.info(f"  🎉 Visual #{idx} exportado com sucesso!")
    return True

async def export_selected_visuals(tab, visuals, selected_indexes):
    """
    Exporta os visuais escolhidos pelo usuário.
    """
    results = []

    valid_indexes = [i for i in selected_indexes if 0 <= i < len(visuals)]
    log.info("======================================================================")
    log.info("📌 Iniciando exportação")
    log.info("======================================================================")

    for pos, i in enumerate(valid_indexes, start=1):
        visual = visuals[i]
        title = visual.get("title", f"Visual #{i}")
        idx = visual.get("index", i)

        log.info(f"📦 [{pos}/{len(valid_indexes)}] Exportando: {title} (container #{idx})")

        ok = await export_single_visual(tab, visual)
        results.append({
            "visual_list_index": i,
            "container_index": idx,
            "title": title,
            "success": ok,
        })

        await close_open_menus_and_overlays(tab, aggressive=True)
        await dismiss_sensitive_data_popup(tab)
        await asyncio.sleep(SHORT_WAIT)

    return results

def display_export_summary(results):
    """
    Exibe resumo final da exportação.
    """
    print("\n" + "=" * 70)
    print("📊 RESUMO DA EXPORTAÇÃO")
    print("=" * 70)

    success_count = 0
    fail_count = 0

    for item in results:
        icon = "✅" if item.get("success") else "❌"
        print(f"  {icon} {item.get('title', 'Visual')}")

        if item.get("success"):
            success_count += 1
        else:
            fail_count += 1

    print(f"\n  Total: {success_count} sucesso, {fail_count} falha")
    print("  📂 Arquivos salvos na pasta Downloads padrão")

# ---------------------------------------------------------------------------
# Aceitar aviso de dados sensíveis
# ---------------------------------------------------------------------------

async def dismiss_sensitive_data_warning(tab):
    """
    Trata avisos de confidencialidade/exportação do Power BI.
    """
    handled_any = False

    for _ in range(6):
        result = await tab.evaluate(r"""
            (() => {
                const lowerText = el => (el?.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();
                const isVisible = el => {
                    if (!el) return false;
                    const r = el.getBoundingClientRect();
                    return r.width > 8 && r.height > 8;
                };

                const dialogSelectors = [
                    'mat-dialog-container', '.cdk-overlay-pane', '[role="dialog"]',
                    '[role="alertdialog"]', '.modal', '.pbi-modal',
                    '[class*="dialog"]', '[class*="Dialog"]',
                    '[class*="popup"]', '[class*="Popup"]',
                    '[class*="overlay"]', '[class*="Overlay"]'
                ];

                const actionWords = [
                    'copiar', 'copy', 'entendi', 'continuar', 'continue', 'aceitar',
                    'accept', 'concordo', 'agree', 'confirmar', 'confirm',
                    'ok', 'exportar', 'export', 'permitir', 'allow', 'prosseguir'
                ];

                const contextWords = [
                    'confid', 'sensív', 'sensit', 'privac', 'warning', 'aviso',
                    'dados subjacentes', 'underlying data', 'relatório', 'report',
                    'copiar relatório', 'copy report', 'exportação', 'export'
                ];

                for (const sel of dialogSelectors) {
                    const dialogs = document.querySelectorAll(sel);
                    for (const dialog of dialogs) {
                        if (!isVisible(dialog)) continue;
                        const text = lowerText(dialog);
                        const looksRelevant = contextWords.some(w => text.includes(w));
                        if (!looksRelevant) continue;

                        const buttons = dialog.querySelectorAll('button, [role="button"], a');
                        for (const btn of buttons) {
                            if (!isVisible(btn)) continue;
                            const textBtn = lowerText(btn);
                            if (actionWords.some(w => textBtn.includes(w))) {
                                btn.click();
                                return JSON.stringify({action: 'button', label: btn.textContent.trim().substring(0, 40)});
                            }
                        }

                        const closeBtn = dialog.querySelector(
                            'button[aria-label*="Close"], button[aria-label*="Fechar"], ' +
                            'button[title*="Close"], button[title*="Fechar"], ' +
                            '.close, .dialog-close, .popup-close, .ms-Dialog-button'
                        );
                        if (isVisible(closeBtn)) {
                            closeBtn.click();
                            return JSON.stringify({action: 'close-x', label: 'X'});
                        }
                    }
                }

                return JSON.stringify(null);
            })()
        """)

        try:
            parsed = json.loads(str(result)) if result is not None else None
        except Exception:
            parsed = None

        if parsed:
            handled_any = True
            log.info(f"  ⚠️🔓 Tratado popup: {parsed.get('action')} / {parsed.get('label')}")
            await asyncio.sleep(1.4)
            continue

        if handled_any:
            break
        await asyncio.sleep(0.8)

    return handled_any


# ---------------------------------------------------------------------------
# CDP helpers
# ---------------------------------------------------------------------------

async def _cdp_click(tab, x: int, y: int) -> bool:
    """
    Clique real via CDP.
    """
    try:
        await tab.send(cdp_dict_event("mouseMoved", x, y))
        await asyncio.sleep(0.05)
        await tab.send(cdp_dict_event("mousePressed", x, y))
        await asyncio.sleep(0.05)
        await tab.send(cdp_dict_event("mouseReleased", x, y))
        return True
    except Exception:
        pass

    try:
        await tab.mouse_move(x, y)
        await asyncio.sleep(0.05)
        await tab.mouse_click(x, y)
        return True
    except Exception:
        pass

    try:
        await tab.evaluate(f"""
            (() => {{
                const el = document.elementFromPoint({x}, {y});
                if (!el) return false;
                const cx = {x}, cy = {y};
                const init = {{bubbles:true, cancelable:true, composed:true,
                               clientX:cx, clientY:cy, screenX:cx, screenY:cy,
                               view:window}};
                el.dispatchEvent(new PointerEvent('pointerover',  init));
                el.dispatchEvent(new PointerEvent('pointerenter', init));
                el.dispatchEvent(new MouseEvent('mouseover',  init));
                el.dispatchEvent(new MouseEvent('mouseenter', init));
                el.dispatchEvent(new MouseEvent('mousemove',  init));
                el.dispatchEvent(new PointerEvent('pointerdown', init));
                el.dispatchEvent(new MouseEvent('mousedown',  init));
                el.dispatchEvent(new PointerEvent('pointerup',   init));
                el.dispatchEvent(new MouseEvent('mouseup',    init));
                el.dispatchEvent(new MouseEvent('click',      init));
                if (el.focus) el.focus();
                return true;
            }})()
        """)
        return True
    except Exception:
        return False


def cdp_dict_event(event_type: str, x: int, y: int) -> dict:
    button = "left" if event_type != "mouseMoved" else "none"
    click_count = 1 if event_type == "mousePressed" else 0
    return {
        "method": "Input.dispatchMouseEvent",
        "params": {
            "type": event_type,
            "x": x,
            "y": y,
            "button": button,
            "clickCount": click_count,
            "modifiers": 0,
            "deltaX": 0,
            "deltaY": 0,
        }
    }


def cdp_dict_key_event(event_type: str, key_payload: dict) -> dict:
    params = {"type": event_type}
    params.update(key_payload or {})
    return {
        "method": "Input.dispatchKeyEvent",
        "params": params,
    }


def _key_payload(key_name: str) -> dict:
    mapping = {
        "Enter": {
            "key": "Enter",
            "code": "Enter",
            "text": "\r",
            "unmodified_text": "\r",
            "windows_virtual_key_code": 13,
            "native_virtual_key_code": 13,
        },
        "ArrowDown": {
            "key": "ArrowDown",
            "code": "ArrowDown",
            "text": "",
            "unmodified_text": "",
            "windows_virtual_key_code": 40,
            "native_virtual_key_code": 40,
        },
        "ArrowUp": {
            "key": "ArrowUp",
            "code": "ArrowUp",
            "text": "",
            "unmodified_text": "",
            "windows_virtual_key_code": 38,
            "native_virtual_key_code": 38,
        },
        "Tab": {
            "key": "Tab",
            "code": "Tab",
            "text": "\t",
            "unmodified_text": "\t",
            "windows_virtual_key_code": 9,
            "native_virtual_key_code": 9,
        },
    }
    return dict(mapping.get(key_name, {}))


def _key_payload_raw_from_typed(typed_payload: dict) -> dict:
    return {
        "key": typed_payload.get("key", ""),
        "code": typed_payload.get("code", ""),
        "text": typed_payload.get("text", ""),
        "unmodifiedText": typed_payload.get("unmodified_text", ""),
        "windowsVirtualKeyCode": typed_payload.get("windows_virtual_key_code", 0),
        "nativeVirtualKeyCode": typed_payload.get("native_virtual_key_code", 0),
    }


async def _send_raw_cdp(tab, method: str, params: dict) -> bool:
    msg = {"method": method, "params": params or {}}
    attempts = []

    if hasattr(tab, "send"):
        attempts.append(("tab.send(dict)", tab.send, msg))

    conn = getattr(tab, "connection", None) or getattr(tab, "_connection", None)
    if conn and hasattr(conn, "send"):
        attempts.append(("tab.connection.send", conn.send, method, params))

    browser = getattr(tab, "browser", None)
    bconn = getattr(browser, "connection", None) if browser else None
    if bconn and hasattr(bconn, "send"):
        attempts.append(("browser.connection.send", bconn.send, method, params))

    for item in attempts:
        try:
            fn = item[1]
            args = item[2:]
            out = fn(*args)
            if asyncio.iscoroutine(out):
                await out
            return True
        except Exception:
            continue
    return False


async def _cdp_key_event(tab, key_name: str) -> dict:
    payload = _key_payload(key_name)
    if not payload:
        log.info(f"    ❌ [_cdp_key_event] tecla não suportada: {key_name}")
        return {"cdp_ok": False, "typed_api_ok": False, "raw_api_ok": False}

    typed_ok = False
    try:
        await tab.send(uc.cdp.input_.dispatch_key_event(type_="keyDown", **payload))
        if payload.get("text"):
            await tab.send(uc.cdp.input_.dispatch_key_event(type_="char", **payload))
        await tab.send(uc.cdp.input_.dispatch_key_event(type_="keyUp", **payload))
        typed_ok = True
    except Exception:
        pass

    raw_ok = False
    raw_payload = _key_payload_raw_from_typed(payload)
    try:
        down_ok = await _send_raw_cdp(tab, "Input.dispatchKeyEvent", {"type": "keyDown", **raw_payload})
        char_ok = True
        if raw_payload.get("text"):
            char_ok = await _send_raw_cdp(tab, "Input.dispatchKeyEvent", {"type": "char", **raw_payload})
        up_ok = await _send_raw_cdp(tab, "Input.dispatchKeyEvent", {"type": "keyUp", **raw_payload})
        raw_ok = bool(down_ok and char_ok and up_ok)
    except Exception:
        pass

    return {
        "cdp_ok": bool(typed_ok or raw_ok),
        "typed_api_ok": bool(typed_ok),
        "raw_api_ok": bool(raw_ok),
    }


async def _simple_keyboard_probe(tab) -> dict:
    result = {"ok": False, "enter_ok": False, "arrow_ok": False}
    try:
        await tab.evaluate("""
            (() => {
                const id = '__pbi_keyboard_probe__';
                let wrap = document.getElementById(id);
                if (!wrap) {
                    wrap = document.createElement('div');
                    wrap.id = id;
                    wrap.style.cssText = 'position:fixed;left:8px;bottom:8px;z-index:2147483647;background:#fff;padding:4px;border:1px solid #999;';
                    wrap.innerHTML = '<input id="__pbi_keyboard_probe_input__" style="width:180px" value="" />';
                    document.body.appendChild(wrap);
                }
            })()
        """)
        await tab.evaluate("document.getElementById('__pbi_keyboard_probe_input__')?.focus()")
        await asyncio.sleep(0.1)

        enter_diag = await _cdp_key_event(tab, "Enter")
        await asyncio.sleep(0.1)
        arrow_diag = await _cdp_key_event(tab, "ArrowDown")
        await asyncio.sleep(0.1)

        probe_raw = await tab.evaluate("""
            (() => {
                const i = document.getElementById('__pbi_keyboard_probe_input__');
                if (!i) return JSON.stringify({exists:false});
                return JSON.stringify({
                    exists: true,
                    focused: document.activeElement === i,
                    value: i.value || ''
                });
            })()
        """)
        probe_data = json.loads(str(probe_raw)) if probe_raw else {}
        result = {
            "ok": bool(probe_data.get("exists")),
            "enter_ok": bool((enter_diag or {}).get("cdp_ok")),
            "arrow_ok": bool((arrow_diag or {}).get("cdp_ok")),
            "focused": probe_data.get("focused"),
        }
    except Exception as e:
        result["error"] = str(e)

    # CLEANUP: remove o probe input e desfoca para não interferir nos slicers
    try:
        await tab.evaluate("""
            (() => {
                const wrap = document.getElementById('__pbi_keyboard_probe__');
                if (wrap) wrap.remove();
                document.activeElement?.blur?.();
            })()
        """)
    except Exception:
        pass

    return result


# ---------------------------------------------------------------------------
# Slicer activation experiment
# ---------------------------------------------------------------------------

async def _experiment_activate_slicer(tab, idx: int, slicer_title: str) -> dict:
    import json as _j

    async def _snapshot(label: str) -> dict:
        try:
            raw = await tab.evaluate(f"""
                (() => {{
                    const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                    if (!vc) return JSON.stringify({{error: 'no_vc'}});
                    const ae = document.activeElement;
                    const host = vc.closest('.visualContainerHost') || vc.parentElement;
                    const hostClass = host ? host.className : '';
                    const inp = vc.querySelector('input[type="text"], input[class*="search"], input[class*="Search"]');
                    const inpRect = inp ? inp.getBoundingClientRect() : null;
                    const inpVisible = !!(inpRect && inpRect.width > 0 && inpRect.height > 0);
                    const listbox = vc.querySelector('[role="listbox"], [role="option"], .dropdown-content, [class*="dropdown"], [class*="listbox"]');
                    const lbRect = listbox ? listbox.getBoundingClientRect() : null;
                    const lbVisible = !!(lbRect && lbRect.width > 0 && lbRect.height > 0);
                    const items = Array.from(vc.querySelectorAll('.slicerItemContainer, div.row, [class*="slicerItem"]')).filter(e => e.getBoundingClientRect().height > 0);
                    const vcClass = vc.className;
                    return JSON.stringify({{
                        host_class: hostClass, vc_class: vcClass,
                        ae_tag: ae ? ae.tagName : 'null',
                        ae_in_slicer: vc.contains(ae),
                        inp_visible: inpVisible, lb_visible: lbVisible,
                        item_count: items.length,
                    }});
                }})()
            """)
            data = _j.loads(str(raw)) if raw else {}
            data["_label"] = label
            return data
        except Exception as e:
            return {"error": str(e), "_label": label}

    async def _log_snapshot(s: dict):
        log.info(f"      host_class='{s.get('host_class','?')}' vc_class='{s.get('vc_class','?')}'")
        log.info(f"      ae={s.get('ae_tag','?')} in_slicer={s.get('ae_in_slicer','?')} inp_visible={s.get('inp_visible','?')} lb_visible={s.get('lb_visible','?')} items={s.get('item_count','?')}")

    def _is_activated(before: dict, after: dict) -> bool:
        host_changed = before.get("host_class") != after.get("host_class")
        focus_gained = (not before.get("ae_in_slicer")) and after.get("ae_in_slicer")
        inp_appeared = (not before.get("inp_visible")) and after.get("inp_visible")
        lb_appeared = (not before.get("lb_visible")) and after.get("lb_visible")
        items_grew = (after.get("item_count", 0) or 0) > (before.get("item_count", 0) or 0)
        return any([host_changed, focus_gained, inp_appeared, lb_appeared, items_grew])

    async def _get_targets() -> dict:
        try:
            raw = await tab.evaluate(f"""
                (() => {{
                    const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                    if (!vc) return JSON.stringify({{error:'no_vc'}});
                    const r = (el) => {{
                        if (!el) return null;
                        const rect = el.getBoundingClientRect();
                        if (!rect || rect.width <= 0) return null;
                        return {{ x: Math.round(rect.left + rect.width/2), y: Math.round(rect.top + rect.height/2) }};
                    }};
                    const a = r(vc);
                    const body = vc.querySelector('.slicer-content-wrapper, .slicer-body, .slicerBody, [class*="slicerBody"], [class*="content"]');
                    const b = r(body);
                    const transform = vc.querySelector('transform');
                    const c = r(transform ? transform.firstElementChild : null);
                    const header = vc.querySelector('visual-container-header button, visual-header-item-container button, .visualContainerHeader button');
                    const d = r(header);
                    return JSON.stringify({{a, b, c, d}});
                }})()
            """)
            return _j.loads(str(raw)) if raw else {}
        except Exception:
            return {}

    log.info(f"  🔬 [{slicer_title}] Iniciando 4 experimentos de ativação...")
    targets = await _get_targets()
    log.info(f"    Alvos: A={targets.get('a')} B={targets.get('b')} C={targets.get('c')} D={targets.get('d')}")

    experiments = [
        ("A", "centro do visual-container", targets.get("a")),
        ("B", "área de conteúdo (body/items)", targets.get("b")),
        ("C", "wrapper intermediário (transform > div)", targets.get("c")),
        ("D", "botão do header", targets.get("d")),
    ]

    result = {"winner": None, "winner_target": None, "winner_coords": None, "evidence": {}}

    for letter, description, coords in experiments:
        log.info(f"    🔹 Tentativa {letter} — {description} coords={coords}")
        await press_escape(tab, times=1, wait_each=0.2)
        await asyncio.sleep(0.3)

        before = await _snapshot(f"{letter}_before")
        log.info(f"    ANTES [{letter}]:")
        await _log_snapshot(before)

        if not coords:
            log.info(f"    ⚠️ [{letter}] alvo não encontrado no DOM — pulando")
            continue

        cx, cy = coords["x"], coords["y"]
        if cx <= 0 or cy <= 0:
            log.info(f"    ⚠️ [{letter}] coordenada inválida ({cx},{cy}) — pulando")
            continue

        try:
            await tab.evaluate(f"""
                (() => {{
                    const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                    if (!vc) return;
                    const init = {{bubbles:true, composed:true, view:window, clientX:{cx}, clientY:{cy}}};
                    ['pointerenter','pointerover','mouseenter','mouseover','mousemove'].forEach(t =>
                        vc.dispatchEvent(new PointerEvent(t, init))
                    );
                }})()
            """)
        except Exception:
            pass
        await asyncio.sleep(0.5)

        clicked = await _cdp_click(tab, cx, cy)
        log.info(f"    [CLICK {letter}] cdp_click={clicked} coords=({cx},{cy})")
        await asyncio.sleep(1.2)

        after = await _snapshot(f"{letter}_after")
        log.info(f"    DEPOIS [{letter}]:")
        await _log_snapshot(after)

        activated = _is_activated(before, after)
        log.info(
            f"    {'✅ ATIVADO' if activated else '❌ sem mudança'} [{letter}] "
            f"host_changed={before.get('host_class') != after.get('host_class')} "
            f"focus_gained={after.get('ae_in_slicer')} "
            f"inp_appeared={after.get('inp_visible') and not before.get('inp_visible')} "
            f"items_before={before.get('item_count')} items_after={after.get('item_count')}"
        )

        result["evidence"][letter] = {
            "coords": coords, "clicked": clicked, "activated": activated,
            "before": before, "after": after,
        }

        if activated and result["winner"] is None:
            result["winner"] = letter
            result["winner_target"] = description
            result["winner_coords"] = coords
            log.info(f"    🏆 Tentativa {letter} é o VENCEDOR para este slicer")

    if result["winner"]:
        log.info(f"  ✅ Experimento concluído: vencedor={result['winner']} ({result['winner_target']})")
    else:
        log.info(f"  ❌ Experimento concluído: NENHUMA tentativa ativou o visual")

    return result


# ---------------------------------------------------------------------------
# NEW: Micro-scroll safe enumeration for slicers (v10.2)
# ---------------------------------------------------------------------------

# Textos de UI que devem ser descartados na enumeração
_SLICER_UI_NOISE = {
    'pressionar enter para explorar os dados', 'selecionar tudo', 'select all',
    'buscar', 'search', 'pesquisar', 'ainda não aplicado', 'not yet applied',
    'apply changes', 'aplicar alterações', '(ainda não aplicado)', 'basic',
    'limpar', 'clear', 'clear filter', 'limpar filtro',
}


def _is_slicer_noise(text: str, field_name: str) -> bool:
    """Retorna True se o texto é ruído de UI e não um valor real do slicer."""
    low = text.lower().strip()
    return (
        not low
        or low in _SLICER_UI_NOISE
        or low == field_name
        or low.startswith('pressionar enter')
        or (low.startswith('(') and low.endswith(')'))
        or len(text) > 80
    )


async def _click_slicer_header_safe(tab, idx: int) -> dict:
    """
    Clica na área do CABEÇALHO do slicer (texto do título, área acima dos valores).
    Isso é seguro porque não altera seleção de valores.
    """
    import json as _j
    try:
        raw = await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return JSON.stringify({{ok:false, reason:'no_vc'}});
                const vcRect = vc.getBoundingClientRect();

                // Estratégia 1: elemento de texto do cabeçalho do slicer
                const headerSels = [
                    '.slicer-header-text', '[class*="header-text"]', '[class*="headerText"]',
                    '.visual-title', '.visualTitle',
                ];
                for (const sel of headerSels) {{
                    const el = vc.querySelector(sel);
                    if (el) {{
                        const r = el.getBoundingClientRect();
                        if (r.width > 3 && r.height > 3 && r.top >= vcRect.top - 5 && r.top < vcRect.top + 45) {{
                            return JSON.stringify({{
                                ok:true,
                                x: Math.round(r.left + r.width/2),
                                y: Math.round(r.top + r.height/2),
                                method:'header_text_el'
                            }});
                        }}
                    }}
                }}

                // Estratégia 2: filterRestatement
                const restatement = vc.querySelector('.filterRestatement, [class*="filterRestatement"]');
                if (restatement) {{
                    const r = restatement.getBoundingClientRect();
                    if (r.width > 3 && r.height > 3) {{
                        return JSON.stringify({{ok:true, x: Math.round(r.left + r.width/2), y: Math.round(r.top + r.height/2), method:'restatement'}});
                    }}
                }}

                // Estratégia 3: faixa segura no topo do visual-container
                const firstItem = vc.querySelector('.slicerItemContainer, [class*="slicerItem"]');
                let safeY = Math.round(vcRect.top + 10);
                if (firstItem) {{
                    const itemRect = firstItem.getBoundingClientRect();
                    if (itemRect.top > vcRect.top + 5) {{
                        safeY = Math.round((vcRect.top + itemRect.top) / 2);
                    }}
                }}
                return JSON.stringify({{ok:true, x: Math.round(vcRect.left + vcRect.width / 2), y: safeY, method:'vc_top_safe_band'}});
            }})()
        """)
        return _j.loads(str(raw)) if raw else {"ok": False}
    except Exception as e:
        return {"ok": False, "error": str(e)}


async def _find_scroll_container(tab, idx: int, slicer_title: str) -> dict:
    """
    ETAPA 1: Identifica o container rolável real dentro do slicer.
    Loga todos os candidatos e escolhe o melhor.
    """
    import json as _j
    try:
        raw = await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return JSON.stringify({{found:false, reason:'no_vc', candidates:[]}});

                const ITEM_SEL = '.slicerItemContainer, [role="option"], [role="listitem"], div.row, [class*="slicerItem"]';
                const candidates = [];
                const allEls = Array.from(vc.querySelectorAll('*'));

                // Gera caminho DOM resumido
                const domPath = (el) => {{
                    const parts = [];
                    let cur = el;
                    for (let i = 0; i < 5 && cur && cur !== vc; i++) {{
                        const tag = cur.tagName.toLowerCase();
                        const cls = (cur.className || '').toString().split(' ').filter(Boolean).slice(0,2).join('.');
                        parts.unshift(cls ? tag + '.' + cls : tag);
                        cur = cur.parentElement;
                    }}
                    return parts.join(' > ');
                }};

                for (const el of allEls) {{
                    const cs = getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    if (r.width < 20 || r.height < 20) continue;

                    const items = el.querySelectorAll(ITEM_SEL);
                    const itemCount = Array.from(items).filter(i => i.getBoundingClientRect().height > 0).length;
                    if (itemCount < 1) continue;

                    const ov = (cs.overflow || '').toLowerCase();
                    const ovY = (cs.overflowY || '').toLowerCase();
                    const scrollable = el.scrollHeight > el.clientHeight + 2;
                    const hasOverflow = ovY.includes('auto') || ovY.includes('scroll') || ov.includes('auto') || ov.includes('scroll');

                    if (!scrollable && !hasOverflow && itemCount < 2) continue;

                    candidates.push({{
                        tag: el.tagName,
                        id: el.id || '',
                        class: (el.className || '').toString().slice(0, 120),
                        role: el.getAttribute('role') || '',
                        scrollTop: Math.round(el.scrollTop || 0),
                        scrollHeight: Math.round(el.scrollHeight || 0),
                        clientHeight: Math.round(el.clientHeight || 0),
                        offsetHeight: Math.round(el.offsetHeight || 0),
                        overflow: ov,
                        overflowY: ovY,
                        items_inside: itemCount,
                        dom_path: domPath(el),
                        _scrollable: scrollable,
                        _hasOverflow: hasOverflow,
                        _score: (scrollable ? 100 : 0) + (hasOverflow ? 50 : 0) + itemCount * 10 + (el.scrollHeight - el.clientHeight),
                    }});
                }}

                // Fallback: listbox
                if (candidates.length === 0) {{
                    const lb = vc.querySelector('[role="listbox"]');
                    if (lb) {{
                        const cs = getComputedStyle(lb);
                        const items = Array.from(lb.querySelectorAll(ITEM_SEL)).filter(i => i.getBoundingClientRect().height > 0);
                        candidates.push({{
                            tag: lb.tagName, id: lb.id || '',
                            class: (lb.className || '').toString().slice(0,120),
                            role: lb.getAttribute('role') || '',
                            scrollTop: Math.round(lb.scrollTop||0), scrollHeight: Math.round(lb.scrollHeight||0),
                            clientHeight: Math.round(lb.clientHeight||0), offsetHeight: Math.round(lb.offsetHeight||0),
                            overflow: (cs.overflow||''), overflowY: (cs.overflowY||''),
                            items_inside: items.length, dom_path: domPath(lb),
                            _scrollable: lb.scrollHeight > lb.clientHeight + 2, _hasOverflow: true,
                            _score: 200 + items.length * 10,
                        }});
                    }}
                }}

                candidates.sort((a,b) => b._score - a._score);
                const best = candidates[0] || null;

                return JSON.stringify({{
                    found: !!best,
                    reason: best ? 'selected' : 'no_candidates',
                    candidates: candidates.slice(0, 8),
                    selected: best,
                }});
            }})()
        """)
        data = _j.loads(str(raw)) if raw else {"found": False, "candidates": []}
    except Exception as e:
        data = {"found": False, "error": str(e), "candidates": []}

    # Log de cada candidato
    for ci, cand in enumerate(data.get("candidates", [])):
        log.info(f"    [SCROLL_CONTAINER_CANDIDATE]")
        log.info(f"    slicer={slicer_title}")
        log.info(f"    candidate_index={ci}")
        for k in ("tag", "id", "class", "role", "scrollTop", "scrollHeight", "clientHeight", "offsetHeight", "overflow", "overflowY", "items_inside", "dom_path"):
            log.info(f"    {k}={cand.get(k, '')}")

    sel = data.get("selected")
    if sel:
        log.info(f"    [SCROLL_CONTAINER_SELECTED]")
        log.info(f"    slicer={slicer_title}")
        for k in ("tag", "id", "class", "role"):
            log.info(f"    {k}={sel.get(k, '')}")
        log.info(f"    initial_scrollTop={sel.get('scrollTop')}")
        log.info(f"    initial_scrollHeight={sel.get('scrollHeight')}")
        log.info(f"    initial_clientHeight={sel.get('clientHeight')}")
        log.info(f"    selection_reason=highest_score({sel.get('_score',0)})")
    else:
        log.info(f"    [SCROLL_CONTAINER_SELECTED] slicer={slicer_title} — NENHUM encontrado")

    return data


async def _read_visible_box(tab, idx: int, field_name: str, step: int = -1) -> dict:
    """
    Lê a "caixa" atual de elementos visíveis no slicer.
    v10.4 — CORRIGIDO:
      - SELECTOR_TEST é a fonte oficial da caixa. Se encontrou textos, a caixa NÃO pode ficar vazia.
      - BOX_ITEM logado por elemento.
      - RAW_TEXT_CAPTURE construído diretamente dos elementos aceitos.
      - BOX_CONTRADICTION logado se SELECTOR_TEST encontrou textos mas raw_texts ficou vazio.
    """
    import json as _j
    slicer_label = f"container#{idx}"

    try:
        raw = await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return JSON.stringify({{error:'no_vc'}});

                // ── Re-encontrar o scroll container ──
                const allEls = Array.from(vc.querySelectorAll('*'));
                let scrollContainer = null;
                let bestScore = -1;
                for (const el of allEls) {{
                    const cs = getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    if (r.width < 20 || r.height < 20) continue;
                    const ov = (cs.overflowY || cs.overflow || '').toLowerCase();
                    const scrollable = el.scrollHeight > el.clientHeight + 2;
                    const hasOv = ov.includes('auto') || ov.includes('scroll');
                    if (!scrollable && !hasOv) continue;
                    const score = (scrollable ? 100 : 0) + (hasOv ? 50 : 0) + (el.scrollHeight - el.clientHeight);
                    if (score > bestScore) {{ bestScore = score; scrollContainer = el; }}
                }}
                if (!scrollContainer) scrollContainer = vc.querySelector('[role="listbox"]') || vc;

                const scTag = scrollContainer.tagName;
                const scClass = (scrollContainer.className || '').toString().slice(0, 120);
                const scScrollTop = Math.round(scrollContainer.scrollTop || 0);

                // ── DOM Snapshot ──
                const childNodes = Array.from(scrollContainer.children || []);
                const firstChildrenTags = childNodes.slice(0, 15).map(c => {{
                    const tag = c.tagName.toLowerCase();
                    const cls = (c.className || '').toString().split(' ').filter(Boolean).slice(0,2).join('.');
                    const txt = (c.textContent || '').trim().slice(0, 40);
                    return tag + (cls ? '.' + cls : '') + (txt ? '="' + txt + '"' : '');
                }});
                const innerTextSample = (scrollContainer.innerText || '').slice(0, 500);

                // ── SELECTOR_TEST: testar múltiplos seletores DENTRO do scrollContainer ──
                const testSelectors = [
                    '.slicerItemContainer',
                    '[role="option"]',
                    '[role="listitem"]',
                    '[class*="slicerItem"]',
                    'div.row',
                    '.slicer-text',
                    '.slicerText',
                    'span.slicerText',
                    'span[class*="slicerText"]',
                    'label',
                ];

                // Para cada seletor, coleta elementos COM bbox válida
                const selectorResults = [];
                for (const sel of testSelectors) {{
                    const found = Array.from(scrollContainer.querySelectorAll(sel)).filter(e => {{
                        const r = e.getBoundingClientRect();
                        return r.height > 0 && r.width > 0;
                    }});
                    // Extrai texto de cada elemento encontrado (mesma lógica usada depois)
                    const itemData = found.map(e => {{
                        const span = e.querySelector('.slicerText, span[class*="slicerText"], span');
                        const textContent = (e.textContent || '').replace(/\\s+/g, ' ').trim();
                        const innerText   = ((span || e).innerText || '').replace(/\\s+/g, ' ').trim();
                        const normalized  = textContent || innerText;
                        return {{
                            tag: e.tagName.toLowerCase(),
                            cls: (e.className || '').toString().slice(0, 60),
                            textContent,
                            innerText,
                            normalized,
                        }};
                    }});
                    const texts = itemData
                        .map(d => d.normalized)
                        .filter(t => t && t.length > 0 && t.length <= 80);
                    selectorResults.push({{
                        selector: sel,
                        elements_found: found.length,
                        texts,
                        item_data: itemData,
                    }});
                }}

                // ── Escolher o melhor seletor: mais textos não-vazios ──
                let bestSr = null;
                for (const sr of selectorResults) {{
                    if (sr.texts.length > 0) {{
                        if (!bestSr || sr.texts.length > bestSr.texts.length) {{
                            bestSr = sr;
                        }}
                    }}
                }}

                // ── Fallback AMPLO se nenhum seletor encontrou texto ──
                if (!bestSr || bestSr.texts.length === 0) {{
                    const fallbackItems = [];
                    const walk = (parent, depth) => {{
                        if (depth > 4) return;
                        for (const child of Array.from(parent.children || [])) {{
                            const r = child.getBoundingClientRect();
                            if (r.height <= 0 || r.width <= 0) continue;
                            const tc = (child.textContent || '').replace(/\\s+/g, ' ').trim();
                            if (tc && tc.length > 0 && tc.length <= 80) {{
                                fallbackItems.push({{
                                    tag: child.tagName.toLowerCase(),
                                    cls: (child.className || '').toString().slice(0,60),
                                    textContent: tc,
                                    innerText: (child.innerText || '').replace(/\\s+/g,' ').trim(),
                                    normalized: tc,
                                }});
                            }}
                            walk(child, depth + 1);
                        }}
                    }};
                    walk(scrollContainer, 0);
                    if (fallbackItems.length > 0) {{
                        const texts = [...new Set(fallbackItems.map(d => d.normalized).filter(t => t))];
                        bestSr = {{
                            selector: 'fallback_walk',
                            elements_found: fallbackItems.length,
                            texts,
                            item_data: fallbackItems,
                        }};
                    }}
                }}

                // ── Agora constrói raw_texts E box_items DIRETO do bestSr ──
                // Fonte única de verdade: item_data do seletor vencedor.
                const box_items = [];
                const raw_texts = [];
                const selected = [];
                const seen = new Set();

                if (bestSr && bestSr.item_data) {{
                    // Precisamos dos elementos reais para checar seleção;
                    // re-query APENAS o seletor vencedor (ou usa item_data para texto)
                    const winnerEls = bestSr.selector !== 'fallback_walk'
                        ? Array.from(scrollContainer.querySelectorAll(bestSr.selector)).filter(e => {{
                            const r = e.getBoundingClientRect();
                            return r.height > 0 && r.width > 0;
                          }})
                        : []; // fallback: sem re-query de elementos DOM

                    for (let i = 0; i < bestSr.item_data.length; i++) {{
                        const d = bestSr.item_data[i];
                        const normalized = d.normalized || '';
                        if (!normalized) {{
                            box_items.push({{ index:i, tag:d.tag, cls:d.cls,
                                textContent:d.textContent, innerText:d.innerText,
                                normalized:'', accepted:false, discard_reason:'empty_text' }});
                            continue;
                        }}
                        if (normalized.length > 80) {{
                            box_items.push({{ index:i, tag:d.tag, cls:d.cls,
                                textContent:d.textContent, innerText:d.innerText,
                                normalized, accepted:false, discard_reason:'too_long' }});
                            continue;
                        }}
                        if (seen.has(normalized)) {{
                            box_items.push({{ index:i, tag:d.tag, cls:d.cls,
                                textContent:d.textContent, innerText:d.innerText,
                                normalized, accepted:false, discard_reason:'duplicate' }});
                            continue;
                        }}
                        seen.add(normalized);
                        raw_texts.push(normalized);

                        // Checa seleção via elemento DOM real (se disponível)
                        let isSel = false;
                        if (winnerEls[i]) {{
                            const el = winnerEls[i];
                            const itemOrParent = el.closest('.slicerItemContainer, [role="option"], [role="listitem"]') || el;
                            isSel = (
                                itemOrParent.classList.contains('selected') ||
                                itemOrParent.classList.contains('isSelected') ||
                                itemOrParent.getAttribute('aria-selected') === 'true' ||
                                itemOrParent.getAttribute('aria-checked') === 'true' ||
                                !!itemOrParent.querySelector('.selected, .isSelected, .partiallySelected')
                            );
                        }}
                        if (isSel) selected.push(normalized);

                        box_items.push({{ index:i, tag:d.tag, cls:d.cls,
                            textContent:d.textContent, innerText:d.innerText,
                            normalized, accepted:true, discard_reason:'' }});
                    }}
                }}

                // node_ids para VIRTUALIZATION_CHECK
                const nodeIds = (bestSr && bestSr.selector !== 'fallback_walk')
                    ? Array.from(scrollContainer.querySelectorAll(bestSr.selector))
                        .filter(e => {{ const r = e.getBoundingClientRect(); return r.height > 0 && r.width > 0; }})
                        .slice(0, 10)
                        .map(e => (e.tagName||'') + '#' + (e.id||'') + '.' + ((e.className||'').toString().slice(0,30)))
                    : raw_texts.slice(0, 10); // fallback: usa textos como identidade

                return JSON.stringify({{
                    raw_texts,
                    selected,
                    sc_tag: scTag,
                    sc_class: scClass,
                    sc_scrollTop: scScrollTop,
                    dom_snapshot: {{
                        innerText_sample: innerTextSample,
                        child_count: childNodes.length,
                        first_children_tags: firstChildrenTags,
                    }},
                    selector_results: selectorResults.map(sr => ({{
                        selector: sr.selector,
                        elements_found: sr.elements_found,
                        texts: sr.texts,
                    }})),
                    best_selector: bestSr ? bestSr.selector : 'none',
                    best_selector_count: bestSr ? bestSr.elements_found : 0,
                    best_selector_texts: bestSr ? bestSr.texts : [],
                    box_items,
                    node_ids: nodeIds,
                }});
            }})()
        """)
        data = _j.loads(str(raw)) if raw else {"raw_texts": [], "selected": []}
    except Exception as e:
        log.info(f"    ⚠️ [_read_visible_box] evaluate error: {e}")
        data = {"raw_texts": [], "selected": [], "error": str(e)}

    raw_texts            = data.get("raw_texts", [])
    selected_raw         = data.get("selected", [])
    best_selector        = data.get("best_selector", "none")
    best_selector_texts  = data.get("best_selector_texts", [])
    box_items            = data.get("box_items", [])

    # ── Log DOM Snapshot ──
    dom_snap = data.get("dom_snapshot", {})
    if step >= 0 or not raw_texts:
        log.info(f"    [DOM_SNAPSHOT] slicer={slicer_label} step={step}")
        log.info(f"    sc_tag={data.get('sc_tag','')} sc_class={data.get('sc_class','')[:80]} sc_scrollTop={data.get('sc_scrollTop','')}")
        log.info(f"    innerText_sample={dom_snap.get('innerText_sample','')[:300]}")
        log.info(f"    child_count={dom_snap.get('child_count',0)}")
        log.info(f"    first_children_tags={dom_snap.get('first_children_tags',[])}")

    # ── Log SELECTOR_TEST (todos os seletores com elementos) ──
    for sr in data.get("selector_results", []):
        if sr.get("elements_found", 0) > 0:
            log.info(
                f"    [SELECTOR_TEST] selector={sr['selector']} "
                f"elements_found={sr['elements_found']} "
                f"texts={sr.get('texts',[])[:8]}"
            )

    # ── Log BOX_ITEM por elemento ──
    for bi in box_items:
        log.info(
            f"    [BOX_ITEM] step={step} selector={best_selector} index={bi.get('index','')} "
            f"tag={bi.get('tag','')} class={bi.get('cls','')[:40]} "
            f"textContent={bi.get('textContent','')[:60]} "
            f"innerText={bi.get('innerText','')[:60]} "
            f"normalized_text={bi.get('normalized','')[:60]} "
            f"accepted={bi.get('accepted',False)} "
            f"discard_reason={bi.get('discard_reason','')}"
        )

    # ── Log RAW_TEXT_CAPTURE (fonte: elementos aceitos do seletor vencedor) ──
    log.info(
        f"    [RAW_TEXT_CAPTURE] best_selector={best_selector} "
        f"count={data.get('best_selector_count',0)} "
        f"raw_texts={raw_texts[:15]}"
    )

    # ── BOX_CONTRADICTION: SELECTOR_TEST encontrou textos mas raw_texts está vazio ──
    if best_selector_texts and not raw_texts:
        log.warning(
            f"    [BOX_CONTRADICTION] selector={best_selector} "
            f"selector_texts={best_selector_texts[:10]} "
            f"raw_texts=[] "
            f"reason=selector_found_texts_but_box_pipeline_dropped_them"
        )
        # FALLBACK OFICIAL: usar os textos do SELECTOR_TEST como fonte de verdade
        log.info(
            f"    [BOX_CONTRADICTION_RECOVERY] usando best_selector_texts como fallback oficial"
        )
        raw_texts = list(best_selector_texts)

    # ── Filtrar ruído (TEXT_FILTER) ──
    accepted  = []
    discarded = []
    discard_reasons = {}
    for t in raw_texts:
        if _is_slicer_noise(t, field_name):
            discarded.append(t)
            low = t.lower().strip()
            reason = (
                "ui_noise"    if low in _SLICER_UI_NOISE else
                "field_name"  if low == field_name else
                "pattern"
            )
            discard_reasons[t] = reason
        else:
            accepted.append(t)

    # ── Log TEXT_FILTER ──
    log.info(f"    [TEXT_FILTER] step={step}")
    log.info(f"    raw_texts={raw_texts[:15]}")
    log.info(f"    discarded_texts={discarded}")
    log.info(f"    discard_reasons={discard_reasons}")
    log.info(f"    accepted_values={accepted}")

    # ── Log VISIBLE_BOX resumo ──
    log.info(
        f"    [VISIBLE_BOX] step={step} "
        f"raw_visible_elements_count={len(accepted)} "
        f"raw_texts={raw_texts[:15]} "
        f"filtered_candidate_values={accepted}"
    )

    return {
        "raw_texts"       : raw_texts,
        "accepted"        : accepted,
        "discarded"       : discarded,
        "discard_reasons" : discard_reasons,
        "selected"        : [s for s in selected_raw if not _is_slicer_noise(s, field_name)],
        "node_ids"        : data.get("node_ids", []),
        "best_selector"   : best_selector,

    }


async def _read_slicer_selected_snapshot(tab, idx: int) -> list:
    """Lê estado de seleção do slicer para guarda de seleção indevida.
    Busca dentro do scroll container E no visual-container para cobertura máxima."""
    import json as _j
    try:
        raw = await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return JSON.stringify([]);

                // Busca em TODO o visual-container (garante cobertura)
                const ITEM_SEL = '.slicerItemContainer, [role="option"], [role="listitem"], div.row, [class*="slicerItem"]';
                const out = [];
                const seen = new Set();

                // Busca por itens com atributo de seleção
                vc.querySelectorAll(ITEM_SEL).forEach(item => {{
                    const isSel = item.classList.contains('selected') ||
                        item.classList.contains('isSelected') ||
                        item.getAttribute('aria-selected') === 'true' ||
                        item.getAttribute('aria-checked') === 'true';
                    if (isSel) {{
                        const t = (item.textContent || '').replace(/\\s+/g, ' ').trim();
                        if (t && !seen.has(t)) {{ seen.add(t); out.push(t); }}
                    }}
                }});

                // Fallback: busca por classes .selected em qualquer descendente
                if (out.length === 0) {{
                    vc.querySelectorAll('.selected, .isSelected, [aria-selected="true"], [aria-checked="true"]').forEach(el => {{
                        const t = (el.textContent || '').replace(/\\s+/g, ' ').trim();
                        if (t && t.length < 80 && !seen.has(t)) {{ seen.add(t); out.push(t); }}
                    }});
                }}

                return JSON.stringify(out.slice(0, 80));
            }})()
        """)
        return _j.loads(str(raw)) if raw else []
    except Exception:
        return []


async def _micro_scroll_step(tab, idx: int, slicer_title: str, step: int, delta: int = 16) -> dict:
    """
    ETAPA 4: Aplica um micro-scroll incremental no container rolável do slicer.
    Tenta scrollTop direto primeiro, depois wheel como fallback.
    """
    import json as _j
    try:
        raw = await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return JSON.stringify({{ok:false, reason:'no_vc'}});

                const ITEM_SEL = '.slicerItemContainer, [role="option"], [role="listitem"], div.row, [class*="slicerItem"]';

                // Encontra o melhor container rolável (mesmo critério de _find_scroll_container)
                const candidates = Array.from(vc.querySelectorAll('*')).filter(el => {{
                    const cs = getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    if (r.width < 20 || r.height < 20) return false;
                    const items = Array.from(el.querySelectorAll(ITEM_SEL)).filter(i => i.getBoundingClientRect().height > 0);
                    if (items.length < 1) return false;
                    const ov = (cs.overflowY || cs.overflow || '').toLowerCase();
                    return el.scrollHeight > el.clientHeight + 2 || ov.includes('auto') || ov.includes('scroll');
                }});

                let target = null;
                if (candidates.length > 0) {{
                    candidates.sort((a,b) => (b.scrollHeight - b.clientHeight) - (a.scrollHeight - a.clientHeight));
                    target = candidates[0];
                }}
                if (!target) {{
                    target = vc.querySelector('[role="listbox"]');
                }}
                if (!target) return JSON.stringify({{ok:false, reason:'no_scroll_target'}});

                const before = target.scrollTop || 0;
                target.scrollTop = before + {delta};
                const after = target.scrollTop || 0;

                return JSON.stringify({{
                    ok: true,
                    scroll_method: 'scrollTop',
                    scrollTop_before: Math.round(before),
                    scrollTop_after: Math.round(after),
                    scroll_changed: Math.abs(after - before) > 0.5,
                    scrollHeight: Math.round(target.scrollHeight || 0),
                    clientHeight: Math.round(target.clientHeight || 0),
                }});
            }})()
        """)
        data = _j.loads(str(raw)) if raw else {"ok": False}
    except Exception as e:
        data = {"ok": False, "error": str(e)}

    # Se scrollTop não mudou, tenta wheel
    if data.get("ok") and not data.get("scroll_changed"):
        try:
            center_raw = await tab.evaluate(f"""
                (() => {{
                    const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                    if (!vc) return JSON.stringify({{ok:false}});
                    const r = vc.getBoundingClientRect();
                    return JSON.stringify({{ok:true, x:Math.round(r.left+r.width/2), y:Math.round(r.top+r.height/2)}});
                }})()
            """)
            center = _j.loads(str(center_raw)) if center_raw else {"ok": False}
            if center.get("ok"):
                try:
                    await tab.send(uc.cdp.input_.dispatch_mouse_event(
                        type_="mouseWheel",
                        x=center["x"], y=center["y"],
                        delta_x=0, delta_y=delta * 3,
                    ))
                    data["scroll_method"] = "wheel_fallback"
                    await asyncio.sleep(0.15)
                    # Re-check scrollTop
                    recheck_raw = await tab.evaluate(f"""
                        (() => {{
                            const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                            if (!vc) return JSON.stringify({{st:0}});
                            const ITEM_SEL = '.slicerItemContainer, [role="option"], [role="listitem"], div.row, [class*="slicerItem"]';
                            const cands = Array.from(vc.querySelectorAll('*')).filter(el => {{
                                const items = Array.from(el.querySelectorAll(ITEM_SEL)).filter(i => i.getBoundingClientRect().height > 0);
                                return items.length > 0 && el.scrollHeight > el.clientHeight + 2;
                            }});
                            const t = cands[0] || vc.querySelector('[role="listbox"]');
                            return JSON.stringify({{st: t ? Math.round(t.scrollTop||0) : 0}});
                        }})()
                    """)
                    recheck = _j.loads(str(recheck_raw)) if recheck_raw else {}
                    new_st = recheck.get("st", data.get("scrollTop_after", 0))
                    data["scrollTop_after"] = new_st
                    data["scroll_changed"] = abs(new_st - data.get("scrollTop_before", 0)) > 0.5
                except Exception:
                    pass
        except Exception:
            pass

    log.info(f"    [MICRO_SCROLL]")
    log.info(f"    slicer={slicer_title}")
    log.info(f"    step={step}")
    log.info(f"    delta={delta}")
    log.info(f"    scroll_method={data.get('scroll_method', 'unknown')}")
    log.info(f"    scrollTop_before={data.get('scrollTop_before')}")
    log.info(f"    scrollTop_after={data.get('scrollTop_after')}")
    log.info(f"    scroll_changed={data.get('scroll_changed')}")
    log.info(f"    render_wait_ms=120")

    return data


async def _enumerate_slicer_via_micro_scroll(tab, idx: int, slicer_title: str, field_name: str,
                                              max_steps: int = 60, delta: int = 16) -> dict:
    """
    Enumeração segura de valores do slicer por micro-scroll incremental.

    Fluxo:
    1. Ativa o slicer com clique seguro no cabeçalho
    2. Localiza o container rolável
    3. Faz micro-scroll repetido
    4. A cada passo lê a "caixa" visível, compara com a anterior
    5. Acumula valores novos
    6. Verifica seleção indevida
    7. Para quando estabilizar
    """
    import json as _j

    log.info(f"    [MICRO_SCROLL_ENUM] Iniciando enumeração por micro-scroll para '{slicer_title}'...")

    # ── Estado operacional ──
    slicer_activated = False
    scroll_container_found = False
    enumeration_started = False
    selection_side_effect = False

    # ── 1. Ativar slicer com clique seguro no cabeçalho ──
    header_pos = await _click_slicer_header_safe(tab, idx)
    log.info(f"    [HEADER_CLICK] {header_pos}")
    if header_pos.get("ok"):
        await _cdp_click(tab, header_pos["x"], header_pos["y"])
        await asyncio.sleep(0.5)
        slicer_activated = True
    else:
        log.info(f"    ⚠️ Cabeçalho não encontrado para ativação")
        return {"values": [], "selected": [], "steps": 0, "success": False, "method": "header_not_found",
                "selection_side_effect": False}

    log.info(f"    slicer_found=True slicer_activated={slicer_activated}")

    # ── 2. Identificar container rolável ──
    sc_data = await _find_scroll_container(tab, idx, slicer_title)
    scroll_container_found = sc_data.get("found", False)
    log.info(f"    scroll_container_found={scroll_container_found}")

    # Snapshot de seleção inicial (ETAPA 8)
    selected_before = await _read_slicer_selected_snapshot(tab, idx)
    log.info(f"    [SELECTION_GUARD] initial_selected={selected_before}")

    # ── 3. Ler caixa inicial (antes de qualquer scroll) ──
    previous_box = await _read_visible_box(tab, idx, field_name, step=0)
    all_unique = set(previous_box["accepted"])
    # ── value_index + enum_context: metadata para locator Fase 2 ──
    _best_sel = previous_box.get("best_selector", ".slicerItemContainer")
    enum_context = {
        "best_selector": _best_sel,
        "initial_scrollTop": 0,
    }
    value_index = {}
    for val in all_unique:
        value_index[val] = {
            "first_seen_step": 0,
            "first_seen_scrollTop": 0,
            "selector": _best_sel,
            "total_enum_steps": 0,
        }
    all_selected = set(previous_box["selected"])
    previous_node_ids = previous_box.get("node_ids", [])

    log.info(f"    [VISIBLE_BOX] slicer={slicer_title} step=0 box_index=0")
    log.info(f"    raw_visible_elements_count={len(previous_box['accepted'])}")
    log.info(f"    raw_texts={previous_box['raw_texts']}")
    log.info(f"    filtered_candidate_values={previous_box['accepted']}")
    log.info(f"    discarded_texts={previous_box['discarded']}")
    log.info(f"    new_values_in_this_box={sorted(list(all_unique))}")
    log.info(f"    all_unique_values_so_far={sorted(list(all_unique))}")
    # TEXT_FILTER já logado internamente por _read_visible_box (v10.4) — não duplicar


    # Se não tem container rolável, retorna o que temos passivamente
    if not scroll_container_found:
        log.info(f"    ⚠️ Sem container rolável — usando apenas valores passivos")
        return {
            "values": sorted(list(all_unique)),
            "selected": sorted(list(all_selected)),
            "steps": 0,
            "success": len(all_unique) > 0,
            "method": "passive_only",
            "selection_side_effect": False,
        }

    # ── 4-7: Loop de micro-scroll com comparação de caixas ──
    enumeration_started = True
    no_new_rounds = 0
    scroll_stuck_rounds = 0
    MAX_NO_NEW = 5
    MAX_SCROLL_STUCK = 4
    steps_done = 0

    for step in range(1, max_steps + 1):
        steps_done = step

        # Micro-scroll
        scroll_result = await _micro_scroll_step(tab, idx, slicer_title, step, delta=delta)
        await asyncio.sleep(0.12)  # render wait

        # Scroll stuck?
        if not scroll_result.get("scroll_changed"):
            scroll_stuck_rounds += 1
            if scroll_stuck_rounds >= MAX_SCROLL_STUCK:
                log.info(f"    [ENUM_STOP] slicer={slicer_title} reason=scroll_stuck steps={steps_done} all_unique_values={sorted(list(all_unique))}")
                break
        else:
            scroll_stuck_rounds = 0

        # Ler caixa atual — DENTRO do scroll container
        current_box = await _read_visible_box(tab, idx, field_name, step=step)

        # ── ETAPA 5: Virtualization check ──
        current_node_ids = current_box.get("node_ids", [])
        node_identity_changed = current_node_ids != previous_node_ids
        # Comparação de texto: usa accepted (valores filtrados) como identidade semântica
        prev_accepted_sorted = sorted(previous_box["accepted"])
        curr_accepted_sorted = sorted(current_box["accepted"])
        text_changed = curr_accepted_sorted != prev_accepted_sorted

        log.info(f"    [VIRTUALIZATION_CHECK] step={step}")
        log.info(f"    previous_box_values={prev_accepted_sorted}")
        log.info(f"    current_box_values={curr_accepted_sorted}")
        log.info(f"    text_changed={text_changed}")
        log.info(f"    node_identity_changed={node_identity_changed}")

        # ── SCROLL_READ_VALIDATION ──
        log.info(
            f"    [SCROLL_READ_VALIDATION] step={step} "
            f"scrollTop={scroll_result.get('scrollTop_after','')} "
            f"visible_values={current_box['accepted'][:10]} "
            f"changed_from_previous={text_changed}"
        )

        # ── Comparação de caixas (ETAPA 6 — BOX_DIFF) ──
        current_vals  = set(current_box["accepted"])
        previous_vals = set(previous_box["accepted"])
        new_vals      = current_vals - all_unique
        removed_vals  = previous_vals - current_vals
        intersection  = current_vals & previous_vals
        window_changed = current_vals != previous_vals

        log.info(f"    [VISIBLE_BOX] slicer={slicer_title} step={step} box_index={step}")
        log.info(f"    raw_visible_elements_count={len(current_box['accepted'])}")
        log.info(f"    raw_texts={current_box['raw_texts'][:15]}")
        log.info(f"    filtered_candidate_values={current_box['accepted']}")
        log.info(f"    new_values_in_this_box={sorted(list(new_vals))}")
        log.info(f"    all_unique_values_so_far={sorted(list(all_unique | new_vals))}")

        log.info(f"    [BOX_DIFF] step={step}")
        log.info(f"    previous_box_values={sorted(list(previous_vals))}")
        log.info(f"    current_box_values={sorted(list(current_vals))}")
        log.info(f"    new_values={sorted(list(new_vals))}")
        log.info(f"    removed_values={sorted(list(removed_vals))}")
        log.info(f"    intersection={sorted(list(intersection))}")
        log.info(f"    window_changed={window_changed}")

        # TEXT_FILTER já foi logado dentro de _read_visible_box (v10.4)
        # — não duplicar aqui. Apenas registra se houve descarte.
        if current_box["discarded"]:
            log.info(
                f"    [TEXT_FILTER_SUMMARY] step={step} "
                f"discarded_count={len(current_box['discarded'])} "
                f"accepted_count={len(current_box['accepted'])}"
            )

        # ── Acumula ──
        all_unique   |= new_vals

        # ── Registrar novos valores no value_index ──
        for val in new_vals:
            if val not in value_index:
                value_index[val] = {
                    "first_seen_step": step,
                    "first_seen_scrollTop": scroll_result.get("scrollTop_after", 0),
                    "selector": current_box.get("best_selector", _best_sel),
                    "total_enum_steps": 0,
                }
        all_selected |= set(current_box["selected"])

        # Guarda de seleção (ETAPA 8)
        selected_after = await _read_slicer_selected_snapshot(tab, idx)
        log.info(f"    [SELECTION_GUARD] slicer={slicer_title} step={step}")
        log.info(f"    selected_before={selected_before}")
        log.info(f"    selected_after={selected_after}")
        sel_changed = sorted(selected_before) != sorted(selected_after)
        log.info(f"    selection_side_effect_detected={sel_changed}")

        if sel_changed:
            selection_side_effect = True
            log.info(f"    [SELECTION_INCIDENT] slicer={slicer_title} step={step}")
            log.info(f"    selected_before={selected_before}")
            log.info(f"    selected_after={selected_after}")
            log.info(f"    last_action=micro_scroll_step_{step}")
            log.info(f"    stop_reason=selection_side_effect")
            log.info(f"    [ENUM_STOP] slicer={slicer_title} reason=selection_side_effect steps={steps_done} all_unique_values={sorted(list(all_unique))}")
            break

        # Critério de parada: sem novos valores
        if not new_vals:
            no_new_rounds += 1
            if no_new_rounds >= MAX_NO_NEW:
                log.info(f"    [ENUM_STOP] slicer={slicer_title} reason=no_new_values steps={steps_done} boxes_processed={step} all_unique_values={sorted(list(all_unique))}")
                break
        else:
            no_new_rounds = 0

        previous_box = current_box
        previous_node_ids = current_node_ids

        # Limite de segurança
        if step >= max_steps:
            log.info(f"    [ENUM_STOP] slicer={slicer_title} reason=max_steps steps={steps_done} all_unique_values={sorted(list(all_unique))}")
            break

    final_values = sorted(list(all_unique))
    final_selected = sorted(list(all_selected))

    # Comparação final (ETAPA 9)
    passive_values = sorted(list(set(previous_box["accepted"]) if steps_done == 0 else set()))
    log.info(f"    [VALUES_COMPARE] slicer={slicer_title}")
    log.info(f"    passive_values={sorted(list(set(previous_box['accepted'])))}")
    log.info(f"    enumerated_values={final_values}")
    log.info(f"    final_clean_values={final_values}")
    log.info(f"    new_values_revealed={len(final_values) > len(set(previous_box['accepted']))}")
    log.info(f"    selection_side_effect_detected={selection_side_effect}")

    log.info(f"    ✅ [MICRO_SCROLL_ENUM] {len(final_values)} valores em {steps_done} passos: {final_values[:20]}")

    # Atualizar total_enum_steps em todos os valores
    for val in value_index:
        value_index[val]["total_enum_steps"] = steps_done

    # Log do índice de valores
    for val_name, meta in sorted(value_index.items()):
        log.info(
            f"    [ENUM_VALUE_INDEX] filter={slicer_title} "
            f"value='{val_name}' "
            f"first_seen_step={meta.get('first_seen_step', '?')} "
            f"first_seen_scrollTop={meta.get('first_seen_scrollTop', '?')} "
            f"selector={meta.get('selector', '?')}"
        )

    return {
        "values": final_values,
        "selected": final_selected,
        "steps": steps_done,
        "success": len(final_values) > 0,
        "method": "micro_scroll",
        "selection_side_effect": selection_side_effect,
        "value_index": value_index,                        # ← LINHA NOVA
        "enum_context": enum_context,
    }

# ---------------------------------------------------------------------------
# Controle Seguro de Filtros (Filter Safe Control Layer)
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Controle Seguro de Filtros — Fase 2
# ---------------------------------------------------------------------------

def normalize_slicer_name(display_name: str) -> str:
    """
    Normaliza título de slicer: remove sufixos de UI, trim, lowercase.
    Usado para matching com FILTER_PLAN.
    """
    normalized = (
        display_name
        .replace("(Ainda não aplicado)", "")
        .replace("(ainda não aplicado)", "")
        .replace("(Not yet applied)", "")
        .replace("(not yet applied)", "")
        .strip()
        .lower()
    )
    return normalized


async def read_current_selection(tab, idx: int, slicer_title: str) -> dict:
    """
    Lê o estado atual do slicer diretamente do DOM:
    - available_values: todos os valores visíveis (exceto controles de UI)
    - selected_values: valores com marcação ativa
    - has_blank, has_select_all, has_clear_control
    - element_map: {text → {x, y, selected, element_type}}

    Log: [FILTER_STATE_BEFORE_APPLY]
    """
    import json as _j
    try:
        raw = await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return JSON.stringify({{error:'no_vc'}});

                const norm = t => (t || '').replace(/\\s+/g, ' ').trim();

                // ── Re-encontrar scroll container ──
                const allEls = Array.from(vc.querySelectorAll('*'));
                let sc = null; let bestScore = -1;
                for (const el of allEls) {{
                    const cs = getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    if (r.width < 20 || r.height < 20) continue;
                    const ov = (cs.overflowY || cs.overflow || '').toLowerCase();
                    const scrollable = el.scrollHeight > el.clientHeight + 2;
                    const hasOv = ov.includes('auto') || ov.includes('scroll');
                    if (!scrollable && !hasOv) continue;
                    const score = (scrollable ? 100 : 0) + (hasOv ? 50 : 0) + (el.scrollHeight - el.clientHeight);
                    if (score > bestScore) {{ bestScore = score; sc = el; }}
                }}
                if (!sc) sc = vc.querySelector('[role="listbox"]') || vc;

                // ── Varredura de itens ──
                const ITEM_SEL = '.slicerItemContainer, [role="option"], [role="listitem"], [class*="slicerItem"]';
                const items = Array.from(sc.querySelectorAll(ITEM_SEL)).filter(el => {{
                    const r = el.getBoundingClientRect();
                    return r.height > 0 && r.width > 0;
                }});

                const available    = [];
                const selected_out = [];
                const element_map  = {{}};
                let has_blank      = false;
                let has_select_all = false;

                items.forEach(item => {{
                    const extractT = (el) => {{
                        const sp = el.querySelector('.slicerText, span[class*="slicerText"]');
                        if (sp) {{ const t1 = norm(sp.textContent); if (t1) return t1;
                                const t2 = norm(sp.innerText); if (t2) return t2; }}
                        const anySp = el.querySelector('span');
                        if (anySp) {{ const t3 = norm(anySp.textContent); if (t3) return t3;
                                    const t4 = norm(anySp.innerText); if (t4) return t4; }}
                        const t5 = norm(el.textContent);
                        if (t5 && t5.length <= 80) return t5;
                        const t6 = norm(el.innerText);
                        if (t6 && t6.length <= 80) return t6;
                        return '';
                    }};
                    const text = extractT(item);
                    if (!text || text.length > 80) return;

                    const low = text.toLowerCase();
                    const isSelectAll = (low === 'selecionar tudo' || low === 'select all');
                    const isBlank     = (low === '(em branco)' || low === '(blank)');

                    const r  = item.getBoundingClientRect();
                    const cx = Math.round(r.left + r.width  / 2);
                    const cy = Math.round(r.top  + r.height / 2);

                    const isSel = (
                        item.classList.contains('selected') ||
                        item.classList.contains('isSelected') ||
                        item.getAttribute('aria-selected') === 'true' ||
                        item.getAttribute('aria-checked')  === 'true' ||
                        !!item.querySelector('.selected, .isSelected, .partiallySelected')
                    );

                    element_map[text] = {{
                        x: cx, y: cy,
                        selected: isSel,
                        element_type: isSelectAll ? 'select_all' : isBlank ? 'blank' : 'value',
                    }};

                    if (isSelectAll) {{ has_select_all = true; return; }}
                    if (isBlank)     {{ has_blank = true; }}

                    available.push(text);
                    if (isSel) selected_out.push(text);
                }});

                // ── Botão Limpar ──
                const clearSelectors = [
                    '.clear-button', '[class*="clearButton"]', '[class*="clear-button"]',
                    'button[title*="Limpar"]', 'button[title*="Clear"]',
                    '[aria-label*="Limpar"]', '[aria-label*="Clear"]',
                    '.slicer-header button', 'visual-container-header button',
                    'visual-header-item-container button',
                ];
                let hasClearControl = false;
                let clearX = 0, clearY = 0;
                for (const sel of clearSelectors) {{
                    const btn = vc.querySelector(sel);
                    if (btn) {{
                        const r = btn.getBoundingClientRect();
                        if (r.width > 0 && r.height > 0) {{
                            hasClearControl = true;
                            clearX = Math.round(r.left + r.width  / 2);
                            clearY = Math.round(r.top  + r.height / 2);
                            break;
                        }}
                    }}
                }}

                return JSON.stringify({{
                    available_values:  available,
                    selected_values:   selected_out,
                    has_blank,
                    has_select_all,
                    has_clear_control: hasClearControl,
                    clear_x: clearX,
                    clear_y: clearY,
                    element_map,
                }});
            }})()
        """)
        data = _j.loads(str(raw)) if raw else {}
    except Exception as e:
        data = {
            "error": str(e),
            "available_values": [], "selected_values": [],
            "has_blank": False, "has_select_all": False,
            "has_clear_control": False, "element_map": {},
        }

    log.info(f"    [FILTER_STATE_BEFORE_APPLY]")
    log.info(f"    filter={slicer_title}")
    log.info(f"    available_values={data.get('available_values', [])[:30]}")
    log.info(f"    selected_values={data.get('selected_values', [])}")
    log.info(f"    has_blank={data.get('has_blank', False)}")
    log.info(f"    has_select_all={data.get('has_select_all', False)}")
    log.info(f"    has_clear_control={data.get('has_clear_control', False)}")

    return data


async def clear_slicer_selection(tab, idx: int, slicer_title: str, state: dict) -> bool:
    """
    Limpa seleção do slicer de forma segura.
    Prioridade:
      1. Botão Limpar explícito
      2. Toggle "Selecionar Tudo" (2 cliques para garantir estado limpo)
      3. Desmarcação individual dos selecionados conhecidos

    Nunca faz clique cego em massa.
    Log: [FILTER_CLEAR]
    """
    sel_before = list(state.get("selected_values", []))

    if not sel_before:
        log.info(
            f"    [FILTER_CLEAR] filter={slicer_title} method=nothing_to_clear "
            f"selected_before=[] selected_after=[] clear_ok=True"
        )
        return True

    # ── Método 1: botão Limpar ──
    if state.get("has_clear_control") and state.get("clear_x") and state.get("clear_y"):
        cx, cy = state["clear_x"], state["clear_y"]
        log.info(
            f"    [FILTER_CLEAR] filter={slicer_title} method=clear_button "
            f"coords=({cx},{cy}) selected_before={sel_before}"
        )
        await _cdp_click(tab, cx, cy)
        await asyncio.sleep(0.6)
        new_state = await read_current_selection(tab, idx, slicer_title)
        sel_after = new_state.get("selected_values", [])
        clear_ok  = len(sel_after) == 0
        log.info(
            f"    [FILTER_CLEAR] filter={slicer_title} method=clear_button "
            f"selected_after={sel_after} clear_ok={clear_ok}"
        )
        if clear_ok:
            return True
        # Atualiza state para próximo método
        state = new_state

    # ── Método 2: Selecionar Tudo (toggle duplo) ──
    emap = state.get("element_map", {})
    select_all_key = next(
        (k for k in emap if k.lower() in ("selecionar tudo", "select all")), None
    )
    if select_all_key:
        entry = emap[select_all_key]
        cx, cy = entry["x"], entry["y"]
        log.info(
            f"    [FILTER_CLEAR] filter={slicer_title} method=select_all_toggle "
            f"coords=({cx},{cy}) selected_before={sel_before}"
        )
        for click_n in range(1, 3):
            await _cdp_click(tab, cx, cy)
            await asyncio.sleep(0.5)
            check = await read_current_selection(tab, idx, slicer_title)
            sel_after = check.get("selected_values", [])
            if not sel_after:
                log.info(
                    f"    [FILTER_CLEAR] filter={slicer_title} method=select_all_toggle "
                    f"clicks={click_n} selected_after=[] clear_ok=True"
                )
                return True
        log.info(
            f"    [FILTER_CLEAR] filter={slicer_title} method=select_all_toggle "
            f"selected_after={sel_after} clear_ok=False — tentando método 3"
        )
        state = check

    # ── Método 3: desmarcação individual ──
    current = await read_current_selection(tab, idx, slicer_title)
    to_deselect = list(current.get("selected_values", []))
    log.info(
        f"    [FILTER_CLEAR] filter={slicer_title} method=individual_deselect "
        f"to_deselect={to_deselect} selected_before={to_deselect}"
    )
    emap = current.get("element_map", {})
    for val in to_deselect:
        entry = emap.get(val)
        if not entry:
            log.info(f"    [FILTER_CLEAR] filter={slicer_title} value='{val}' element_not_found — skip")
            continue
        if not entry.get("selected", False):
            log.info(f"    [FILTER_CLEAR] filter={slicer_title} value='{val}' already_deselected — skip")
            continue
        cx, cy = entry["x"], entry["y"]
        await _cdp_click(tab, cx, cy)
        await asyncio.sleep(0.4)

    final = await read_current_selection(tab, idx, slicer_title)
    sel_after = final.get("selected_values", [])
    clear_ok  = len(sel_after) == 0
    log.info(
        f"    [FILTER_CLEAR] filter={slicer_title} method=individual_deselect "
        f"selected_after={sel_after} clear_ok={clear_ok}"
    )
    return clear_ok


async def validate_filter_final(
    tab, idx: int, slicer_title: str,
    target_values: list, mode: str,
) -> dict:
    """
    Lê estado final e compara com target_values.
    Log: [FILTER_FINAL_VALIDATE]
    """
    final_state    = await read_current_selection(tab, idx, slicer_title)
    final_selected = final_state.get("selected_values", [])

    target_set = {v.lower() for v in target_values}
    final_set  = {s.lower() for s in final_selected}

    if mode == "single":
        missing    = [v for v in target_values if v.lower() not in final_set]
        unexpected = [s for s in final_selected if s.lower() not in target_set]
        validation_ok = (len(missing) == 0 and len(unexpected) == 0)
    else:  # multi
        missing    = [v for v in target_values if v.lower() not in final_set]
        unexpected = []  # multi permite outros selecionados além dos target
        validation_ok = len(missing) == 0

    log.info(f"    [FILTER_FINAL_VALIDATE]")
    log.info(f"    filter={slicer_title}")
    log.info(f"    target_values={target_values}")
    log.info(f"    final_selected={final_selected}")
    log.info(f"    missing={missing}")
    log.info(f"    unexpected={unexpected}")
    log.info(f"    validation_ok={validation_ok}")

    return {
        "validation_ok": validation_ok,
        "final_selected": final_selected,
        "missing": missing,
        "unexpected": unexpected,
    }

# ═══════════════════════════════════════════════════════════════════════
# FUNÇÃO 1: _scroll_to_find_value_in_slicer (REESCRITA COMPLETA)
#
# SUBSTITUI a versão anterior.
# Posição no arquivo: ANTES de apply_filter_plan
# ═══════════════════════════════════════════════════════════════════════

async def _scroll_to_find_value_in_slicer(tab, idx: int, target_value: str,
                                           slicer_title: str,
                                           enum_value_index: dict = None,
                                           max_scroll_steps: int = 40,
                                           enum_context: dict = None,
                                           delta: int = 18) -> dict:
    """
    Locator em 3 camadas para encontrar o nó clicável de um valor no slicer.

    Camada A — procura direta no viewport atual
    Camada B — reposicionamento por índice da enumeração (enum hint)
    Camada C — scroll search dirigido a partir da faixa estimada

    Retorna:
      {"found": bool, "x": int, "y": int, "selected": bool,
       "text": str, "method": str, "steps": int,
       "layer": "A"|"B"|"C", "tag": str, "cls": str}
    """
    import json as _j
    target_low = target_value.lower().strip()

    # ── Resolver enum hint ──
    hint = {}
    if enum_value_index and isinstance(enum_value_index, dict):
        hint = enum_value_index.get(target_value) or {}
        if not hint:
            # Case-insensitive fallback
            for k, v in enum_value_index.items():
                if k.lower().strip() == target_low:
                    hint = v
                    break

    enum_seen = bool(hint)
    first_seen_step = hint.get("first_seen_step", -1)
    first_seen_scrollTop = hint.get("first_seen_scrollTop", 0)
    # Selector real: prioridade enum_context > value_index > default
    _ctx = enum_context or {}
    _default_sel = _ctx.get("best_selector", ".slicerItemContainer")
    selector_hint = hint.get("selector", _default_sel)
    if selector_hint in ("initial_box", "micro_scroll", "unknown"):
        selector_hint = _default_sel

    log.info(f"    [FILTER_LOCATOR_CONTEXT]")
    log.info(f"    filter={slicer_title}")
    log.info(f"    target_value='{target_value}'")
    log.info(f"    best_selector_from_enum={_default_sel}")
    log.info(f"    selector_hint_resolved={selector_hint}")
    log.info(f"    first_seen_scrollTop={first_seen_scrollTop}")
    log.info(f"    estimated_scroll_band={estimated_band}")
    
    total_enum_steps = hint.get("total_enum_steps", 0)

    # Estimar banda de scroll
    if total_enum_steps > 0 and first_seen_step >= 0:
        ratio = first_seen_step / max(total_enum_steps, 1)
        if ratio < 0.25:
            estimated_band = "top"
        elif ratio < 0.75:
            estimated_band = "middle"
        else:
            estimated_band = "bottom"
    else:
        estimated_band = "unknown"

    log.info(f"    [FILTER_VALUE_INDEX_HINT]")
    log.info(f"    filter={slicer_title}")
    log.info(f"    target_value='{target_value}'")
    log.info(f"    enum_seen={enum_seen}")
    log.info(f"    first_seen_step={first_seen_step}")
    log.info(f"    first_seen_scrollTop={first_seen_scrollTop}")
    log.info(f"    estimated_scroll_band={estimated_band}")
    log.info(f"    selector_hint={selector_hint}")

    # ── Helper JS: encontrar item no viewport atual ──
    async def _find_in_viewport(attempt_label: str, search_mode: str) -> dict:
        """Procura o target no viewport usando a mesma lógica da enumeração."""
        escaped_target = _j.dumps(target_low)
        escaped_selector = _j.dumps(selector_hint)
        try:
            raw = await tab.evaluate(f"""
                (() => {{
                    const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                    if (!vc) return JSON.stringify({{found: false, reason: 'no_vc', items: []}});
                    const target = {escaped_target};
                    const hintSel = {escaped_selector};

                    // Mesma cadeia de seletores da enumeração (SELECTOR_TEST)
                    const selectors = [
                        hintSel,
                        '.slicerItemContainer',
                        '[role="option"]',
                        '[role="listitem"]',
                        '[class*="slicerItem"]',
                        'div.row',
                    ];
                    const uniqueSelectors = [...new Set(selectors)];

                    // scrollTop atual
                    let scrollTop = 0;
                    const allEls = Array.from(vc.querySelectorAll('*'));
                    for (const el of allEls) {{
                        const cs = getComputedStyle(el);
                        const ov = (cs.overflowY || cs.overflow || '').toLowerCase();
                        if ((ov.includes('auto') || ov.includes('scroll')) && el.scrollHeight > el.clientHeight + 2) {{
                            scrollTop = Math.round(el.scrollTop);
                            break;
                        }}
                    }}

                    // ── Normalização ÚNICA (mesma da enumeração) ──
                    const normalize = (t) => (t || '').replace(/\\s+/g, ' ').trim();
                    const normLower = (t) => normalize(t).toLowerCase();

                    // ── Extração de texto IDÊNTICA à _read_visible_box ──
                    const extractText = (el) => {{
                        // Cadeia 1: .slicerText span (mesma da enumeração)
                        const span = el.querySelector('.slicerText, span[class*="slicerText"]');
                        if (span) {{
                            const t1 = normalize(span.textContent);
                            if (t1) return t1;
                            const t2 = normalize(span.innerText);
                            if (t2) return t2;
                        }}
                        // Cadeia 2: primeiro span direto
                        const anySpan = el.querySelector('span');
                        if (anySpan) {{
                            const t3 = normalize(anySpan.textContent);
                            if (t3) return t3;
                            const t4 = normalize(anySpan.innerText);
                            if (t4) return t4;
                        }}
                        // Cadeia 3: textContent do próprio elemento
                        const t5 = normalize(el.textContent);
                        if (t5 && t5.length <= 80) return t5;
                        // Cadeia 4: innerText
                        const t6 = normalize(el.innerText);
                        if (t6 && t6.length <= 80) return t6;
                        return '';
                    }};

                    const itemsDebug = [];
                    let bestSelectorUsed = '';

                    for (const sel of uniqueSelectors) {{
                        const elements = Array.from(vc.querySelectorAll(sel)).filter(el => {{
                            const r = el.getBoundingClientRect();
                            return r.height > 0 && r.width > 0;
                        }});
                        if (elements.length === 0) continue;

                        for (const item of elements) {{
                            const rawText = extractText(item);
                            const normText = normLower(rawText);

                            // Debug: gravar todos os itens (até 15)
                            if (itemsDebug.length < 15) {{
                                const span = item.querySelector('.slicerText, span[class*="slicerText"], span');
                                itemsDebug.push({{
                                    textContent: normalize(item.textContent).slice(0, 60),
                                    innerText: normalize(item.innerText).slice(0, 60),
                                    spanText: span ? normalize(span.textContent).slice(0, 60) : '(no span)',
                                    spanInner: span ? normalize(span.innerText).slice(0, 60) : '(no span)',
                                    extracted: rawText,
                                    normalized: normText,
                                }});
                            }}

                            if (normText === target) {{
                                const r = item.getBoundingClientRect();
                                const isSel = (
                                    item.classList.contains('selected') ||
                                    item.classList.contains('isSelected') ||
                                    item.getAttribute('aria-selected') === 'true' ||
                                    item.getAttribute('aria-checked') === 'true' ||
                                    !!item.querySelector('.selected, .isSelected')
                                );
                                return JSON.stringify({{
                                    found: true,
                                    x: Math.round(r.left + r.width / 2),
                                    y: Math.round(r.top + r.height / 2),
                                    text: rawText,
                                    selected: isSel,
                                    tag: item.tagName.toLowerCase(),
                                    cls: (item.className || '').toString().slice(0, 80),
                                    selector_used: sel,
                                    visible: r.width > 0 && r.height > 0,
                                    in_viewport: r.top >= 0 && r.bottom <= window.innerHeight,
                                    scrollTop: scrollTop,
                                    texts_visible: itemsDebug.map(d => d.extracted),
                                    items_debug: itemsDebug,
                                }});
                            }}
                        }}

                        // Se este seletor encontrou itens, é o melhor
                        if (elements.length > 0 && !bestSelectorUsed) {{
                            bestSelectorUsed = sel;
                        }}
                    }}

                    return JSON.stringify({{
                        found: false,
                        scrollTop: scrollTop,
                        texts_visible: itemsDebug.map(d => d.extracted),
                        items_debug: itemsDebug,
                        selector_used: bestSelectorUsed,
                        reason: itemsDebug.length === 0 ? 'no_visible_elements' :
                                itemsDebug.every(d => !d.extracted) ? 'empty_texts_extracted' :
                                'text_not_matched',
                    }});
                }})()
            """)
            data = _j.loads(str(raw)) if raw else {"found": False}
        except Exception as e:
            data = {"found": False, "reason": f"js_error:{str(e)[:60]}"}

        # ── Log VISIBLE_ITEMS_DEBUG ──
        items_debug = data.get("items_debug", [])
        if items_debug:
            log.info(f"    [VISIBLE_ITEMS_DEBUG]")
            log.info(f"    filter={slicer_title}")
            log.info(f"    selector={data.get('selector_used', '?')}")
            log.info(f"    elements_found={len(items_debug)}")
            for di, item_d in enumerate(items_debug[:10]):
                log.info(
                    f"    item[{di}]: textContent='{item_d.get('textContent','')}' "
                    f"innerText='{item_d.get('innerText','')}' "
                    f"spanText='{item_d.get('spanText','')}' "
                    f"spanInner='{item_d.get('spanInner','')}' "
                    f"extracted='{item_d.get('extracted','')}' "
                    f"normalized='{item_d.get('normalized','')}'"
                )

        # ── Log DOM_INVALID_STATE se textos vazios ──
        texts_vis = data.get("texts_visible", [])
        if items_debug and all(not d.get("extracted") for d in items_debug):
            log.warning(
                f"    [DOM_INVALID_STATE] filter={slicer_title} "
                f"reason=empty_texts_extracted "
                f"elements_found={len(items_debug)} "
                f"action=will_retry_or_fail"
            )

        # ── Log LOCATOR_ATTEMPT ──
        log.info(f"    [FILTER_VALUE_LOCATOR_ATTEMPT]")
        log.info(f"    filter={slicer_title}")
        log.info(f"    target_value='{target_value}'")
        log.info(f"    attempt={attempt_label}")
        log.info(f"    search_mode={search_mode}")
        log.info(f"    scrollTop={data.get('scrollTop', '?')}")
        log.info(f"    selector={data.get('selector_used', '?')}")
        log.info(f"    texts_visible={texts_vis[:10]}")
        log.info(f"    dom_locator_found={data.get('found', False)}")

        if data.get("found"):
            # ── Log TEXT_MATCH_DEBUG ──
            log.info(f"    [TEXT_MATCH_DEBUG]")
            log.info(f"    target_raw='{target_value}'")
            log.info(f"    target_norm='{target_low}'")
            log.info(f"    candidate_raw='{data.get('text', '?')}'")
            log.info(f"    candidate_norm='{data.get('text', '').lower().strip()}'")
            log.info(f"    match=True")

            # ── Log CLICK_TARGET ──
            log.info(f"    [FILTER_VALUE_CLICK_TARGET]")
            log.info(f"    filter={slicer_title}")
            log.info(f"    target='{target_value}'")
            log.info(f"    normalized='{target_low}'")
            log.info(f"    found=True")
            log.info(f"    element_tag={data.get('tag', '?')}")
            log.info(f"    element_class={data.get('cls', '?')}")
            log.info(f"    target_text='{data.get('text', '?')}'")
            log.info(f"    visible={data.get('visible', False)}")
            log.info(f"    clickable={data.get('in_viewport', False)}")

        return data

    # ── Helper: set scrollTop on the scroll container ──
    async def _set_scroll_position(target_scrollTop: int) -> bool:
        try:
            raw = await tab.evaluate(f"""
                (() => {{
                    const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                    if (!vc) return JSON.stringify({{ok:false}});
                    const allEls = Array.from(vc.querySelectorAll('*'));
                    for (const el of allEls) {{
                        const cs = getComputedStyle(el);
                        const ov = (cs.overflowY || cs.overflow || '').toLowerCase();
                        if ((ov.includes('auto') || ov.includes('scroll')) && el.scrollHeight > el.clientHeight + 2) {{
                            el.scrollTop = {target_scrollTop};
                            return JSON.stringify({{ok:true, actual: Math.round(el.scrollTop)}});
                        }}
                    }}
                    return JSON.stringify({{ok:false, reason:'no_scroll_container'}});
                }})()
            """)
            data = _j.loads(str(raw)) if raw else {}
            return data.get("ok", False)
        except Exception:
            return False

    # ── Helper: micro-scroll from current position ──
    async def _micro_scroll(scroll_delta: int) -> dict:
        try:
            raw = await tab.evaluate(f"""
                (() => {{
                    const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                    if (!vc) return JSON.stringify({{ok:false}});
                    const allEls = Array.from(vc.querySelectorAll('*'));
                    for (const el of allEls) {{
                        const cs = getComputedStyle(el);
                        const ov = (cs.overflowY || cs.overflow || '').toLowerCase();
                        if ((ov.includes('auto') || ov.includes('scroll')) && el.scrollHeight > el.clientHeight + 2) {{
                            const before = el.scrollTop;
                            el.scrollTop += {scroll_delta};
                            const after = el.scrollTop;
                            return JSON.stringify({{
                                ok: true,
                                changed: Math.abs(after - before) > 0.5,
                                scrollTop: Math.round(after),
                                scrollHeight: Math.round(el.scrollHeight),
                                clientHeight: Math.round(el.clientHeight),
                            }});
                        }}
                    }}
                    return JSON.stringify({{ok:false}});
                }})()
            """)
            return _j.loads(str(raw)) if raw else {"ok": False}
        except Exception:
            return {"ok": False}

# ── Forçar scrollTop=0 para estado DOM consistente ──
    try:
        await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return;
                for (const el of Array.from(vc.querySelectorAll('*'))) {{
                    const cs = getComputedStyle(el);
                    const ov = (cs.overflowY || cs.overflow || '').toLowerCase();
                    if ((ov.includes('auto') || ov.includes('scroll')) && el.scrollHeight > el.clientHeight + 2) {{
                        el.scrollTop = 0; return;
                    }}
                }}
            }})()
        """)
        await asyncio.sleep(0.2)
    except Exception:
        pass

    log.info(
        f"    [FILTER_SCROLL_REPOSITION] filter={slicer_title} "
        f"target_value='{target_value}' "
        f"from_scrollTop=unknown to_scrollTop=0 reason=initial_reset"
    )
    
    # ══════════════════════════════════════════════════════════════════
    # CAMADA A — Procura direta no viewport atual
    # ══════════════════════════════════════════════════════════════════
    result_a = await _find_in_viewport("A1", "viewport_direct")
    if result_a.get("found"):
        return {
            "found": True,
            "x": result_a["x"], "y": result_a["y"],
            "selected": result_a.get("selected", False),
            "text": result_a.get("text", target_value),
            "method": "viewport_direct",
            "layer": "A",
            "steps": 0,
            "tag": result_a.get("tag", ""),
            "cls": result_a.get("cls", ""),
        }
        # ── Reativação se DOM inválido ──
    a_items = result_a.get("items_debug", [])
    a_all_empty = a_items and all(not d.get("extracted") for d in a_items)
    if a_all_empty:
        log.info(
            f"    [DOM_INVALID_STATE] filter={slicer_title} "
            f"reason=empty_texts_after_camada_A "
            f"action=retry_activation"
        )
        # Reativar slicer
        header_retry = await _click_slicer_header_safe(tab, idx)
        if header_retry.get("ok"):
            await _cdp_click(tab, header_retry["x"], header_retry["y"])
            await asyncio.sleep(0.8)
            # Retry Camada A
            result_a2 = await _find_in_viewport("A2_reactivated", "viewport_after_reactivation")
            if result_a2.get("found"):
                return {
                    "found": True,
                    "x": result_a2["x"], "y": result_a2["y"],
                    "selected": result_a2.get("selected", False),
                    "text": result_a2.get("text", target_value),
                    "method": "viewport_after_reactivation",
                    "layer": "A",
                    "steps": 0,
                    "tag": result_a2.get("tag", ""),
                    "cls": result_a2.get("cls", ""),
                }


    # ══════════════════════════════════════════════════════════════════
    # CAMADA B — Reposicionamento por índice da enumeração
    # ══════════════════════════════════════════════════════════════════
    if enum_seen and first_seen_scrollTop > 0:
        log.info(
            f"    [FILTER_VALUE_LOCATOR] Camada B: reposicionando scroll "
            f"para scrollTop={first_seen_scrollTop} (enum hint step {first_seen_step})"
        )
        reposition_ok = await _set_scroll_position(first_seen_scrollTop)
        await asyncio.sleep(0.15)

        if reposition_ok:
            result_b = await _find_in_viewport("B1_exact", "enum_hint_exact")
            if result_b.get("found"):
                return {
                    "found": True,
                    "x": result_b["x"], "y": result_b["y"],
                    "selected": result_b.get("selected", False),
                    "text": result_b.get("text", target_value),
                    "method": "enum_hint_exact",
                    "layer": "B",
                    "steps": 0,
                    "tag": result_b.get("tag", ""),
                    "cls": result_b.get("cls", ""),
                }

            # Não exato: tenta vizinhança do hint (-2δ a +4δ)
            for offset_mult in [-2, -1, 1, 2, 3, 4]:
                nearby_scrollTop = max(0, first_seen_scrollTop + (offset_mult * delta))
                await _set_scroll_position(nearby_scrollTop)
                await asyncio.sleep(0.1)
                result_b2 = await _find_in_viewport(
                    f"B2_offset_{offset_mult}", "enum_hint_nearby"
                )
                if result_b2.get("found"):
                    return {
                        "found": True,
                        "x": result_b2["x"], "y": result_b2["y"],
                        "selected": result_b2.get("selected", False),
                        "text": result_b2.get("text", target_value),
                        "method": f"enum_hint_nearby_offset_{offset_mult}",
                        "layer": "B",
                        "steps": abs(offset_mult),
                        "tag": result_b2.get("tag", ""),
                        "cls": result_b2.get("cls", ""),
                    }

    # ══════════════════════════════════════════════════════════════════
    # CAMADA C — Scroll search dirigido
    # Começa do ponto de hint (ou do topo se sem hint)
    # ══════════════════════════════════════════════════════════════════
    start_scrollTop = first_seen_scrollTop if enum_seen else 0

    # Reseta para o ponto de partida
    await _set_scroll_position(start_scrollTop)
    await asyncio.sleep(0.15)

    log.info(
        f"    [FILTER_VALUE_LOCATOR] Camada C: scroll dirigido "
        f"start_scrollTop={start_scrollTop} delta={delta} max_steps={max_scroll_steps}"
    )

    # Fase C1: scroll para FRENTE (para baixo) a partir do ponto de partida
    for step in range(max_scroll_steps):
        result_c = await _find_in_viewport(f"C_fwd_{step}", "scroll_search_forward")
        if result_c.get("found"):
            return {
                "found": True,
                "x": result_c["x"], "y": result_c["y"],
                "selected": result_c.get("selected", False),
                "text": result_c.get("text", target_value),
                "method": "scroll_search_forward",
                "layer": "C",
                "steps": step,
                "tag": result_c.get("tag", ""),
                "cls": result_c.get("cls", ""),
            }

        scroll_data = await _micro_scroll(delta)
        await asyncio.sleep(0.1)

        if scroll_data.get("ok") and not scroll_data.get("changed"):
            log.info(
                f"    [FILTER_VALUE_LOCATOR] Camada C forward: "
                f"scroll exhausted at step {step}"
            )
            break

    # Fase C2: se começou do meio (enum hint), tenta scroll para TRÁS
    if start_scrollTop > 0:
        await _set_scroll_position(max(0, start_scrollTop - delta))
        await asyncio.sleep(0.15)

        backward_steps = min(max_scroll_steps // 2, 15)
        for step in range(backward_steps):
            result_c2 = await _find_in_viewport(f"C_bwd_{step}", "scroll_search_backward")
            if result_c2.get("found"):
                return {
                    "found": True,
                    "x": result_c2["x"], "y": result_c2["y"],
                    "selected": result_c2.get("selected", False),
                    "text": result_c2.get("text", target_value),
                    "method": "scroll_search_backward",
                    "layer": "C",
                    "steps": step,
                    "tag": result_c2.get("tag", ""),
                    "cls": result_c2.get("cls", ""),
                }

            scroll_data = await _micro_scroll(-delta)
            await asyncio.sleep(0.1)

            if scroll_data.get("ok") and not scroll_data.get("changed"):
                break

    # ── Falha final: diagnóstico do motivo ──
    failure_reason = "unknown"
    failure_details = ""
    if not enum_seen:
        failure_reason = "no_enum_hint_available"
        failure_details = "valor não tinha metadata de enumeração"
    elif first_seen_scrollTop == 0 and estimated_band == "top":
        failure_reason = "stale_dom_after_activation"
        failure_details = "valor era visível no topo durante enum mas desapareceu após reativação"
    else:
        # Tenta diagnóstico: verifica se o texto existe em QUALQUER parte do slicer
        try:
            diag_raw = await tab.evaluate(f"""
                (() => {{
                    const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                    if (!vc) return JSON.stringify({{diag: 'no_vc'}});
                    const target = {_j.dumps(target_low)};
                    const allText = (vc.innerText || '').toLowerCase();
                    const textExists = allText.includes(target);
                    const allItems = Array.from(vc.querySelectorAll(
                        '.slicerItemContainer, [role="option"], [role="listitem"], [class*="slicerItem"]'
                    ));
                    const totalItems = allItems.length;
                    const visibleItems = allItems.filter(e => {{
                        const r = e.getBoundingClientRect();
                        return r.height > 0 && r.width > 0;
                    }}).length;
                    return JSON.stringify({{
                        text_in_innerText: textExists,
                        total_items: totalItems,
                        visible_items: visibleItems,
                    }});
                }})()
            """)
            diag = _j.loads(str(diag_raw)) if diag_raw else {}
        except Exception:
            diag = {}

        if diag.get("text_in_innerText"):
            failure_reason = "selector_found_text_but_not_clickable"
            failure_details = (
                f"texto existe no innerText do slicer mas nenhum seletor "
                f"capturou como item clicável. total_items={diag.get('total_items',0)} "
                f"visible_items={diag.get('visible_items',0)}"
            )
        elif diag.get("visible_items", 0) == 0:
            failure_reason = "stale_dom_after_activation"
            failure_details = (
                f"nenhum item visível no slicer. total_items={diag.get('total_items',0)}"
            )
        else:
            failure_reason = "text_normalization_mismatch"
            failure_details = (
                f"itens visíveis existem ({diag.get('visible_items',0)}) "
                f"mas nenhum match textual para '{target_value}'"
            )

    log.info(f"    [FILTER_LOCATOR_FAILURE_REASON]")
    log.info(f"    filter={slicer_title}")
    log.info(f"    target_value='{target_value}'")
    log.info(f"    reason={failure_reason}")
    log.info(f"    details={failure_details}")

    return {
        "found": False,
        "method": f"all_layers_exhausted_{failure_reason}",
        "layer": "exhausted",
        "steps": max_scroll_steps,
        "failure_reason": failure_reason,
        "failure_details": failure_details,
    }

# ═══════════════════════════════════════════════════════════════════════
# FUNÇÃO 2: apply_filter_plan (ATUALIZADA)
#
# Mudança: recebe enum_value_index e repassa ao locator.
# SUBSTITUI a versão da v07 (que já tem o enum fallback).
# Posição no arquivo: APÓS _scroll_to_find_value_in_slicer,
#                      ANTES de _read_filter_state
# ═══════════════════════════════════════════════════════════════════════

async def apply_filter_plan(
    tab, idx: int,
    slicer_title: str,
    final_clean_values: list,
    plan_config: dict,
    enum_value_index: dict = None,
    enum_context: dict = None,
) -> dict:
    """
    Orquestra a Fase 2 para UM slicer.

    CORREÇÃO v2: além de usar final_clean_values como fonte de existência,
    agora usa enum_value_index para localizar elementos via scroll dirigido.

    Fluxo:
      [FILTER_ENUM_SNAPSHOT] → [FILTER_ACTIVATION] → [FILTER_AVAILABLE_SOURCE]
      → [FILTER_CLEAR]
      → por valor:
          [FILTER_VALUE_EXISTENCE] → [FILTER_VALUE_INDEX_HINT]
          → [FILTER_VALUE_LOCATOR_ATTEMPT] (camadas A/B/C)
          → [FILTER_VALUE_CLICK_TARGET] → [FILTER_VALUE_APPLY]
      → [FILTER_FINAL_VALIDATE]
    """
    mode          = plan_config.get("mode", "single")
    clear_first   = plan_config.get("clear_first", True)
    target_values = list(plan_config.get("target_values", []))
    if enum_value_index is None:
        enum_value_index = {}
    if enum_context is None:
        enum_context = {}

    result = {
        "validation_ok": False,
        "final_selected": [],
        "per_value": {},
        "aborted": False,
    }

    # ══════════════════════════════════════════════════════════════════
    # ETAPA 1: Congelar snapshot da Fase 1 (enumeração)
    # ══════════════════════════════════════════════════════════════════
    enum_available = list(final_clean_values) if final_clean_values else []
    log.info(f"    [FILTER_ENUM_SNAPSHOT]")
    log.info(f"    filter={slicer_title}")
    log.info(f"    enumerated_values={enum_available}")
    log.info(f"    enum_value_index_keys={list(enum_value_index.keys())[:15]}")

    log.info(f"    [FILTER_PLAN_MATCH]")
    log.info(f"    matched_key={normalize_slicer_name(slicer_title)}")
    log.info(f"    mode={mode}")
    log.info(f"    clear_first={clear_first}")
    log.info(f"    target_values={target_values}")

    # ══════════════════════════════════════════════════════════════════
    # ETAPA 2: REATIVAR o slicer
    # ══════════════════════════════════════════════════════════════════
    header_pos = await _click_slicer_header_safe(tab, idx)
    activation_ok = False
    activation_method = "none"
    if header_pos.get("ok"):
        await _cdp_click(tab, header_pos["x"], header_pos["y"])
        await asyncio.sleep(0.6)
        activation_ok = True
        activation_method = "header_click"
    else:
        try:
            import json as _jfb
            fb_raw = await tab.evaluate(f"""
                (() => {{
                    const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                    if (!vc) return JSON.stringify({{ok:false}});
                    const r = vc.getBoundingClientRect();
                    return JSON.stringify({{
                        ok: r.width > 0,
                        x: Math.round(r.left + r.width / 2),
                        y: Math.round(r.top  + r.height / 2),
                    }});
                }})()
            """)
            fb = _jfb.loads(str(fb_raw)) if fb_raw else {}
            if fb.get("ok"):
                await _cdp_click(tab, fb["x"], fb["y"])
                await asyncio.sleep(0.6)
                activation_ok = True
                activation_method = "center_fallback"
        except Exception:
            pass

    log.info(
        f"    [FILTER_ACTIVATION] filter={slicer_title} "
        f"method={activation_method} ok={activation_ok}"
    )

    if not activation_ok:
        log.warning(
            f"    [FILTER_INCIDENT] filter={slicer_title} "
            f"stage=activation reason=could_not_activate_slicer "
            f"action_taken=abort"
        )
        result["aborted"] = True
        return result

    # ══════════════════════════════════════════════════════════════════
    # ETAPA 3: Ler estado atual do DOM
    # ══════════════════════════════════════════════════════════════════
    # Forçar scrollTop=0 para estado consistente antes de ler
    try:
        await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return;
                const allEls = Array.from(vc.querySelectorAll('*'));
                for (const el of allEls) {{
                    const cs = getComputedStyle(el);
                    const ov = (cs.overflowY || cs.overflow || '').toLowerCase();
                    if ((ov.includes('auto') || ov.includes('scroll')) && el.scrollHeight > el.clientHeight + 2) {{
                        el.scrollTop = 0;
                        return;
                    }}
                }}
            }})()
        """)
        await asyncio.sleep(0.3)
    except Exception:
        pass
    state         = await read_current_selection(tab, idx, slicer_title)
    dom_available = list(state.get("available_values", []))

    if not dom_available and enum_available:
        log.info(
            f"    [FILTER_ACTIVATION_RETRY] filter={slicer_title} "
            f"dom_available_empty — waiting 1s and retrying"
        )
        await asyncio.sleep(1.0)
        state         = await read_current_selection(tab, idx, slicer_title)
        dom_available = list(state.get("available_values", []))

    # ══════════════════════════════════════════════════════════════════
    # ETAPA 4: effective_available (enum é fonte oficial)
    # ══════════════════════════════════════════════════════════════════
    if dom_available:
        dom_set = set(v.lower() for v in dom_available)
        enum_extras = [v for v in enum_available if v.lower() not in dom_set]
        effective_available = dom_available + enum_extras
        source_used = "dom" if not enum_extras else "merged"
    elif enum_available:
        effective_available = list(enum_available)
        source_used = "enum_fallback"
    else:
        effective_available = []
        source_used = "empty"

    log.info(f"    [FILTER_AVAILABLE_SOURCE]")
    log.info(f"    filter={slicer_title}")
    log.info(f"    enum_available={enum_available[:20]}")
    log.info(f"    dom_available={dom_available[:20]}")
    log.info(f"    effective_available={effective_available[:20]}")
    log.info(f"    source_used={source_used}")

    log.info(f"    [FILTER_STATE_BEFORE_APPLY]")
    log.info(f"    filter={slicer_title}")
    log.info(f"    available_values={effective_available[:30]}")
    log.info(f"    selected_values={state.get('selected_values', [])}")
    log.info(f"    source_used={source_used}")

    if not effective_available:
        log.warning(
            f"    [FILTER_INCIDENT] filter={slicer_title} "
            f"stage=read_state reason=no_values_from_any_source "
            f"action_taken=abort"
        )
        result["aborted"] = True
        return result

    # ══════════════════════════════════════════════════════════════════
    # ETAPA 5: Verificar existência
    # ══════════════════════════════════════════════════════════════════
    effective_lower = {v.lower() for v in effective_available}
    for tv in target_values:
        exists = tv.lower() in effective_lower
        log.info(
            f"    [FILTER_VALUE_EXISTENCE] filter={slicer_title} "
            f"target_value='{tv}' "
            f"exists_in_effective_available={exists} "
            f"source_used={source_used}"
        )

    # ══════════════════════════════════════════════════════════════════
    # ETAPA 6: Limpar se clear_first
    # ══════════════════════════════════════════════════════════════════
    if clear_first:
        clear_ok = await clear_slicer_selection(tab, idx, slicer_title, state)
        if not clear_ok:
            log.warning(
                f"    [FILTER_INCIDENT] filter={slicer_title} "
                f"stage=clear reason=clear_failed action_taken=abort"
            )
            result["aborted"] = True
            return result
        state = await read_current_selection(tab, idx, slicer_title)

    # ══════════════════════════════════════════════════════════════════
    # ETAPA 7: Decisão e aplicação por valor (com locator em camadas)
    # ══════════════════════════════════════════════════════════════════
    current_state = state
    for tv in target_values:
        emap       = current_state.get("element_map", {})
        sel_before = list(current_state.get("selected_values", []))

        # ── Pergunta A: o valor existe? ──
        tv_exists = tv.lower() in effective_lower
        if not tv_exists:
            log.info(
                f"    [FILTER_VALUE_EXISTENCE] filter={slicer_title} "
                f"target_value='{tv}' exists_in_effective_available=False "
                f"action=error reason=value_not_in_any_source"
            )
            result["per_value"][tv] = {
                "action": "error", "click_ok": False,
                "reason": "value_not_in_any_source",
            }
            continue

        # ── Pergunta B: localizar elemento clicável ──
        # Primeiro: tenta element_map direto (leitura recente do DOM)
        entry      = emap.get(tv)
        matched_tv = tv
        if entry is None:
            tv_low      = tv.lower()
            matched_key = next((k for k in emap if k.lower() == tv_low), None)
            if matched_key:
                entry      = emap[matched_key]
                matched_tv = matched_key

        dom_locator_found = entry is not None
        locator_strategy  = "element_map_direct" if entry else "not_found_yet"

        log.info(
            f"    [FILTER_VALUE_LOCATOR] filter={slicer_title} "
            f"target_value='{tv}' "
            f"dom_locator_found={dom_locator_found} "
            f"locator_strategy={locator_strategy}"
        )

        # Se não encontrou no element_map, usa o locator em 3 camadas
        if entry is None:
            scroll_result = await _scroll_to_find_value_in_slicer(
                tab, idx, tv, slicer_title,
                enum_value_index=enum_value_index,
                enum_context=enum_context,
            )
            if scroll_result.get("found"):
                entry = {
                    "x": scroll_result["x"],
                    "y": scroll_result["y"],
                    "selected": scroll_result.get("selected", False),
                    "element_type": "value",
                }
                matched_tv = scroll_result.get("text", tv)
                dom_locator_found = True
                locator_strategy  = f"layered_{scroll_result.get('layer','?')}_{scroll_result.get('method','?')}"
                log.info(
                    f"    [FILTER_VALUE_LOCATOR] filter={slicer_title} "
                    f"target_value='{tv}' dom_locator_found=True "
                    f"locator_strategy={locator_strategy} "
                    f"coords=({entry['x']},{entry['y']})"
                )
            else:
                log.info(
                    f"    [FILTER_INCIDENT] filter={slicer_title} "
                    f"stage=locator "
                    f"reason=target_exists_but_dom_locator_failed "
                    f"target_value='{tv}' "
                    f"failure_reason={scroll_result.get('failure_reason', 'unknown')} "
                    f"failure_details={scroll_result.get('failure_details', '')}"
                )
                result["per_value"][tv] = {
                    "action": "error", "click_ok": False,
                    "reason": "exists_but_locator_failed",
                    "failure_reason": scroll_result.get("failure_reason", "unknown"),
                }
                continue

        # ── Verificar se item ainda não está selecionado antes de clicar ──
        currently_selected = bool(entry.get("selected", False))
        element_type       = entry.get("element_type", "value")

        if element_type != "value":
            log.info(
                f"    [FILTER_VALUE_DECISION] filter={slicer_title} "
                f"target_value='{tv}' exists=True "
                f"action=error reason=special_element_type={element_type}"
            )
            result["per_value"][tv] = {
                "action": "error", "click_ok": False,
                "reason": f"special_element_{element_type}",
            }
            continue

        if currently_selected:
            action = "skip"
            reason = "already_selected"
        else:
            action = "click_to_select"
            reason = "not_yet_selected"

        log.info(
            f"    [FILTER_VALUE_DECISION] filter={slicer_title} "
            f"target_value='{tv}' exists=True "
            f"currently_selected={currently_selected} "
            f"action={action} reason={reason}"
        )

        if action == "skip":
            result["per_value"][tv] = {
                "action": "skip", "click_ok": True,
                "selected_before": sel_before,
                "selected_after":  sel_before,
            }
            continue

        # ── Clique ──
        cx, cy   = entry["x"], entry["y"]
        click_ok = await _cdp_click(tab, cx, cy)
        await asyncio.sleep(0.5)

        new_state  = await read_current_selection(tab, idx, slicer_title)
        sel_after  = list(new_state.get("selected_values", []))

        log.info(
            f"    [FILTER_VALUE_APPLY] filter={slicer_title} "
            f"target_value='{tv}' click_target_text='{matched_tv}' "
            f"click_ok={click_ok} "
            f"selected_before={sel_before} "
            f"selected_after={sel_after}"
        )

        value_entered       = any(s.lower() == tv.lower() for s in sel_after)
        unexpected_removals = [
            v for v in sel_before
            if v not in sel_after and v.lower() != tv.lower()
        ]

        if not value_entered:
            log.warning(
                f"    [FILTER_INCIDENT] filter={slicer_title} "
                f"stage=apply_value target_value='{tv}' "
                f"reason=value_not_entered_after_click "
                f"selected_after={sel_after} action_taken=continue"
            )

        if unexpected_removals:
            log.warning(
                f"    [FILTER_INCIDENT] filter={slicer_title} "
                f"stage=apply_value target_value='{tv}' "
                f"reason=unexpected_removals={unexpected_removals} "
                f"action_taken=abort"
            )
            result["aborted"] = True
            result["per_value"][tv] = {
                "action": action, "click_ok": bool(click_ok),
                "selected_before": sel_before, "selected_after": sel_after,
                "unexpected_removals": unexpected_removals,
            }
            return result

        result["per_value"][tv] = {
            "action":   action, "click_ok": bool(click_ok),
            "selected_before": sel_before, "selected_after": sel_after,
            "unexpected_removals": [],
        }
        current_state = new_state

    # ══════════════════════════════════════════════════════════════════
    # ETAPA 8: Validação final
    # ══════════════════════════════════════════════════════════════════
    val_result = await validate_filter_final(
        tab, idx, slicer_title, target_values, mode
    )
    result["validation_ok"]  = val_result["validation_ok"]
    result["final_selected"] = val_result["final_selected"]

    log.info(f"    [FILTER_FINAL_VALIDATE]")
    log.info(f"    filter={slicer_title}")
    log.info(f"    target_values={target_values}")
    log.info(f"    final_selected={val_result.get('final_selected', [])}")
    log.info(f"    missing={val_result.get('missing', [])}")
    log.info(f"    validation_ok={val_result.get('validation_ok', False)}")

    return result

async def _read_filter_state(tab, idx: int, slicer_title: str, field_name: str) -> dict:
    """
    Lê o estado atual do slicer: valores disponíveis, selecionados,
    existência de controles especiais (Em branco, Selecionar tudo, Limpar).

    Retorna dict com:
      available_values, selected_values,
      has_blank, has_select_all, has_clear_control,
      element_map: {normalized_text -> {x, y, selected, element_type}}
    """
    import json as _j
    try:
        raw = await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return JSON.stringify({{error:'no_vc'}});

                const norm = t => (t || '').replace(/\\s+/g, ' ').trim();

                // ── Encontrar scroll container ──
                const allEls = Array.from(vc.querySelectorAll('*'));
                let sc = null; let bestScore = -1;
                for (const el of allEls) {{
                    const cs = getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    if (r.width < 20 || r.height < 20) continue;
                    const ov = (cs.overflowY || cs.overflow || '').toLowerCase();
                    const scrollable = el.scrollHeight > el.clientHeight + 2;
                    const hasOv = ov.includes('auto') || ov.includes('scroll');
                    if (!scrollable && !hasOv) continue;
                    const score = (scrollable ? 100 : 0) + (hasOv ? 50 : 0) + (el.scrollHeight - el.clientHeight);
                    if (score > bestScore) {{ bestScore = score; sc = el; }}
                }}
                if (!sc) sc = vc.querySelector('[role="listbox"]') || vc;

                // ── Varredura de itens ──
                const ITEM_SEL = '.slicerItemContainer, [role="option"], [role="listitem"], [class*="slicerItem"]';
                const items = Array.from(sc.querySelectorAll(ITEM_SEL)).filter(el => {{
                    const r = el.getBoundingClientRect();
                    return r.height > 0 && r.width > 0;
                }});

                const available = [];
                const selected  = [];
                const element_map = {{}};
                let has_blank      = false;
                let has_select_all = false;

                items.forEach(item => {{
                    const span = item.querySelector('.slicerText, span[class*="slicerText"], span');
                    const text = norm((span || item).textContent);
                    if (!text || text.length > 80) return;

                    const low = text.toLowerCase();
                    const isSelectAll = low === 'selecionar tudo' || low === 'select all';
                    const isBlank     = low === '(em branco)' || low === '(blank)';

                    const r = item.getBoundingClientRect();
                    const cx = Math.round(r.left + r.width / 2);
                    const cy = Math.round(r.top  + r.height / 2);

                    const isSel = (
                        item.classList.contains('selected') ||
                        item.classList.contains('isSelected') ||
                        item.getAttribute('aria-selected') === 'true' ||
                        item.getAttribute('aria-checked') === 'true' ||
                        !!item.querySelector('.selected, .isSelected, .partiallySelected')
                    );

                    const entry = {{
                        x: cx, y: cy,
                        selected: isSel,
                        element_type: isSelectAll ? 'select_all' : isBlank ? 'blank' : 'value',
                    }};
                    element_map[text] = entry;

                    if (isSelectAll) {{ has_select_all = true; if (isSel) selected.push(text); return; }}
                    if (isBlank)     {{ has_blank = true; }}

                    available.push(text);
                    if (isSel) selected.push(text);
                }});

                // ── Botão Limpar ──
                const clearSelectors = [
                    '.clear-button', '[class*="clearButton"]', '[class*="clear-button"]',
                    'button[title*="Limpar"]', 'button[title*="Clear"]',
                    '[aria-label*="Limpar"]', '[aria-label*="Clear"]',
                    '.slicer-header button', 'visual-container-header button',
                ];
                let hasClearControl = false;
                let clearX = 0, clearY = 0;
                for (const sel of clearSelectors) {{
                    const btn = vc.querySelector(sel);
                    if (btn) {{
                        const r = btn.getBoundingClientRect();
                        if (r.width > 0 && r.height > 0) {{
                            hasClearControl = true;
                            clearX = Math.round(r.left + r.width / 2);
                            clearY = Math.round(r.top  + r.height / 2);
                            break;
                        }}
                    }}
                }}

                return JSON.stringify({{
                    available_values: available,
                    selected_values: selected,
                    has_blank,
                    has_select_all,
                    has_clear_control: hasClearControl,
                    clear_x: clearX,
                    clear_y: clearY,
                    element_map,
                }});
            }})()
        """)
        data = _j.loads(str(raw)) if raw else {}
    except Exception as e:
        data = {"error": str(e), "available_values": [], "selected_values": [],
                "has_blank": False, "has_select_all": False, "has_clear_control": False,
                "element_map": {}}

    log.info(f"    [FILTER_STATE_BEFORE]")
    log.info(f"    filter={slicer_title}")
    log.info(f"    available_values={data.get('available_values', [])[:30]}")
    log.info(f"    selected_values={data.get('selected_values', [])}")
    log.info(f"    has_blank={data.get('has_blank', False)}")
    log.info(f"    has_select_all={data.get('has_select_all', False)}")
    log.info(f"    has_clear_control={data.get('has_clear_control', False)}")

    return data


async def _safe_clear_filter(tab, idx: int, slicer_title: str, state: dict) -> bool:
    """
    Limpa o slicer de forma segura.
    Prioridade:
      1. Botão Limpar explícito (mais seguro)
      2. Clicar em "Selecionar tudo" se marcado (toggle para desmarcar tudo)
      3. Desmarcar individualmente cada selecionado (último recurso)

    Em NENHUM caso faz clique cego em massa.
    """
    sels = state.get("selected_values", [])
    if not sels:
        log.info(f"    [FILTER_CLEAR] filter={slicer_title} method=nothing_to_clear clear_ok=True")
        return True

    # ── Método 1: botão Limpar ──
    if state.get("has_clear_control") and state.get("clear_x") and state.get("clear_y"):
        cx, cy = state["clear_x"], state["clear_y"]
        log.info(f"    [FILTER_CLEAR] filter={slicer_title} method=clear_button coords=({cx},{cy})")
        await _cdp_click(tab, cx, cy)
        await asyncio.sleep(0.6)

        # Valida limpeza
        new_state = await _read_filter_state(tab, idx, slicer_title, "")
        if not new_state.get("selected_values"):
            log.info(f"    [FILTER_CLEAR] filter={slicer_title} method=clear_button clear_ok=True")
            return True
        log.info(f"    [FILTER_CLEAR] filter={slicer_title} method=clear_button clear_ok=False remaining={new_state['selected_values']}")
        # Cai para método 2

    # ── Método 2: Selecionar Tudo (toggle) ──
    # Lógica: se alguns estão marcados mas não todos, clicar em "Selecionar Tudo"
    # uma vez marca todos; clicar de novo desmarca todos.
    # Só usamos se tiver "Selecionar Tudo" no mapa.
    emap = state.get("element_map", {})
    select_all_key = next(
        (k for k in emap if k.lower() in ("selecionar tudo", "select all")), None
    )
    if select_all_key and select_all_key in emap:
        entry = emap[select_all_key]
        cx, cy = entry["x"], entry["y"]
        log.info(f"    [FILTER_CLEAR] filter={slicer_title} method=select_all_toggle coords=({cx},{cy})")

        # Clica duas vezes: 1ª marca tudo, 2ª desmarca tudo
        for click_n in range(1, 3):
            await _cdp_click(tab, cx, cy)
            await asyncio.sleep(0.5)
            check_state = await _read_filter_state(tab, idx, slicer_title, "")
            if not check_state.get("selected_values"):
                log.info(f"    [FILTER_CLEAR] filter={slicer_title} method=select_all_toggle clicks={click_n} clear_ok=True")
                return True

        log.info(f"    [FILTER_CLEAR] filter={slicer_title} method=select_all_toggle clear_ok=False")

    # ── Método 3: desmarcar individualmente (último recurso) ──
    # NUNCA desmarca mais do que os selecionados conhecidos.
    current_state = await _read_filter_state(tab, idx, slicer_title, "")
    to_deselect = list(current_state.get("selected_values", []))
    log.info(f"    [FILTER_CLEAR] filter={slicer_title} method=individual_deselect count={len(to_deselect)}")

    for val in to_deselect:
        entry = current_state.get("element_map", {}).get(val)
        if not entry:
            log.info(f"    [FILTER_CLEAR] filter={slicer_title} value='{val}' coords=not_found — skip")
            continue
        if not entry.get("selected", False):
            log.info(f"    [FILTER_CLEAR] filter={slicer_title} value='{val}' already_deselected — skip")
            continue
        cx, cy = entry["x"], entry["y"]
        log.info(f"    [FILTER_CLEAR] filter={slicer_title} deselecting='{val}' coords=({cx},{cy})")
        await _cdp_click(tab, cx, cy)
        await asyncio.sleep(0.4)

    # Validação final
    final_state = await _read_filter_state(tab, idx, slicer_title, "")
    remaining = final_state.get("selected_values", [])
    clear_ok = len(remaining) == 0
    log.info(f"    [FILTER_CLEAR] filter={slicer_title} method=individual_deselect clear_ok={clear_ok} remaining={remaining}")
    return clear_ok


async def _apply_single_filter_value(
    tab, idx: int, slicer_title: str,
    target_value: str, current_state: dict,
    plan_mode: str,
) -> dict:
    """
    Aplica um único valor ao slicer com leitura de estado antes e depois.

    Retorna:
      action, click_ok, selected_before, selected_after, validation_ok, unexpected_removals
    """
    emap       = current_state.get("element_map", {})
    sel_before = list(current_state.get("selected_values", []))

    # ── Decisão ──
    entry = emap.get(target_value)

    if entry is None:
        # Tenta correspondência case-insensitive
        target_low = target_value.lower()
        matched_key = next((k for k in emap if k.lower() == target_low), None)
        if matched_key:
            entry = emap[matched_key]
            log.info(f"    [FILTER_VALUE_DECISION] filter={slicer_title} target_value='{target_value}' match_via_lowercase='{matched_key}'")
        else:
            log.info(
                f"    [FILTER_VALUE_DECISION] filter={slicer_title} "
                f"target_value='{target_value}' currently_selected=False "
                f"action=error reason=value_not_found_in_slicer"
            )
            return {
                "action": "error", "click_ok": False,
                "selected_before": sel_before, "selected_after": sel_before,
                "validation_ok": False, "unexpected_removals": [],
                "reason": "value_not_found",
            }

    currently_selected = bool(entry.get("selected", False))

    # Regra de ouro: NUNCA clicar sem saber o efeito
    if currently_selected:
        action = "skip"
        reason = "already_selected_skip_to_avoid_unintended_deselect"
    else:
        action = "click_to_select"
        reason = "not_yet_selected"

    log.info(
        f"    [FILTER_VALUE_DECISION] filter={slicer_title} "
        f"target_value='{target_value}' "
        f"currently_selected={currently_selected} "
        f"action={action} "
        f"reason={reason}"
    )

    if action == "skip":
        return {
            "action": "skip", "click_ok": True,
            "selected_before": sel_before, "selected_after": sel_before,
            "validation_ok": True, "unexpected_removals": [], "reason": reason,
        }

    # ── Executa clique ──
    cx, cy = entry["x"], entry["y"]
    log.info(f"    [FILTER_VALUE_APPLY] filter={slicer_title} target_value='{target_value}' coords=({cx},{cy})")
    click_ok = await _cdp_click(tab, cx, cy)
    await asyncio.sleep(0.5)

    # ── Releitura de estado ──
    new_state  = await _read_filter_state(tab, idx, slicer_title, "")
    sel_after  = list(new_state.get("selected_values", []))

    log.info(
        f"    [FILTER_VALUE_APPLY] filter={slicer_title} "
        f"target_value='{target_value}' "
        f"selected_before={sel_before} "
        f"selected_after={sel_after} "
        f"click_ok={click_ok}"
    )

    # ── Validação ──
    value_entered    = target_value in sel_after or any(s.lower() == target_value.lower() for s in sel_after)
    unexpected_removals = [v for v in sel_before if v not in sel_after and v != target_value]
    validation_ok    = value_entered and not unexpected_removals

    log.info(
        f"    [FILTER_VALUE_VALIDATE] filter={slicer_title} "
        f"target_value='{target_value}' "
        f"validation_ok={validation_ok} "
        f"unexpected_removals={unexpected_removals}"
    )

    return {
        "action": action, "click_ok": bool(click_ok),
        "selected_before": sel_before, "selected_after": sel_after,
        "validation_ok": validation_ok,
        "unexpected_removals": unexpected_removals,
        "reason": reason,
        "new_state": new_state,
    }


async def apply_filter_safe_control(
    tab,
    slicers: list,
    filter_plan: dict,
) -> dict:
    """
    Camada de Controle Seguro de Filtros.

    Recebe:
      - tab: aba do browser
      - slicers: lista retornada por scan_slicers (cada item tem index, title, allValues, etc.)
      - filter_plan: FILTER_PLAN declarado pelo usuário

    Fluxo por filtro:
      [FILTER_PLAN] → [FILTER_STATE_BEFORE] → [FILTER_CLEAR] →
      [FILTER_VALUE_DECISION] → [FILTER_VALUE_APPLY] → [FILTER_VALUE_VALIDATE] →
      [FILTER_FINAL_VALIDATE]

    Retorna dict com resultado por filtro.
    """
    results = {}

    if not filter_plan:
        log.info("    [FILTER_PLAN] filter_plan vazio — nenhum filtro a aplicar")
        return results

    # ── Indexar slicers por título normalizado ──
    slicer_by_title: dict[str, dict] = {}
    for s in slicers:
        raw_title  = s.get("title", "")
        norm_title = _normalize_slicer_title(raw_title)
        slicer_by_title[norm_title] = s

    # ── Log FILTER_PLAN ──
    plan_keys     = [_normalize_slicer_title(k) for k in filter_plan.keys()]
    available_keys = list(slicer_by_title.keys())
    matched       = [k for k in plan_keys if k in available_keys]
    ignored_plan  = [k for k in plan_keys if k not in available_keys]
    ignored_slice = [k for k in available_keys if k not in plan_keys]

    log.info(f"    [FILTER_PLAN] filters_declared={plan_keys}")
    log.info(f"    [FILTER_PLAN] filters_matched={matched}")
    log.info(f"    [FILTER_PLAN] filters_ignored_not_in_slicers={ignored_plan}")
    log.info(f"    [FILTER_PLAN] filters_ignored_not_in_plan={ignored_slice}")

    # ── Loop por filtro do plano ──
    for raw_plan_key, plan_cfg in filter_plan.items():
        norm_key     = _normalize_slicer_title(raw_plan_key)
        slicer_info  = slicer_by_title.get(norm_key)

        if slicer_info is None:
            log.info(f"    [FILTER_PLAN] '{norm_key}' não encontrado nos slicers — ignorando")
            results[norm_key] = {"ok": False, "reason": "slicer_not_found"}
            continue

        idx          = slicer_info["index"]
        slicer_title = slicer_info.get("title", raw_plan_key)
        field_name   = norm_key
        mode         = plan_cfg.get("mode", "single")
        clear_first  = plan_cfg.get("clear_first", True)
        target_values = list(plan_cfg.get("target_values", []))

        log.info(f"\n  ── Aplicando filtro '{slicer_title}' (container #{idx}) ──")
        log.info(f"    mode={mode} clear_first={clear_first} target_values={target_values}")

        filter_result = {
            "ok": False, "mode": mode,
            "target_values": target_values,
            "final_selected": [],
            "per_value": {},
        }

        # ── 1. Ativar slicer (clique seguro no cabeçalho) ──
        header_pos = await _click_slicer_header_safe(tab, idx)
        if header_pos.get("ok"):
            await _cdp_click(tab, header_pos["x"], header_pos["y"])
            await asyncio.sleep(0.5)
        else:
            log.info(f"    ⚠️ Cabeçalho não encontrado para '{slicer_title}' — tentando sem ativação explícita")

        # ── 2. Leitura de estado inicial ──
        state = await _read_filter_state(tab, idx, slicer_title, field_name)

        # ── 3. Limpeza (se clear_first) ──
        if clear_first:
            clear_ok = await _safe_clear_filter(tab, idx, slicer_title, state)
            if not clear_ok:
                log.warning(f"    ⚠️ Limpeza não confirmada para '{slicer_title}' — prosseguindo com cautela")
            # Reler estado após limpeza
            state = await _read_filter_state(tab, idx, slicer_title, field_name)

        # ── 4-5. Decisão e aplicação por valor ──
        current_state = state
        for tv in target_values:
            res = await _apply_single_filter_value(
                tab, idx, slicer_title, tv, current_state, mode
            )
            filter_result["per_value"][tv] = res

            # Atualiza estado atual para próxima iteração
            if "new_state" in res:
                current_state = res["new_state"]
            else:
                # Reler se não veio atualizado
                current_state = await _read_filter_state(tab, idx, slicer_title, field_name)

        # ── 6. Validação final ──
        final_state    = await _read_filter_state(tab, idx, slicer_title, field_name)
        final_selected = final_state.get("selected_values", [])

        if mode == "single":
            # Espera exatamente 1 valor selecionado, igual ao target
            validation_ok = (
                len(target_values) == 1 and
                any(s.lower() == target_values[0].lower() for s in final_selected)
            )
        else:
            # multi: todos os target_values devem estar presentes
            target_set = {v.lower() for v in target_values}
            final_set  = {s.lower() for s in final_selected}
            validation_ok = target_set.issubset(final_set)

        log.info(f"    [FILTER_FINAL_VALIDATE]")
        log.info(f"    filter={slicer_title}")
        log.info(f"    target_values={target_values}")
        log.info(f"    final_selected_values={final_selected}")
        log.info(f"    validation_ok={validation_ok}")

        filter_result["ok"]            = validation_ok
        filter_result["final_selected"] = final_selected
        results[norm_key]              = filter_result

        # Fecha com ESC para não deixar dropdown aberto
        await press_escape(tab, times=2, wait_each=0.3)
        await asyncio.sleep(0.4)

    return results
# ---------------------------------------------------------------------------
# Scan de filtros/slicers
# ---------------------------------------------------------------------------

async def scan_slicers(tab):
    """
    Escaneia filtros/slicers na página.
    v10.2 — Micro-scroll safe enumeration:
      - Keyboard probe executado ANTES do loop (não rouba foco)
      - Ativação via clique no cabeçalho (zona segura)
      - Enumeração via micro-scroll incremental (NUNCA clica em valores)
      - Guarda de seleção a cada passo (para se detectar side-effect)
      - Comparação de caixas consecutivas para descobrir novos valores
    """
    log.info("🎚️ Escaneando filtros/slicers (com expansão de dropdowns)...")
    await get_report_dom_context_info(tab, caller="scan_slicers (pré-scan)")
    await cleanup_residual_ui(tab, stage_label="scan de slicers - início", aggressive=True)

    # 1. Scan estático
    payload_raw = await tab.evaluate(r"""
        (() => {
            try {
                const results = [];
                const resolveDoc = () => {
                    if (document.querySelectorAll('visual-container').length > 0)
                        return { doc: document, source: 'document' };
                    for (const fr of document.querySelectorAll('iframe')) {
                        try {
                            const d = fr.contentDocument;
                            if (d && d.querySelectorAll('visual-container').length > 0)
                                return { doc: d, source: 'iframe' };
                        } catch(e) {}
                    }
                    return { doc: document, source: 'document' };
                };
                const ctx = resolveDoc();
                const norm = t => (t || '').replace(/\s+/g, ' ').trim();
                const containers = Array.from(ctx.doc.querySelectorAll('visual-container'));
                let rawSlicer = 0, discardedBySize = 0;

                containers.forEach((vc, index) => {
                    const el = vc.querySelector('transform') || vc;
                    const rect = el.getBoundingClientRect();
                    if (rect.width < 15 || rect.height < 15) { discardedBySize++; return; }

                    const allClasses = [vc.className||'', el.className||'',
                        ...Array.from(vc.querySelectorAll('[class]')).slice(0,40).map(e=>String(e.className||''))
                    ].join(' ').toLowerCase();

                    const isSlicer = (
                        allClasses.includes('slicer') || allClasses.includes('chiclet') ||
                        !!vc.querySelector('.slicer-container,.slicer-content-wrapper,.slicer-body,.slicerBody,[class*="Slicer"],[class*="slicer"],.chiclet-slicer')
                    );
                    if (!isSlicer) return;
                    rawSlicer++;

                    let title = '';
                    const headerEl = vc.querySelector('.slicer-header-text,[class*="header-text"],h3,h4,.visual-title,.visualTitle,[class*="title"]');
                    if (headerEl) title = norm(headerEl.textContent);
                    if (!title) {
                        const aria = vc.getAttribute('aria-label') || el.getAttribute('aria-label') || '';
                        title = norm(aria.replace(/\(.*?\)/g, ''));
                    }

                    let slicerType = 'lista';
                    if (allClasses.includes('chiclet')) slicerType = 'chiclet';
                    else if (vc.querySelector('input[type="range"],.slider,[class*="range"],[class*="Range"]')) slicerType = 'range';
                    else if (vc.querySelector('select,.dropdown,[class*="dropdown"],[class*="Dropdown"]')) slicerType = 'dropdown';
                    else if (vc.querySelector('.date-slicer,[class*="date-slicer"],[class*="DateSlicer"]')) slicerType = 'date';
                    else if (vc.querySelector('input[type="text"],input.searchInput')) slicerType = 'busca';

                    const allValues = [], selectedValues = [];
                    const seen = new Set(), selectedSet = new Set();

                    const items = vc.querySelectorAll(
                        '.slicerItemContainer,[class*="slicerItem"],[role="option"],[role="listitem"],.row,[class*="chiclet"]'
                    );
                    items.forEach(item => {
                        const value = norm(
                            item.querySelector('.slicerText,span,[class*="slicerText"],[class*="text"],label')?.textContent || item.textContent
                        );
                        if (!value || value.length > 80) return;
                        const lower = value.toLowerCase();
                        if (lower === 'selecionar tudo' || lower === 'select all') return;
                        if (!seen.has(value)) { seen.add(value); allValues.push(value); }

                        const checkbox = item.querySelector('.slicerCheckbox,input[type="checkbox"],[class*="checkbox"],[class*="Checkbox"]');
                        const isSelected = (
                            item.classList.contains('selected') || item.classList.contains('isSelected') ||
                            item.querySelector('.selected,.isSelected,.partiallySelected') !== null ||
                            item.getAttribute('aria-selected') === 'true' ||
                            item.getAttribute('aria-checked') === 'true' ||
                            (checkbox && (checkbox.checked === true ||
                                checkbox.getAttribute('aria-checked') === 'true' ||
                                checkbox.classList.contains('selected') ||
                                checkbox.classList.contains('partiallySelected')))
                        );
                        if (isSelected && !selectedSet.has(value)) { selectedSet.add(value); selectedValues.push(value); }
                    });

                    const hasSearchInput = !!vc.querySelector('input[type="text"],input[class*="search"],input[class*="Search"]');
                    const hasDropdownBtn = !!vc.querySelector('[class*="dropdown"],[class*="Dropdown"],button[class*="expand"],button[class*="Expand"]');
                    const needsExpansion = (
                        (slicerType === 'busca' || slicerType === 'dropdown') && (hasSearchInput || hasDropdownBtn || allValues.length === 0)
                    ) || (slicerType === 'chiclet' && allValues.length === 0);

                    const visibleText = norm(vc.textContent).toLowerCase();
                    const hasPending = ['ainda não aplicado','not yet applied','apply changes','aplicar alterações'].some(t => visibleText.includes(t));
                    const inferredFiltered = (
                        selectedValues.length > 0 ||
                        ['múltiplos selecionados','multiple selections','selecionado','selected'].some(t => visibleText.includes(t)) ||
                        vc.querySelector('[aria-selected="true"],[aria-checked="true"],.selected,.isSelected,.partiallySelected') !== null
                    );

                    results.push({
                        index,
                        title: title || `Slicer #${index + 1}`,
                        type: slicerType,
                        allValues: allValues.slice(0, 50),
                        selectedValues: selectedValues.slice(0, 50),
                        totalValues: allValues.length,
                        totalSelected: selectedValues.length,
                        hasPending,
                        applied: inferredFiltered && !hasPending,
                        needsExpansion,
                        x: Math.round(rect.x),
                        y: Math.round(rect.y),
                        width: Math.round(rect.width),
                        height: Math.round(rect.height),
                    });
                });

                return JSON.stringify({
                    ok: true,
                    payload: {
                        slicers: results,
                        diagnostics: { rawSlicer, discardedBySize, contextSource: ctx.source }
                    }
                });
            } catch(err) {
                return JSON.stringify({ ok: false, error: String(err && err.message ? err.message : err) });
            }
        })()
    """)

    try:
        wrapper = json.loads(str(payload_raw)) if payload_raw else {}
    except Exception as exc:
        log.error(f"❌ [scan_slicers] parse error: {exc}")
        return []

    if not wrapper.get("ok"):
        log.error(f"❌ [scan_slicers] JS error: {wrapper.get('error')}")
        return []

    payload = wrapper.get("payload") or {}
    slicers = list(payload.get("slicers") or [])
    diag = payload.get("diagnostics") or {}
    log.info(
        f"🧪 [diagnóstico] visual-container bruto no DOM: "
        f"{diag.get('rawSlicer', 0) + diag.get('discardedBySize', 0)} "
        f"(contexto={diag.get('contextSource', '?')})"
    )
    log.info(f"🧪 [diagnóstico] candidatos brutos a slicer: {diag.get('rawSlicer', 0)}")
    log.info(f"🧪 [diagnóstico] descartados por tamanho (slicer scan): {diag.get('discardedBySize', 0)}")

    # ══════════════════════════════════════════════════════════════════════
    # KEYBOARD PROBE: executa UMA VEZ ANTES do loop de slicers
    # Isso evita que o probe roube o foco durante a interação com slicers
    # ══════════════════════════════════════════════════════════════════════
    probe = await _simple_keyboard_probe(tab)
    log.info(
        f"  [KEYBOARD_PROBE] ok={probe.get('ok')} "
        f"enter_ok={probe.get('enter_ok')} arrow_ok={probe.get('arrow_ok')} "
        f"focused={probe.get('focused')}"
    )
    if probe.get("error"):
        log.info(f"  [KEYBOARD_PROBE] error={probe.get('error')}")

    # Limpa qualquer estado residual do probe
    await press_escape(tab, times=2, wait_each=0.3)
    await asyncio.sleep(0.5)

    import json as _j

    for slicer in slicers:
        idx = slicer["index"]
        title = slicer.get("title", f"Slicer #{idx}")
        field_name = title.replace("(Ainda não aplicado)", "").strip().lower()
        slicer_type = slicer.get("type", "lista")
        log.info(f"  🔽 Slicer '{title}' (container #{idx})...")

        # ── Resolve posição real ─────────────────────────────────────────
        try:
            pos_raw = await tab.evaluate(f"""
                (() => {{
                    const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                    if (!vc) return JSON.stringify({{error:'no_vc'}});
                    const candidates = [vc, ...Array.from(vc.children)];
                    let bestRect = null;
                    for (const el of candidates) {{
                        const r = el.getBoundingClientRect();
                        if (r.width > 20 && r.height > 20) {{ bestRect = r; break; }}
                    }}
                    if (!bestRect) {{
                        for (const el of Array.from(vc.querySelectorAll('*'))) {{
                            const r = el.getBoundingClientRect();
                            if (r.width > 20 && r.height > 20 && r.left > 0 && r.top > 0) {{ bestRect = r; break; }}
                        }}
                    }}
                    if (!bestRect) return JSON.stringify({{error: 'no_valid_rect'}});
                    const cx = Math.round(bestRect.left + bestRect.width / 2);
                    const cy = Math.round(bestRect.top  + bestRect.height / 2);
                    const inp = vc.querySelector('input[type="text"],input[class*="search"]');
                    let inpX = cx, inpY = cy;
                    if (inp) {{
                        const ir = inp.getBoundingClientRect();
                        if (ir.width > 0 && ir.height > 0) {{
                            inpX = Math.round(ir.left + ir.width / 2);
                            inpY = Math.round(ir.top  + ir.height / 2);
                        }}
                    }}
                    return JSON.stringify({{
                        ok: true, method: 'direct_rect',
                        vcW: Math.round(bestRect.width), vcH: Math.round(bestRect.height),
                        screenX: cx, screenY: cy,
                        inp_found: !!inp, inpX, inpY,
                    }});
                }})()
            """)
            pos = _j.loads(str(pos_raw)) if isinstance(pos_raw, str) else {}
        except Exception as e:
            pos = {"error": str(e)}

        log.info(f"    [POS] {pos}")

        if pos.get("error") or not pos.get("ok"):
            log.info(f"    ⚠️ posição não resolvida: {pos.get('error')} — pulando")
            continue

        # ── Chiclet: lê botões diretamente ─────────────────────────────
        if slicer_type == "chiclet":
            log.info("    [CHICLET] lendo valores diretamente dos botões...")
            try:
                chiclet_raw = await tab.evaluate(f"""
                    (() => {{
                        const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                        if (!vc) return JSON.stringify({{values: [], selected: []}});
                        const btns = Array.from(vc.querySelectorAll(
                            'button, .chiclet, [class*="chiclet"], [role="option"], [class*="slicerItem"], .slicerItemContainer'
                        )).filter(b => {{
                            const r = b.getBoundingClientRect();
                            return r.width > 10 && r.height > 10;
                        }});
                        const values = [], selected = [], seen = new Set();
                        btns.forEach(b => {{
                            const t = (b.textContent || '').trim();
                            if (!t || t.length > 80) return;
                            const lower = t.toLowerCase();
                            if (lower === 'selecionar tudo' || lower === 'select all') return;
                            if (!seen.has(t)) {{ seen.add(t); values.push(t); }}
                            const isSel = (
                                b.classList.contains('selected') || b.classList.contains('isSelected') ||
                                b.getAttribute('aria-selected') === 'true' || b.getAttribute('aria-pressed') === 'true' ||
                                b.getAttribute('aria-checked') === 'true'
                            );
                            if (isSel && !selected.includes(t)) selected.push(t);
                        }});
                        return JSON.stringify({{values, selected}});
                    }})()
                """)
                chiclet_data = _j.loads(str(chiclet_raw)) if chiclet_raw else {}
                chiclet_vals = chiclet_data.get("values", [])
                chiclet_sel = chiclet_data.get("selected", [])

                if chiclet_vals:
                    slicer["allValues"] = chiclet_vals
                    slicer["totalValues"] = len(chiclet_vals)
                    slicer["selectedValues"] = chiclet_sel
                    slicer["totalSelected"] = len(chiclet_sel)
                    if chiclet_sel:
                        slicer["applied"] = True
                    log.info(f"    ✅ [chiclet] {len(chiclet_vals)} valores: {chiclet_vals} | sel: {chiclet_sel}")
                else:
                    log.info("    ⚠️ [chiclet] nenhum valor nos botões — tentando ativar visual")
                    await _experiment_activate_slicer(tab, idx, title)

            except Exception as che:
                log.info(f"    ⚠️ [chiclet] leitura falhou: {che}")

            await press_escape(tab, times=2, wait_each=0.3)
            await asyncio.sleep(0.5)
            continue

        # ══════════════════════════════════════════════════════════════════
        # BUSCA/DROPDOWN: enumeração via micro-scroll seguro
        # Estratégia: "ativar uma vez, depois nunca mais clicar em item;
        #              apenas micro-rolar e observar"
        # ══════════════════════════════════════════════════════════════════

        # Primeiro: experimento para descobrir qual clique ativa o slicer
        await _experiment_activate_slicer(tab, idx, title)

        # Limpa estado dos experimentos (ESC para sair de qualquer modo)
        await press_escape(tab, times=3, wait_each=0.3)
        await asyncio.sleep(0.5)

        # Agora: enumeração via micro-scroll
        enum_result = await _enumerate_slicer_via_micro_scroll(tab, idx, title, field_name)

        if enum_result.get("selection_side_effect"):
            log.warning(f"    ⚠️ INCIDENTE: seleção alterada durante enumeração de '{title}'!")

        if enum_result.get("success") and enum_result.get("values"):
            scroll_values = enum_result["values"]
            scroll_selected = enum_result.get("selected", [])

            # Mescla com valores estáticos
            static_vals = set(slicer.get("allValues") or [])
            merged = list(
                dict.fromkeys(
                    scroll_values + [v for v in sorted(static_vals) if v not in set(scroll_values)]
                )
            )

            slicer["allValues"] = merged
            slicer["totalValues"] = len(merged)
            slicer["selectedValues"] = scroll_selected
            slicer["totalSelected"] = len(scroll_selected)
            if scroll_selected:
                slicer["applied"] = True

            log.info(
                f"    ✅ {len(merged)} valores (scroll={len(scroll_values)}, merged): "
                f"{merged[:20]} | sel: {scroll_selected}"
            )
        else:
            # Fallback: usa valores do scan estático
            static_vals = slicer.get("allValues") or []
            if static_vals:
                log.info(
                    f"    ⚠️ Enumeração por scroll falhou — mantendo "
                    f"{len(static_vals)} valores estáticos: {static_vals}"
                )
            else:
                # Último recurso: coleta passiva do DOM
                log.info("    ⚠️ Sem valores do scroll ou estáticos — coletando DOM passivo...")
                try:
                    val_raw = await tab.evaluate(f"""
                        (() => {{
                            const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                            if (!vc) return JSON.stringify({{texts:[]}});
                            const texts = [];
                            vc.querySelectorAll('*').forEach(el => {{
                                if (el.childNodes.length === 1 && el.childNodes[0].nodeType === 3) {{
                                    const t = el.childNodes[0].textContent.trim();
                                    if (t && t.length < 80) texts.push(t);
                                }}
                            }});
                            return JSON.stringify({{texts: [...new Set(texts)]}});
                        }})()
                    """)
                    val = _j.loads(str(val_raw)) if val_raw else {}
                    all_texts = val.get("texts", [])
                    values_clean = [
                        v for v in all_texts
                        if not _is_slicer_noise(v, field_name)
                    ]
                    if values_clean:
                        slicer["allValues"] = values_clean
                        slicer["totalValues"] = len(values_clean)
                        log.info(f"    ✅ {len(values_clean)} valores (DOM passivo): {values_clean}")
                except Exception as ve:
                    log.info(f"    ⚠️ Coleta passiva falhou: {ve}")

        # ══════════════════════════════════════════════════════════════════
        # TRANSIÇÃO FASE 1 → FASE 2
        # Após enumeração completa deste slicer, verificar se há plano
        # de filtros declarado e aplicá-lo antes de avançar ao próximo.
        # ══════════════════════════════════════════════════════════════════
        final_clean_values = list(slicer.get("allValues") or [])
        normalized_title = normalize_slicer_name(title)

        log.info(
            f"    [FILTER_NAME_NORMALIZATION] "
            f"display_name='{title}' "
            f"normalized='{normalized_title}' "
            f"matched_key={'yes' if normalized_title in FILTER_PLAN else 'no'}"
        )

        if normalized_title in FILTER_PLAN:
            log.info(
                f"    [FILTER_CONTROL_START] filter={normalized_title} "
                f"available_values={final_clean_values} "
                f"plan_found=True"
            )
            plan_cfg = FILTER_PLAN[normalized_title]
            enum_vi = enum_result.get("value_index", {}) if isinstance(enum_result, dict) else {}
            enum_ctx = enum_result.get("enum_context", {}) if isinstance(enum_result, dict) else {}
            log.info(
                f"    [ENUM_CONTEXT_SNAPSHOT] filter={normalized_title} "
                f"best_selector={enum_ctx.get('best_selector', '?')} "
                f"initial_scrollTop={enum_ctx.get('initial_scrollTop', '?')}"
            )
            filter_result = await apply_filter_plan(
                tab, idx, title, final_clean_values, plan_cfg,
                enum_value_index=enum_vi,
                enum_context=enum_ctx,
            )

            if filter_result.get("validation_ok"):
                log.info(
                    f"    ✅ [FILTER_PHASE2_OK] filter={normalized_title} "
                    f"final_selected={filter_result.get('final_selected', [])}"
                )
                # Atualiza slicer com seleção real aplicada
                slicer["selectedValues"] = filter_result.get("final_selected", [])
                slicer["totalSelected"] = len(slicer["selectedValues"])
                if slicer["selectedValues"]:
                    slicer["applied"] = True
            else:
                if filter_result.get("aborted"):
                    log.warning(
                        f"    [FILTER_INCIDENT] filter={normalized_title} "
                        f"stage=phase2 reason=apply_aborted action_taken=abort"
                    )
                else:
                    log.warning(
                        f"    ❌ [FILTER_PHASE2_FAIL] filter={normalized_title} "
                        f"final_selected={filter_result.get('final_selected', [])} "
                        f"validation_ok=False"
                    )
        else:
            log.info(
                f"    [FILTER_CONTROL_SKIP] filter={normalized_title} "
                f"reason=not_in_plan"
            )

        # Fecha com ESC
        await press_escape(tab, times=3, wait_each=0.3)
        await asyncio.sleep(0.5)

    slicers.sort(key=lambda s: (s.get("y", 0), s.get("x", 0)))
    await cleanup_residual_ui(tab, stage_label="scan de slicers - final", aggressive=True)
    return slicers

async def get_current_url(tab) -> str:
    try:
        result = await tab.evaluate("window.location.href")
        return str(result) if result else ""
    except Exception:
        return ""


async def is_powerbi_loaded(tab) -> bool:
    try:
        count = await tab.evaluate("document.querySelectorAll('visual-container').length")
        return count is not None and int(str(count)) > 0
    except Exception:
        return False


async def is_on_correct_url(tab, target_url: str) -> bool:
    current = await get_current_url(tab)
    if not current:
        return False
    target_lower = target_url.lower()
    current_lower = current.lower()
    if current_lower.startswith(target_lower[:60]):
        return True
    if "reportid=" in target_lower:
        report_id = target_lower.split("reportid=")[1].split("&")[0]
        if report_id in current_lower:
            return True
    if "powerbi.com" in current_lower and "reportid" in current_lower:
        return True
    return False


async def find_correct_tab(browser, target_url: str):
    try:
        targets = await browser.get_targets()
        if not targets:
            return None
        for target in targets:
            t_url = str(getattr(target, 'url', '') or '').lower()
            if 'powerbi.com' in t_url or 'reportembed' in t_url:
                try:
                    tab = await browser.get_tab(target)
                    if tab:
                        return tab
                except Exception:
                    continue
    except Exception:
        pass
    return None


async def ensure_correct_page(browser, tab, target_url: str, owned_tab_refs: set[str] | None = None):
    MAX_RETRIES = 5
    RETRY_WAIT = 8

    for attempt in range(1, MAX_RETRIES + 1):
        log.info(f"🔄 Verificação de página (tentativa {attempt}/{MAX_RETRIES})...")
        current_url = await get_current_url(tab)
        log.info(f"  📍 URL atual: {current_url[:100]}")
        on_correct = await is_on_correct_url(tab, target_url)

        if on_correct:
            log.info("  ✅ URL correta!")
            log.info("  ⏳ Aguardando Power BI renderizar...")
            for load_attempt in range(8):
                if await is_powerbi_loaded(tab):
                    log.info(f"  ✅ Power BI carregado!")
                    return tab
                log.info(f"    Tentativa {load_attempt + 1}/8: aguardando visual-containers...")
                await asyncio.sleep(MEDIUM_WAIT)
            log.warning("  ⚠️ URL correta mas Power BI não renderizou, recarregando...")
            await tab.evaluate("window.location.reload()")
            await asyncio.sleep(PAGE_LOAD_WAIT)
            continue

        log.warning(f"  ⚠️ URL incorreta!")
        log.info("  🔍 Procurando aba do Power BI em outras tabs...")
        pbi_tab = await find_correct_tab(browser, target_url)
        if pbi_tab:
            log.info("  ✅ Encontrou aba do Power BI!")
            tab = pbi_tab
            await asyncio.sleep(MEDIUM_WAIT)
            continue

        log.info(f"  🔄 Forçando navegação para URL do Power BI...")
        await tab.evaluate(f'window.location.href = "{target_url}"')
        await asyncio.sleep(PAGE_LOAD_WAIT)

        if await is_on_correct_url(tab, target_url):
            log.info("  ✅ Navegação forçada funcionou!")
            for load_attempt in range(6):
                if await is_powerbi_loaded(tab):
                    log.info("  ✅ Power BI carregado!")
                    return tab
                await asyncio.sleep(MEDIUM_WAIT)
        else:
            log.warning("  ⚠️ Navegação falhou, abrindo nova aba...")
            try:
                tab = await browser.get(target_url, new_tab=True)
                if owned_tab_refs is not None:
                    owned_tab_refs.add(_tab_ref(tab))
                await asyncio.sleep(PAGE_LOAD_WAIT)
            except Exception as e:
                log.error(f"  ❌ Erro ao abrir nova aba: {e}")

    log.warning("⚠️ Não foi possível confirmar carregamento completo.")
    return tab


# ---------------------------------------------------------------------------
# Fluxo principal
# ---------------------------------------------------------------------------

async def navigate_to_report_page(tab, target_page: str) -> bool:
    if not target_page:
        return True

    target_norm = target_page.strip().lower()
    log.info(f"📄 Tentando navegar para página interna do relatório: '{target_page}'")

    async def _get_page_signature():
        context_info = await get_report_dom_context_info(tab, caller="navigate_signature")
        try:
            raw = await tab.evaluate(r"""
                (() => {
                    const textOf = (el) => (el?.innerText || el?.textContent || '').replace(/\s+/g, ' ').trim();
                    const norm = (s) => (s || '').toLowerCase();
                    let selectedTab = '';
                    const selectedCandidates = Array.from(document.querySelectorAll(
                        '[role="tab"][aria-selected="true"], li.section.active, li[class*="active"]'
                    ));
                    for (const el of selectedCandidates) {
                        const t = textOf(el);
                        if (t && t.length <= 80) { selectedTab = t; break; }
                    }
                    const visualCount = document.querySelectorAll('visual-container').length;
                    return JSON.stringify({ selectedTab, visualCount });
                })()
            """)
            parsed = json.loads(str(raw)) if raw else {}
            return {
                "selected_tab": str(parsed.get("selectedTab") or "").strip(),
                "visual_count": int(context_info.get("visual_count", 0)),
                "context_source": str(context_info.get("context_source") or "unknown"),
            }
        except Exception:
            return {"selected_tab": "", "visual_count": 0}

    async def _legacy_click_page():
        try:
            result = await tab.evaluate(f"""
                (() => {{
                    const name = {json.dumps(target_page)}.toUpperCase();
                    const navSelectors = ['li.section', 'li[class*="section"]', 'ul.pane li', '[role="tab"]', '[role="listitem"]'];
                    for (const sel of navSelectors) {{
                        for (const item of document.querySelectorAll(sel)) {{
                            const text = item.textContent?.trim();
                            if (text && text.toUpperCase().includes(name)) {{
                                const target = item.querySelector('span, div, a') || item;
                                target.scrollIntoView({{block: 'center'}});
                                target.click();
                                return `Aba: "${{text.substring(0, 50)}}"`;
                            }}
                        }}
                    }}
                    for (const el of document.querySelectorAll('span, a, button, label, h3, h4')) {{
                        const text = el.textContent?.trim();
                        if (text && text.toUpperCase().includes(name) && text.length < 40) {{
                            const rect = el.getBoundingClientRect();
                            if (rect.width > 5 && rect.height > 5) {{
                                el.scrollIntoView({{block: 'center'}});
                                el.click();
                                return `Texto: "${{text.substring(0, 40)}}"`;
                            }}
                        }}
                    }}
                    const ariaEl = document.querySelector(`[aria-label*="${{name}}"]`) || document.querySelector(`[title*="${{name}}"]`);
                    if (ariaEl) {{ ariaEl.click(); return "aria-label/title"; }}
                    return null;
                }})()
            """)
        except Exception:
            result = None
        clicked_detail = str(result or "").strip()
        return bool(clicked_detail), clicked_detail

    before = await _get_page_signature()
    log.info(f"  ℹ️ Estado antes da navegação: selectedTab='{before.get('selected_tab','')}', visuals={before.get('visual_count', 0)} (contexto={before.get('context_source','unknown')})")

    await close_open_menus_and_overlays(tab, aggressive=True)
    await asyncio.sleep(1)

    for attempt in range(1, 6):
        log.info(f"  🔎 Tentativa {attempt}/5 de navegar para '{target_page}' (estratégia legada v05)")
        clicked, clicked_detail = await _legacy_click_page()

        if not clicked:
            log.warning(f"  ⚠️ Tentativa {attempt}/5 sem clique válido na página '{target_page}'.")
            await asyncio.sleep(1.5)
            continue

        log.info(f"  ✅ Clique realizado: {clicked_detail}")
        log.info("  ⏳ Aguardando renderização pós-clique...")
        await asyncio.sleep(LONG_WAIT)

        loaded = await wait_for_visual_containers(tab, retries=6, wait_seconds=3)
        after = await _get_page_signature()

        selected_now = str(after.get("selected_tab", "")).lower()
        selected_matches_target = bool(selected_now and target_norm in selected_now)
        changed_signature = (
            after.get("selected_tab") != before.get("selected_tab")
            or after.get("visual_count") != before.get("visual_count")
        )

        if loaded and (selected_matches_target or changed_signature):
            log.info(f"  ✅ Navegação confirmada para '{target_page}': selectedTab='{after.get('selected_tab','')}', visuals={after.get('visual_count', 0)} (contexto={after.get('context_source','unknown')})")
            return True

        log.warning(f"  ⚠️ Clique ocorreu, mas sem confirmação sólida da troca de página.")
        await asyncio.sleep(1.5)

    log.error(f"❌ Não foi possível navegar para a página '{target_page}'.")
    return False

async def wait_for_visual_containers(tab, retries: int = 10, wait_seconds: int = 3) -> bool:
    for attempt in range(1, retries + 1):
        context_info = await get_report_dom_context_info(tab, caller=f"wait_for_visual_containers tentativa {attempt}/{retries}")
        count = int(context_info.get("visual_count", 0))
        context_source = str(context_info.get("context_source") or "unknown")
        if count > 0:
            log.info(f"✅ Visual-containers detectados: {count} (contexto={context_source})")
            return True
        log.info(f"⏳ Aguardando visuais renderizarem... tentativa {attempt}/{retries}")
        await asyncio.sleep(wait_seconds)
    return False


async def get_report_dom_context_info(tab, caller: str = ""):
    try:
        raw = await tab.evaluate("""
            (() => {
                const result = { contextSource: 'document', visualCount: 0, frameIndex: -1 };
                try {
                    const mainCount = document.querySelectorAll('visual-container').length;
                    if (mainCount > 0) {
                        result.contextSource = 'document';
                        result.visualCount = mainCount;
                        return JSON.stringify(result);
                    }
                } catch (e) {}
                const iframes = Array.from(document.querySelectorAll('iframe'));
                for (let i = 0; i < iframes.length; i++) {
                    try {
                        const d = iframes[i].contentDocument;
                        if (!d) continue;
                        const c = d.querySelectorAll('visual-container').length;
                        if (c > 0) {
                            result.contextSource = `iframe[${i}]`;
                            result.visualCount = c;
                            result.frameIndex = i;
                            return JSON.stringify(result);
                        }
                    } catch (e) {}
                }
                return JSON.stringify(result);
            })()
        """)
        parsed = json.loads(str(raw)) if raw else {}
        info = {
            "context_source": str(parsed.get("contextSource") or "document"),
            "visual_count": int(parsed.get("visualCount") or 0),
            "frame_index": int(parsed.get("frameIndex") or -1),
        }
    except Exception:
        info = {"context_source": "unavailable", "visual_count": 0, "frame_index": -1}

    if caller:
        log.info(f"🧪 [dom-context] {caller}: source={info['context_source']}, visuals={info['visual_count']}, frame={info['frame_index']}")
    return info


async def wait_for_visuals_or_abort(tab, stage_label: str, retries: int = 10, wait_seconds: int = 3, allow_reload: bool = True) -> bool:
    log.info(f"⏳ Aguardando visual-containers ({stage_label})...")
    loaded = await wait_for_visual_containers(tab, retries=retries, wait_seconds=wait_seconds)
    if loaded:
        return True
    log.warning(f"⚠️ Nenhum visual-container detectado em '{stage_label}'.")
    if allow_reload:
        log.warning(f"🔄 Tentando reload controlado em '{stage_label}'...")
        try:
            await tab.evaluate("window.location.reload()")
            await asyncio.sleep(PAGE_LOAD_WAIT)
        except Exception:
            pass
        loaded_after_reload = await wait_for_visual_containers(tab, retries=max(6, retries // 2), wait_seconds=wait_seconds)
        if loaded_after_reload:
            return True
    log.error(f"❌ Abortando: página sem visual-containers em '{stage_label}'.")
    return False


def _display_slicers_inline(slicers: list):
    print("\n" + "=" * 100)
    print("🎚️  FILTROS / SLICERS ENCONTRADOS")
    print("=" * 100)

    if not slicers:
        print("  Nenhum slicer/filtro identificado.")
        print("=" * 100)
        return

    for i, s in enumerate(slicers):
        raw_title = s.get("title", f"Slicer #{i+1}")
        title = raw_title.replace('(Ainda não aplicado)', '').strip()
        stype = s.get("type", "?")
        applied = s.get("applied", False)
        pending = s.get("hasPending", False)
        idx = s.get("index", i)

        if applied:
            status = "✅ APLICADO"
        elif pending:
            status = "⏳ PENDENTE (não aplicado)"
        else:
            status = "⬜ SEM SELEÇÃO"

        print(f"\n  [{i}] {title}  (tipo: {stype}, container #{idx})  {status}")
        print(f"      Tamanho: {s.get('width', '?')}x{s.get('height', '?')}")

        all_values = s.get("allValues") or []
        selected_values = s.get("selectedValues") or []

        if all_values:
            exibir = all_values[:10]
            sufixo = f"  (+{len(all_values)-10} mais)" if len(all_values) > 10 else ""
            print(f"      Valores disponíveis ({s.get('totalValues', len(all_values))}): {', '.join(exibir)}{sufixo}")
        else:
            print("      Valores disponíveis: não detectados")

        if selected_values:
            print(f"      ► Valores APLICADOS ({len(selected_values)}): {', '.join(selected_values[:10])}")
        elif applied:
            print("      ► Filtro aplicado (valor exato não detectado)")
        else:
            print("      ► Nenhum valor selecionado / todos selecionados")

    print("\n" + "-" * 100)
    aplicados = sum(1 for s in slicers if s.get("applied"))
    print(f"  Total: {len(slicers)} filtros  |  {aplicados} aplicado(s)")
    print("=" * 100)


async def run_export(url: str, browser_path: str, target_page: str, stop_after_filters: bool = False):
    log.info("🚀 Iniciando automação Power BI")
    log.info(f"📎 URL: {url}")
    log.info(f"🌐 Navegador: {browser_path}")
    if target_page:
        log.info(f"📄 Página alvo: {target_page}")

    browser_path = normalize_browser_path(browser_path)

    if not validate_runtime_config():
        return False

    browser, _profile_dir = await start_isolated_browser(browser_path)
    owned_tab_refs: set[str] = set()
    tab = None

    try:
        tab = await open_report_tab(browser, url, owned_tab_refs)
        focus_ok = await ensure_focus_visibility_emulation(tab, stage_label="post_open_tab")
        log.info("[FOCUS_POLICY]")
        log.info("  stage=post_open_tab")
        log.info("  mode=dom_emulation_first")
        log.info(f"  safe_focus_called={not focus_ok}")
        log.info("  fallback_only=True")
        log.info(f"  reason={'emulation_ok' if focus_ok else 'emulation_failed'}")
        if not focus_ok:
            log.warning("[FOCUS_FALLBACK_USED]")
            log.warning("  stage=post_open_tab")
            log.warning("  reason=focus_emulation_failed")
            await safe_focus_tab(tab)
            focus_ok = await ensure_focus_visibility_emulation(tab, stage_label="post_open_tab_after_fallback")
            if not focus_ok:
                log.error("❌ Falha ao estabilizar foco/visibilidade após fallback.")
                return False

        await close_extra_tabs_created_by_script(browser, tab, owned_tab_refs)
        await allow_multiple_downloads(tab)

        log.info(f"⏳ Aguardando carregamento inicial ({PAGE_LOAD_WAIT}s)...")
        await asyncio.sleep(PAGE_LOAD_WAIT)

        await block_microsoft_learn_and_external_links(tab)
        await cleanup_residual_ui(tab, stage_label="preparação geral após carregamento", aggressive=True)
        tab = await ensure_report_tab_still_valid(tab, url)

        try:
            current_url = await get_tab_url(tab)
            if current_url and "powerbi.com" not in current_url.lower():
                log.warning("⚠️ URL atual não parece ser do relatório. Reabrindo...")
                await tab.get(url)
                await asyncio.sleep(LONG_WAIT)
        except Exception:
            pass

        initial_loaded = await wait_for_visuals_or_abort(tab, stage_label="carregamento inicial do relatório", retries=10, wait_seconds=3, allow_reload=True)
        if not initial_loaded:
            return False

        if target_page:
            page_ok = await navigate_to_report_page(tab, target_page)
            if not page_ok:
                log.error(f"❌ Fluxo interrompido: a navegação para '{target_page}' não foi confirmada.")
                return False
            target_loaded = await wait_for_visuals_or_abort(tab, stage_label=f"navegação para TARGET_PAGE='{target_page}'", retries=8, wait_seconds=3, allow_reload=True)
            if not target_loaded:
                return False
            await ensure_focus_visibility_emulation(tab, stage_label="post_navigation_target_page")
        else:
            ready_without_target = await wait_for_visuals_or_abort(tab, stage_label="página inicial (sem TARGET_PAGE)", retries=8, wait_seconds=3, allow_reload=True)
            if not ready_without_target:
                return False

        await close_open_menus_and_overlays(tab, aggressive=True)

        try:
            current_visual_count = await tab.evaluate("document.querySelectorAll('visual-container').length")
            current_visual_count = int(current_visual_count or 0)
        except Exception:
            current_visual_count = 0

        if current_visual_count == 0:
            log.error("❌ Nenhum visual-container encontrado antes do scan. Abortando.")
            return False

        print("\n" + "=" * 70)
        print("📌 ETAPA 1 - LEITURA DE FILTROS")
        print("=" * 70)

        pre_filters_ok = await ensure_focus_visibility_emulation(tab, stage_label="pre-filters")
        log.info("[FOCUS_POLICY]")
        log.info("  stage=pre-filters")
        log.info("  mode=dom_emulation_first")
        log.info("  safe_focus_called=False")
        log.info("  fallback_only=True")
        log.info(f"  reason={'emulation_ok' if pre_filters_ok else 'emulation_not_fully_confirmed'}")

        await cleanup_residual_ui(tab, stage_label="antes do scan de slicers", aggressive=True)
        slicers = await scan_slicers(tab)
        # ══════════════════════════════════════════════════════════════
        # POST-FILTER GATE — decisão de continuar para exportação
        # ══════════════════════════════════════════════════════════════

#        if stop_after_filters:
#            log.info("🛑 stop_after_filters=True — encerrando após leitura de filtros.")
#            return True

        # ── Coletar resultado dos filtros aplicados ──
        filters_expected = list(FILTER_PLAN.keys()) if FILTER_PLAN else []
        filters_ok = []
        filters_failed = []
        critical_incidents = []

        for s in slicers:
            norm_title = normalize_slicer_name(s.get("title", ""))
            if norm_title not in FILTER_PLAN:
                continue
            plan_cfg = FILTER_PLAN[norm_title]
            target_set = {v.lower() for v in plan_cfg.get("target_values", [])}
            selected_set = {v.lower() for v in (s.get("selectedValues") or [])}
            is_ok = target_set.issubset(selected_set)
            is_required = plan_cfg.get("required", True)

            if is_ok:
                filters_ok.append(norm_title)
            else:
                filters_failed.append(norm_title)
                if is_required:
                    critical_incidents.append(
                        f"required_filter_failed:{norm_title} "
                        f"expected={list(target_set)} got={list(selected_set)}"
                    )

        # ── POST_FILTER_CLEANUP ──
        log.info("🧹 Limpeza pós-filtros...")
        await press_escape(tab, times=3, wait_each=0.3)
        await asyncio.sleep(0.5)
        await cleanup_residual_ui(tab, stage_label="post_filter_cleanup", aggressive=True)
        await asyncio.sleep(0.5)

        try:
            overlay_count = await tab.evaluate("""
                (() => {
                    const visible = (el) => {
                        const r = el.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    };
                    return Array.from(document.querySelectorAll(
                        '[role="menu"], [role="dialog"], [role="listbox"], ' +
                        '[aria-modal="true"], .cdk-overlay-pane, .contextMenu, .popup, .modal'
                    )).filter(visible).length;
                })()
            """)
            overlay_count = int(overlay_count or 0)
        except Exception:
            overlay_count = -1

        ui_clean = overlay_count == 0

        log.info(f"    [POST_FILTER_CLEANUP]")
        log.info(f"    escape_sent=True")
        log.info(f"    cleanup_done=True")
        log.info(f"    visible_overlays={overlay_count}")
        log.info(f"    wait_after_filters_ms=1000")

        # ── POST_FILTER_RENDER_WAIT ──
        log.info("⏳ Aguardando estabilização do relatório pós-filtros...")
        await asyncio.sleep(3)

        try:
            visual_count = await tab.evaluate(
                "document.querySelectorAll('visual-container').length"
            )
            visual_count = int(visual_count or 0)
        except Exception:
            visual_count = 0

        report_stable = visual_count > 0

        log.info(f"    [POST_FILTER_RENDER_WAIT]")
        log.info(f"    wait_ms=3000")
        log.info(f"    visual_count={visual_count}")
        log.info(f"    report_stable={report_stable}")

        # ── POST_FILTER_GATE — decisão final ──
        export_release = (
            len(critical_incidents) == 0
            and ui_clean
            and report_stable
        )

        log.info(f"    [POST_FILTER_GATE]")
        log.info(f"    filters_expected={filters_expected}")
        log.info(f"    filters_ok={filters_ok}")
        log.info(f"    filters_failed={filters_failed}")
        log.info(f"    critical_incidents={critical_incidents}")
        log.info(f"    ui_clean={ui_clean}")
        log.info(f"    report_stable={report_stable}")
        log.info(f"    export_release={export_release}")

        if not export_release:
            log.error(
                f"    [EXPORT_ABORT] "
                f"reason={'required_filter_validation_failed' if critical_incidents else 'ui_or_render_issue'} "
                f"critical_incidents={critical_incidents}"
            )
            return False

        log.info(
            f"    [EXPORT_PHASE_START] "
            f"release_from_gate=True "
            f"filters_applied={filters_ok}"
        )

        log.info("======================================================================")
        log.info("📌 Escaneando visuais disponíveis")
        log.info("======================================================================")
        pre_export_ok = await ensure_focus_visibility_emulation(tab, stage_label="pre-export")
        log.info("[FOCUS_POLICY]")
        log.info("  stage=pre-export")
        log.info("  mode=dom_emulation_first")
        log.info("  safe_focus_called=False")
        log.info("  fallback_only=True")
        log.info(f"  reason={'emulation_ok' if pre_export_ok else 'emulation_not_fully_confirmed'}")

        await cleanup_residual_ui(tab, stage_label="antes do scan de visuais", aggressive=True)
        visuals = await scan_visuals(tab)
        display_visuals(visuals)

        exportable = [i for i, v in enumerate(visuals) if v.get("hasExportData")]
        if not exportable:
            log.warning("⚠️ Nenhum visual com 'Exportar dados' confirmado.")
            return False

        selected_indexes = await ask_user_visual_selection(visuals)
        if not selected_indexes:
            log.warning("⛔ Exportação cancelada pelo usuário.")
            return False

        log.info(f"📥 {len(selected_indexes)} visuais selecionados para exportação")
        results = await export_selected_visuals(tab, visuals, selected_indexes)
        display_export_summary(results)
        return True

    finally:
        await graceful_browser_shutdown(tab, browser)


async def graceful_browser_shutdown(tab, browser):
    if tab is not None:
        with contextlib.suppress(Exception):
            await dismiss_sensitive_data_popup(tab)
        with contextlib.suppress(Exception):
            await close_open_menus_and_overlays(tab, aggressive=True)
        with contextlib.suppress(Exception):
            await asyncio.sleep(0.3)
        with contextlib.suppress(BrokenPipeError, ConnectionResetError, RuntimeError, OSError, Exception):
            await tab.close()
        with contextlib.suppress(Exception):
            await asyncio.sleep(0.2)
    if browser is not None:
        with contextlib.suppress(BrokenPipeError, ConnectionResetError, RuntimeError, OSError, Exception):
            browser.stop()
        with contextlib.suppress(Exception):
            await asyncio.sleep(0.2)


def main():
    if not validate_runtime_config():
        sys.exit(1)
    try:
        success = asyncio.run(run_export(POWERBI_URL, BROWSER_PATH, TARGET_PAGE, stop_after_filters=True))
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        log.warning("⛔ Execução interrompida pelo usuário")
        sys.exit(130)


if __name__ == "__main__":
    main()
