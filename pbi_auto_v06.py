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
    return [
        "--no-sandbox",
        "--disable-blink-features=AutomationControlled",
        "--disable-infobars",
        "--start-maximized",
        "--lang=pt-BR",
        "--disable-popup-blocking",
    ]


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

    aggressive=False:
        limpeza normal entre etapas

    aggressive=True:
        limpeza mais forte, usada antes/depois de exportação ou quando algo ficou preso
    """
    # ESC já resolve bastante coisa no Power BI
    await press_escape(tab, times=3 if aggressive else 2, wait_each=0.4)

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

                const clickableSelectors = [
                    'button',
                    '[role="button"]',
                    '[aria-label]',
                    '.close',
                    '.close-btn',
                    '.close-button',
                    '.dialog-close',
                    '.modal-close',
                    '.popup-close',
                    '[data-testid*="close"]',
                    '[class*="close"]',
                    '[title]',
                ];

                const candidates = Array.from(document.querySelectorAll(clickableSelectors.join(',')));

                for (const el of candidates) {
                    const txt = textOf(el);
                    const aria = (el.getAttribute('aria-label') || '').trim().toLowerCase();
                    const title = (el.getAttribute('title') || '').trim().toLowerCase();

                    if (isLearnLink(el)) {
                        // Nunca clicar nesses links
                        continue;
                    }

                    const rect = el.getBoundingClientRect();
                    const visible = rect.width > 0 && rect.height > 0;
                    if (!visible) continue;

                    const shouldClose =
                        txt === 'x' ||
                        aria === 'x' ||
                        title === 'x' ||
                        txt.includes('fechar') ||
                        txt.includes('close') ||
                        aria.includes('fechar') ||
                        aria.includes('close') ||
                        title.includes('fechar') ||
                        title.includes('close') ||
                        txt.includes('cancelar') ||
                        txt.includes('cancel') ||
                        aria.includes('cancelar') ||
                        aria.includes('cancel');

                    if (shouldClose) {
                        try { el.click(); } catch (e) {}
                    }
                }
            })()
        """)
    except Exception:
        pass

    await asyncio.sleep(OVERLAY_SETTLE_WAIT if aggressive else 1.0)
    await press_escape(tab, times=2 if aggressive else 1, wait_each=0.3)
    await asyncio.sleep(0.5)


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


async def dismiss_sensitive_data_popup(tab, max_rounds: int = 6) -> bool:
    """
    Fecha o popup de 'Você está copiando dados confidenciais' ou semelhantes,
    sem clicar no link do Microsoft Learn.

    Estratégia:
    - procura dialog/modal visível
    - tenta clicar em Fechar/X/Cancelar
    - se houver botão tipo 'Copiar/Copy', ele NÃO deve ficar acionando em loop
    - jamais clicar em links 'Saiba mais...' / Learn
    """
    handled_any = False

    for _ in range(max_rounds):
        try:
            result = await tab.evaluate("""
                (() => {
                    const textOf = (el) => (el?.innerText || el?.textContent || '').trim();
                    const lower = (s) => (s || '').toLowerCase();

                    const dialogs = Array.from(document.querySelectorAll([
                        '[role="dialog"]',
                        '[aria-modal="true"]',
                        '.modal',
                        '.popup',
                        '.dialog',
                    ].join(',')));

                    const visibleDialogs = dialogs.filter(d => {
                        const r = d.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    });

                    const allCandidates = [];

                    for (const dlg of visibleDialogs) {
                        const dlgText = lower(textOf(dlg));
                        const isSensitive =
                            dlgText.includes('confidencial') ||
                            dlgText.includes('confidential') ||
                            dlgText.includes('copiando dados') ||
                            dlgText.includes('copying data') ||
                            dlgText.includes('exportar dados') ||
                            dlgText.includes('exporting data');

                        if (!isSensitive) continue;

                        // Não inclui <a> para impedir clique em "Saiba mais"/Learn.
                        const els = dlg.querySelectorAll('button, [role="button"], [aria-label], [title]');
                        for (const el of els) {
                            const txt = lower(textOf(el));
                            const aria = lower(el.getAttribute('aria-label') || '');
                            const title = lower(el.getAttribute('title') || '');
                            const rect = el.getBoundingClientRect();
                            const visible = rect.width > 0 && rect.height > 0;
                            if (!visible) continue;

                            const isLearn =
                                txt.includes('saiba mais sobre como exportar dados') ||
                                txt.includes('learn more about exporting data') ||
                                txt.includes('saiba mais') ||
                                txt.includes('learn more');

                            if (isLearn) continue;

                            allCandidates.push({
                                el,
                                txt,
                                aria,
                                title
                            });
                        }
                    }

                    // Prioridade 1: fechar/cancelar/x
                    for (const item of allCandidates) {
                        const { el, txt, aria, title } = item;
                        const shouldClose =
                            txt === 'x' ||
                            aria === 'x' ||
                            title === 'x' ||
                            txt.includes('fechar') ||
                            txt.includes('close') ||
                            txt.includes('cancelar') ||
                            txt.includes('cancel') ||
                            aria.includes('fechar') ||
                            aria.includes('close') ||
                            aria.includes('cancelar') ||
                            aria.includes('cancel') ||
                            title.includes('fechar') ||
                            title.includes('close');

                        if (shouldClose) {
                            try { el.click(); } catch (e) {}
                            return "closed";
                        }
                    }

                    // Prioridade 2: se não achou botão claro, tenta X genérico visível
                    for (const item of allCandidates) {
                        const { el, txt, aria, title } = item;
                        if (txt === '×' || txt === 'x' || aria === 'x' || title === 'x') {
                            try { el.click(); } catch (e) {}
                            return "closed-x";
                        }
                    }

                    return "";
                })()
            """)
        except Exception:
            result = ""

        if result:
            handled_any = True
            log.info(f"  ⚠️🔒 Popup confidencial tratado com ação segura: {result}")
            await asyncio.sleep(1.2)
            await press_escape(tab, times=1, wait_each=0.3)
            continue

        # Se não achou nada explícito, tenta limpeza genérica
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
    Rola o visual para o centro da viewport antes de interagir.
    Isso aumenta muito a chance de o header ser renderizado corretamente.
    """
    idx = visual.get("index")
    try:
        await tab.evaluate(f"""
            (() => {{
                const vcs = Array.from(document.querySelectorAll('visual-container'));
                const vc = vcs[{idx}];
                if (!vc) return false;
                vc.scrollIntoView({{ behavior: 'auto', block: 'center', inline: 'center' }});
                return true;
            }})()
        """)
    except Exception:
        pass

    await asyncio.sleep(MEDIUM_WAIT)


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
    Tenta localizar o botão 'Mais opções' do visual atual.
    Retorna um dict com coordenadas ou None.
    """
    idx = visual.get("index")

    try:
        raw = await tab.evaluate(f"""
            (() => {{
                const vcs = Array.from(document.querySelectorAll('visual-container'));
                const vc = vcs[{idx}];
                if (!vc) return null;

                const candidates = Array.from(vc.querySelectorAll([
                    'button[aria-label*="opções"]',
                    'button[aria-label*="Opções"]',
                    'button[aria-label*="options"]',
                    'button[aria-label*="Options"]',
                    '.visual-header-item-container button',
                    '[class*="visualHeader"] button',
                    'button',
                    '[role="button"]'
                ].join(',')));

                const visible = candidates.filter(el => {{
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                }});

                const score = (el) => {{
                    const txt = ((el.innerText || el.textContent || '')).trim().toLowerCase();
                    const aria = ((el.getAttribute('aria-label') || '')).trim().toLowerCase();
                    const title = ((el.getAttribute('title') || '')).trim().toLowerCase();

                    let s = 0;
                    if (aria.includes('opções') || aria.includes('options')) s += 8;
                    if (title.includes('opções') || title.includes('options')) s += 6;
                    if (txt.includes('...')) s += 3;
                    if ((el.closest('.visual-header-item-container'))) s += 4;
                    return s;
                }};

                visible.sort((a, b) => score(b) - score(a));

                const btn = visible[0];
                if (!btn) return null;

                const r = btn.getBoundingClientRect();

                return JSON.stringify({{
                    x: Math.round(r.left + (r.width / 2)),
                    y: Math.round(r.top + (r.height / 2)),
                    width: Math.round(r.width),
                    height: Math.round(r.height)
                }});
            }})()
        """)
        return json.loads(str(raw)) if raw else None
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
    """Verifica rapidamente se algum menu de contexto/menuitem visível foi aberto."""
    try:
        opened = await tab.evaluate("""
            (() => {
                const isVisible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };
                const roots = Array.from(document.querySelectorAll(
                    '[role="menu"], .contextMenu, .dropdown-menu, .menu'
                )).filter(isVisible);
                for (const root of roots) {
                    if (root.querySelector('[role="menuitem"], [role="menuitemcheckbox"], li, button')) {
                        return true;
                    }
                }
                return false;
            })()
        """)
        return bool(opened)
    except Exception:
        return False


async def open_more_options_robust(tab, visual, attempt: int, retries: int) -> bool:
    """
    Estratégia robusta para abrir "Mais opções" sem clicar em elemento errado.

    Etapas:
    1) scroll do visual para centro
    2) hover no centro
    3) hover no header
    4) localizar botão por seletor/score e clicar por coordenada
    5) fallback por clique JS somente dentro do visual
    """
    idx = visual.get("index")
    title = visual.get("title", f"Visual #{idx}")

    log.info(f"    🎯 [{attempt}/{retries}] Abrindo 'Mais opções' em {title} (container #{idx})")

    await scroll_visual_into_view(tab, visual)
    await hover_visual_center(tab, visual)
    await force_hover_visual_header(tab, visual)

    coords = await locate_more_options_button(tab, visual)
    if coords:
        log.info(
            f"      • botão localizado em x={coords.get('x')} y={coords.get('y')} "
            f"(w={coords.get('width')} h={coords.get('height')})"
        )
        clicked = await click_more_options_button(tab, visual)
        if clicked:
            await asyncio.sleep(0.8)
            if await is_visual_menu_open(tab):
                log.info("      ✅ menu aberto via clique por coordenada")
                return True
            log.info("      ⚠️ clique enviado, mas menu ainda não ficou visível")
    else:
        log.info("      ⚠️ botão não localizado por coordenadas nesta tentativa")

    # fallback: click JS somente no container atual, sem tocar fora dele
    try:
        fallback = await tab.evaluate(f"""
            (() => {{
                const vc = Array.from(document.querySelectorAll('visual-container'))[{idx}];
                if (!vc) return false;
                const selectors = [
                    'button[aria-label*="Mais opções"]',
                    'button[aria-label*="More options"]',
                    'button[title*="Mais opções"]',
                    'button[title*="More options"]',
                    '.visual-header-item-container button',
                    '[class*="visualHeader"] button'
                ];
                for (const sel of selectors) {{
                    const btn = vc.querySelector(sel);
                    if (!btn) continue;
                    const r = btn.getBoundingClientRect();
                    if (r.width <= 0 || r.height <= 0) continue;
                    btn.scrollIntoView({{block: 'center', inline: 'center'}});
                    btn.dispatchEvent(new MouseEvent('mouseover', {{bubbles: true}}));
                    btn.click();
                    return true;
                }}
                return false;
            }})()
        """)
    except Exception:
        fallback = False

    if fallback:
        await asyncio.sleep(0.8)
        if await is_visual_menu_open(tab):
            log.info("      ✅ menu aberto via fallback JS restrito ao visual")
            return True
        log.info("      ⚠️ fallback JS clicou, mas menu não ficou visível")

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


async def scan_visuals(tab):
    """
    Escaneia visuais disponíveis na página e tenta confirmar quais suportam exportação.

    Regras:
    - limpa overlays antes de cada teste
    - centraliza o visual na tela
    - faz hover real
    - tenta abrir 'Mais opções'
    - valida se o menu aberto realmente é do visual
    """
    log.info("🔎 Escaneando visuais na página (base v5)...")
    await close_open_menus_and_overlays(tab, aggressive=True)

    payload = await eval_json(tab, """
        (() => {
            const results = [];
            const discardReasons = {};
            const incDiscard = (reason) => {
                discardReasons[reason] = (discardReasons[reason] || 0) + 1;
            };

            const containers = Array.from(document.querySelectorAll('visual-container'));
            const rawCount = containers.length;

            containers.forEach((vc, index) => {
                const el = vc.querySelector('transform') || vc;
                if (!el) { incDiscard('no_host_element'); return; }

                const rect = el.getBoundingClientRect();
                if (rect.width < 15 || rect.height < 15) {
                    incDiscard('tiny_rect(<15x15)');
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
                else if (allClasses.includes('card')) type = 'Card';
                else if (allClasses.includes('kpi')) type = 'KPI';
                else if (allClasses.includes('map')) type = 'Mapa';
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
                    width: Math.round(rect.width),
                    height: Math.round(rect.height),
                    x: Math.round(rect.left),
                    y: Math.round(rect.top),
                    hasOptionsButton: !!optionsBtn,
                    hasHeader: hasHeader,
                    menuOpened: false,
                    hasExportData: false,
                    exportReason: "nao_verificado",
                    rawText: (el.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 250)
                });
            });

            const withHeaderOrBtn = results.filter(v => v.hasHeader || v.hasOptionsButton).length;
            return {
                visuals: results,
                diagnostics: {
                    rawCount,
                    keptCount: results.length,
                    withHeaderOrBtn,
                    discardedCount: rawCount - results.length,
                    discardReasons
                }
            };
        })()
    """, default={"visuals": [], "diagnostics": {}})

    visuals = list((payload or {}).get("visuals") or [])
    diagnostics = (payload or {}).get("diagnostics") or {}

    raw_count = int(diagnostics.get("rawCount", 0))
    kept_count = int(diagnostics.get("keptCount", len(visuals)))
    discarded_count = int(diagnostics.get("discardedCount", max(0, raw_count - kept_count)))
    discard_reasons = diagnostics.get("discardReasons", {}) or {}

    log.info(f"🧪 [diagnóstico] visual-container bruto no DOM: {raw_count}")
    log.info(f"🧪 [diagnóstico] mantidos no scan: {kept_count} | descartados: {discarded_count}")
    if discard_reasons:
        for reason, qty in discard_reasons.items():
            log.info(f"🧪 [diagnóstico] descarte visual: {reason} = {qty}")

    if raw_count > 0 and kept_count == 0:
        log.warning("⚠️ Diagnóstico: havia visual-container bruto, mas todos foram descartados pelos filtros.")

    # comportamento v5: prioriza visuais com header/botão; se zerar, usa todos os mantidos
    exportable = [v for v in visuals if v.get("hasHeader") or v.get("hasOptionsButton")]
    if not exportable and visuals:
        log.warning("⚠️ Nenhum visual com header/botão detectado; usando lista completa (fallback v5).")
        exportable = visuals

    exportable.sort(key=lambda v: (0 if v.get('hasOptionsButton') else 1, v.get('y', 0), v.get('x', 0)))

    log.info(f"📊 Total de visual-containers: {len(exportable)}")
    log.info(f"📋 Visuais com header de opções: {sum(1 for v in exportable if v.get('hasOptionsButton'))}")
    log.info("🔬 Verificando quais visuais suportam 'Exportar dados'...")

    for visual in exportable:
        await close_open_menus_and_overlays(tab, aggressive=True)
        await scroll_visual_into_view(tab, visual)
        await hover_visual_center(tab, visual)

        probe = await probe_visual_export_status(tab, visual, retries=EXPORT_PROBE_RETRIES)
        visual["menuOpened"] = bool(probe.get("menu_opened"))
        visual["hasExportData"] = bool(probe.get("has_export"))
        visual["exportReason"] = str(probe.get("reason") or "indefinido")

        await close_open_menus_and_overlays(tab, aggressive=True)
        await asyncio.sleep(1)

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


async def select_export_type(tab) -> bool:
    """
    Seleciona a opção de exportação.
    Preferência:
    1. .xlsx / Excel
    2. primeira opção de rádio visível
    """
    log.info("  🔘 Selecionando tipo de exportação...")

    try:
        result = await tab.evaluate("""
            (() => {
                const textOf = (el) => (el?.innerText || el?.textContent || '').trim().toLowerCase();
                const isVisible = (el) => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                };

                const dialogs = Array.from(document.querySelectorAll([
                    '[role="dialog"]',
                    '[aria-modal="true"]',
                    '.modal',
                    '.popup',
                    '.dialog'
                ].join(','))).filter(isVisible);

                for (const dlg of dialogs) {
                    const buttons = Array.from(dlg.querySelectorAll('button, [role="radio"], input[type="radio"], label'));
                    for (const el of buttons) {
                        if (!isVisible(el)) continue;
                        const txt = textOf(el);

                        if (
                            txt.includes('.xlsx') ||
                            txt.includes('excel') ||
                            txt.includes('dados resumidos') ||
                            txt.includes('data with current layout')
                        ) {
                            try { el.click(); } catch (e) {}
                            return "preferred";
                        }
                    }

                    const radios = Array.from(dlg.querySelectorAll('[role="radio"], input[type="radio"]')).filter(isVisible);
                    if (radios.length > 0) {
                        try { radios[0].click(); } catch (e) {}
                        return "radio";
                    }
                }

                return "";
            })()
        """)
    except Exception:
        result = ""

    if result:
        log.info("  ✅ Radio button")
        await asyncio.sleep(SHORT_WAIT)
        return True

    # fallback do fluxo antigo: seletores específicos do diálogo de exportação
    for js_expr in [
        'document.querySelector("#pbi-radio-button-1 > label > section > div")',
        'document.querySelector("#pbi-radio-button-1 label")',
        'document.querySelector("#pbi-radio-button-1")',
        'document.querySelector("input[type=radio]:not([disabled])")'
    ]:
        if await js_click(tab, js_expr, "Radio button (fallback)"):
            await asyncio.sleep(SHORT_WAIT)
            return True

    log.warning("  ⚠️ Não foi possível selecionar o tipo de exportação")
    return False


async def confirm_export_dialog(tab) -> bool:
    """
    Clica no botão final 'Exportar' do diálogo.
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

                const dialogs = Array.from(document.querySelectorAll([
                    '[role="dialog"]',
                    '[aria-modal="true"]',
                    '.modal',
                    '.popup',
                    '.dialog'
                ].join(','))).filter(isVisible);

                for (const dlg of dialogs) {
                    const buttons = Array.from(dlg.querySelectorAll('button, [role="button"]')).filter(isVisible);

                    for (const el of buttons) {
                        const txt = textOf(el);

                        const isBad =
                            txt.includes('saiba mais') ||
                            txt.includes('learn more') ||
                            txt === 'cancelar' ||
                            txt === 'cancel';

                        if (isBad) continue;

                        if (txt === 'exportar' || txt.includes('exportar')) {
                            try { el.click(); } catch (e) {}
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
        log.info("  ✅ Botão Exportar")
        await asyncio.sleep(1)
        return True

    # fallback do fluxo antigo para variações de DOM do diálogo
    for js_expr in [
        'document.querySelector("export-data-dialog button.exportButton")',
        'document.querySelector("button.exportButton")',
        'document.querySelector("button.primaryBtn.exportButton")',
        'document.querySelector("mat-dialog-actions button.primaryBtn")',
        'document.querySelector("mat-dialog-actions button:first-child")'
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

    # trata popup confidencial, se aparecer
    await dismiss_sensitive_data_popup(tab)
    tab = await ensure_report_tab_still_valid(tab, POWERBI_URL)

    confirmed = False
    for dialog_attempt in range(1, 3):
        dialog_ready = await wait_export_dialog(tab, retries=6 if dialog_attempt == 1 else 4)
        if not dialog_ready:
            await dismiss_sensitive_data_popup(tab)
            log.info(f"  ⚠️ Tentativa {dialog_attempt}/2 sem diálogo de exportação pronto.")
            continue

        await dismiss_sensitive_data_popup(tab)
        await select_export_type(tab)
        await dismiss_sensitive_data_popup(tab)

        confirmed = await confirm_export_dialog(tab)
        if not confirmed:
            await dismiss_sensitive_data_popup(tab)
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

    # limpa popup final que porventura tenha sobrado
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
    Além de OK/Continuar, cobre o fluxo que pede "Copiar/Copy"
    e depois exige fechar um popup pelo X.
    """
    handled_any = False

    for _ in range(6):
        result = await tab.evaluate("""
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
# Scan de filtros/slicers
# ---------------------------------------------------------------------------

async def scan_slicers(tab):
    """
    Escaneia filtros/slicers visíveis na página sem deixar overlays abertos no final.

    Objetivos:
    - identificar slicers com mais estabilidade
    - detectar melhor se aparentam ter filtro aplicado
    - não deixar dropdown, tooltip ou painel aberto após a leitura
    """
    log.info("🎚️ Escaneando filtros/slicers (base v5)...")
    await cleanup_residual_ui(tab, stage_label="scan de slicers - início", aggressive=True)

    payload = await eval_json(tab, """
        (() => {
            const results = [];
            const containers = Array.from(document.querySelectorAll('visual-container'));
            const norm = txt => (txt || '').replace(/\\s+/g, ' ').trim();
            const diagnostics = {
                rawVisualContainers: containers.length,
                rawSlicerCandidates: 0,
                discardedBySize: 0
            };

            containers.forEach((vc, index) => {
                const el = vc.querySelector('transform') || vc;
                const rect = el.getBoundingClientRect();
                if (rect.width < 15 || rect.height < 15) {
                    diagnostics.discardedBySize += 1;
                    return;
                }

                const allClasses = [
                    vc.className || '',
                    el.className || '',
                    ...Array.from(vc.querySelectorAll('[class]')).slice(0, 40).map(e => String(e.className || ''))
                ].join(' ').toLowerCase();

                const isSlicer = (
                    allClasses.includes('slicer') ||
                    allClasses.includes('chiclet') ||
                    !!vc.querySelector('.slicer-container, .slicer-content-wrapper, .slicer-body, .slicerBody, [class*="Slicer"], [class*="slicer"], .chiclet-slicer')
                );
                if (!isSlicer) return;
                diagnostics.rawSlicerCandidates += 1;

                let title = '';
                const headerEl = vc.querySelector(
                    '.slicer-header-text, [class*="header-text"], h3, h4, .visual-title, .visualTitle, [class*="title"]'
                );
                if (headerEl) title = norm(headerEl.textContent);
                if (!title) {
                    const ariaLabel = vc.getAttribute('aria-label') || el.getAttribute('aria-label') || '';
                    title = norm(ariaLabel.replace(/\\(.*?\\)/g, ''));
                }

                let slicerType = 'lista';
                if (allClasses.includes('chiclet')) slicerType = 'chiclet';
                else if (vc.querySelector('input[type="range"], .slider, [class*="range"], [class*="Range"]')) slicerType = 'range';
                else if (vc.querySelector('select, .dropdown, [class*="dropdown"], [class*="Dropdown"]')) slicerType = 'dropdown';
                else if (vc.querySelector('.date-slicer, [class*="date-slicer"], [class*="DateSlicer"]')) slicerType = 'date';
                else if (vc.querySelector('input[type="text"], input.searchInput')) slicerType = 'busca';

                const allValues = [];
                const selectedValues = [];
                const seen = new Set();
                const selectedSet = new Set();

                const items = vc.querySelectorAll(
                    '.slicerItemContainer, [class*="slicerItem"], [role="option"], [role="listitem"], .row, [class*="chiclet"]'
                );
                items.forEach(item => {
                    const value = norm(
                        item.querySelector('.slicerText, span, [class*="slicerText"], [class*="text"], label')?.textContent || item.textContent
                    );
                    if (!value || value.length > 80) return;
                    const lower = value.toLowerCase();
                    if (lower === 'selecionar tudo' || lower === 'select all') return;
                    if (!seen.has(value)) {
                        seen.add(value);
                        allValues.push(value);
                    }

                    const checkbox = item.querySelector('.slicerCheckbox, input[type="checkbox"], [class*="checkbox"], [class*="Checkbox"]');
                    const isSelected = (
                        item.classList.contains('selected') ||
                        item.classList.contains('isSelected') ||
                        item.querySelector('.selected, .isSelected, .partiallySelected') !== null ||
                        item.getAttribute('aria-selected') === 'true' ||
                        item.getAttribute('aria-checked') === 'true' ||
                        (checkbox && (
                            checkbox.checked === true ||
                            checkbox.getAttribute('aria-checked') === 'true' ||
                            checkbox.classList.contains('selected') ||
                            checkbox.classList.contains('partiallySelected')
                        ))
                    );
                    if (isSelected && !selectedSet.has(value)) {
                        selectedSet.add(value);
                        selectedValues.push(value);
                    }
                });

                const visibleStateCandidates = [
                    vc.querySelector('input[type="text"]')?.value,
                    vc.querySelector('input[aria-autocomplete]')?.value,
                    vc.querySelector('.slicer-dropdown-menu, .dropdown-value, [class*="selectedValue"], [class*="currentValue"]')?.textContent,
                    vc.querySelector('[aria-selected="true"]')?.textContent,
                    vc.querySelector('[aria-checked="true"]')?.textContent
                ].map(norm).filter(Boolean);

                for (const val of visibleStateCandidates) {
                    if (!selectedSet.has(val) && val.length <= 80) {
                        selectedSet.add(val);
                        selectedValues.push(val);
                    }
                }

                const visibleText = norm(vc.textContent).toLowerCase();
                const pendingSignals = ['ainda não aplicado', 'not yet applied', 'apply changes', 'aplicar alterações'];
                const selectedSignals = ['múltiplos selecionados', 'multiple selections', 'selecionado', 'selected'];
                const hasPending = pendingSignals.some(t => visibleText.includes(t));
                const inferredFiltered = (
                    selectedValues.length > 0 ||
                    selectedSignals.some(t => visibleText.includes(t)) ||
                    vc.querySelector('[aria-selected="true"], [aria-checked="true"], .selected, .isSelected, .partiallySelected') !== null
                );

                results.push({
                    index: index,
                    title: title || `Slicer #${index + 1}`,
                    type: slicerType,
                    values: selectedValues.length ? selectedValues.slice(0, 10) : allValues.slice(0, 10),
                    allValues: allValues.slice(0, 50),
                    selectedValues: selectedValues.slice(0, 50),
                    totalValues: allValues.length,
                    totalSelected: selectedValues.length,
                    hasPending: hasPending,
                    applied: inferredFiltered && !hasPending,
                    x: Math.round(rect.x),
                    y: Math.round(rect.y),
                    width: Math.round(rect.width),
                    height: Math.round(rect.height),
                });
            });

            return { slicers: results, diagnostics };
        })()
    """, default={"slicers": [], "diagnostics": {}})

    slicers = list((payload or {}).get("slicers") or [])
    diagnostics = (payload or {}).get("diagnostics") or {}

    log.info(f"🧪 [diagnóstico] visual-container bruto no DOM: {int(diagnostics.get('rawVisualContainers', 0))}")
    log.info(f"🧪 [diagnóstico] candidatos brutos a slicer: {int(diagnostics.get('rawSlicerCandidates', 0))}")
    log.info(f"🧪 [diagnóstico] descartados por tamanho (slicer scan): {int(diagnostics.get('discardedBySize', 0))}")

    slicers.sort(key=lambda s: (s.get("y", 0), s.get("x", 0)))

    await cleanup_residual_ui(tab, stage_label="scan de slicers - final", aggressive=True)
    return slicers


def display_slicers(slicers: list):
    """Exibe slicers/filtros encontrados no terminal."""
    print("\n" + "=" * 100)
    print("🎚️  FILTROS / SLICERS ENCONTRADOS")
    print("=" * 100)

    if not slicers:
        print("  Nenhum slicer/filtro identificado.")
        print("=" * 100)
        return

    for i, s in enumerate(slicers):
        title = s.get("title", f"Slicer #{i+1}")
        stype = s.get("type", "?")
        applied = s.get("applied", False)
        idx = s.get("index", i)

        status = "✅APLICADO" if applied else "⚠️PENDENTE"
        title_display = title if applied else f"{title}(Ainda não aplicado)"

        print(f"\n  [{i}] {title_display} (tipo: {stype}, container #{idx}) {status}")
        print(f"      Tamanho: {s.get('width', '?')}x{s.get('height', '?')}")

        values = s.get("values") or []
        if values:
            print(f"      Valores visíveis: {', '.join(values[:6])}")
        else:
            print("      Valores: não detectados (dropdown fechado ou formato especial)")

    print("\n" + "-" * 100)
    print(f"  Total: {len(slicers)} filtros")
    print("=" * 100)

# ---------------------------------------------------------------------------
# Carregamento robusto — combate redirecionamento por homepage
# ---------------------------------------------------------------------------

async def get_current_url(tab) -> str:
    """Obtém a URL atual da aba."""
    try:
        result = await tab.evaluate("window.location.href")
        return str(result) if result else ""
    except Exception:
        return ""


async def is_powerbi_loaded(tab) -> bool:
    """Verifica se o Power BI realmente renderizou (visual-containers > 0)."""
    try:
        count = await tab.evaluate(
            "document.querySelectorAll('visual-container').length"
        )
        return count is not None and int(str(count)) > 0
    except Exception:
        return False


async def is_on_correct_url(tab, target_url: str) -> bool:
    """Verifica se a aba está na URL correta (ou num redirect legítimo do PBI)."""
    current = await get_current_url(tab)
    if not current:
        return False

    # Extrai o reportId da URL alvo para comparar
    # Funciona mesmo se o PBI redirecionar para uma URL ligeiramente diferente
    target_lower = target_url.lower()
    current_lower = current.lower()

    # Caso 1: URL exata (ou começo dela)
    if current_lower.startswith(target_lower[:60]):
        return True

    # Caso 2: Mesmo reportId (PBI pode mudar parâmetros)
    if "reportid=" in target_lower:
        report_id = target_lower.split("reportid=")[1].split("&")[0]
        if report_id in current_lower:
            return True

    # Caso 3: Está em app.powerbi.com (redirecionamento legítimo)
    if "powerbi.com" in current_lower and "reportid" in current_lower:
        return True

    return False


async def find_correct_tab(browser, target_url: str):
    """
    Procura em todas as abas abertas a que contém o Power BI.
    A homepage pode abrir em outra aba enquanto o PBI fica em background.
    """
    try:
        targets = await browser.get_targets()
        if not targets:
            return None
        for target in targets:
            t_url = str(getattr(target, 'url', '') or '').lower()
            if 'powerbi.com' in t_url or 'reportembed' in t_url:
                # Tenta ativar essa aba
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
    """
    Garante que estamos na URL correta do Power BI e que ele carregou.
    Combate:
    - Homepage roubando foco ao abrir o navegador
    - Redirecionamentos lentos do Power BI
    - Página em branco / about:blank
    
    Retorna a tab correta (pode ser diferente da original se trocou de aba).
    """
    MAX_RETRIES = 5
    RETRY_WAIT = 8

    for attempt in range(1, MAX_RETRIES + 1):
        log.info(f"🔄 Verificação de página (tentativa {attempt}/{MAX_RETRIES})...")

        current_url = await get_current_url(tab)
        log.info(f"  📍 URL atual: {current_url[:100]}")

        # Verifica se está na URL certa
        on_correct = await is_on_correct_url(tab, target_url)

        if on_correct:
            log.info("  ✅ URL correta!")

            # Agora espera o Power BI renderizar
            log.info("  ⏳ Aguardando Power BI renderizar...")
            for load_attempt in range(8):
                if await is_powerbi_loaded(tab):
                    log.info(f"  ✅ Power BI carregado! (visual-containers detectados)")
                    return tab
                log.info(f"    Tentativa {load_attempt + 1}/8: aguardando visual-containers...")
                await asyncio.sleep(MEDIUM_WAIT)

            # Se chegou aqui, URL certa mas PBI não carregou — recarrega
            log.warning("  ⚠️ URL correta mas Power BI não renderizou, recarregando...")
            await tab.evaluate("window.location.reload()")
            await asyncio.sleep(PAGE_LOAD_WAIT)
            continue

        # URL errada — provavelmente a homepage roubou o foco
        log.warning(f"  ⚠️ URL incorreta! Esperava Power BI, encontrou: {current_url[:80]}")

        # Estratégia 1: Procura se o PBI está em outra aba
        log.info("  🔍 Procurando aba do Power BI em outras tabs...")
        pbi_tab = await find_correct_tab(browser, target_url)
        if pbi_tab:
            log.info("  ✅ Encontrou aba do Power BI! Trocando...")
            tab = pbi_tab
            await asyncio.sleep(MEDIUM_WAIT)
            continue

        # Estratégia 2: Navega forçadamente para a URL correta
        log.info(f"  🔄 Forçando navegação para URL do Power BI...")
        await tab.evaluate(f'window.location.href = "{target_url}"')
        await asyncio.sleep(PAGE_LOAD_WAIT)

        # Verifica de novo
        if await is_on_correct_url(tab, target_url):
            log.info("  ✅ Navegação forçada funcionou!")
            # Espera carregar
            for load_attempt in range(6):
                if await is_powerbi_loaded(tab):
                    log.info("  ✅ Power BI carregado!")
                    return tab
                await asyncio.sleep(MEDIUM_WAIT)
        else:
            # Estratégia 3: Abre nova aba com a URL
            log.warning("  ⚠️ Navegação falhou, abrindo nova aba...")
            try:
                tab = await browser.get(target_url, new_tab=True)
                if owned_tab_refs is not None:
                    owned_tab_refs.add(_tab_ref(tab))
                await asyncio.sleep(PAGE_LOAD_WAIT)
            except Exception as e:
                log.error(f"  ❌ Erro ao abrir nova aba: {e}")

    # Última chance: tenta seguir mesmo sem confirmação total
    log.warning("⚠️ Não foi possível confirmar carregamento completo. Tentando prosseguir...")
    return tab


# ---------------------------------------------------------------------------
# Fluxo principal
# ---------------------------------------------------------------------------

async def navigate_to_report_page(tab, target_page: str) -> bool:
    """
    Tenta navegar para uma página/aba interna do relatório Power BI pelo nome visível.
    Ex.: COMPARATIVO

    Retorna True se clicou/encontrou, False se não encontrou.
    """
    if not target_page:
        return True

    target_norm = target_page.strip().lower()
    log.info(f"📄 Tentando navegar para página interna do relatório: '{target_page}'")

    async def _get_page_signature():
        try:
            raw = await tab.evaluate("""
                (() => {
                    const textOf = (el) => (el?.innerText || el?.textContent || '').replace(/\\s+/g, ' ').trim();
                    const norm = (s) => (s || '').toLowerCase();

                    let selectedTab = '';
                    const selectedCandidates = Array.from(document.querySelectorAll(
                        '[role="tab"][aria-selected=\"true\"], li.section.active, li[class*=\"active\"]'
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
                "visual_count": int(parsed.get("visualCount") or 0),
            }
        except Exception:
            return {"selected_tab": "", "visual_count": 0}

    async def _legacy_click_page():
        """
        Restaura o miolo funcional de navegação usado na v05 (mais próximo disponível).
        """
        try:
            result = await tab.evaluate(f"""
                (() => {{
                    const name = {json.dumps(target_page)}.toUpperCase();

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

                    const ariaEl = document.querySelector(`[aria-label*="${{name}}"]`)
                                || document.querySelector(`[title*="${{name}}"]`);
                    if (ariaEl) {{ ariaEl.click(); return "aria-label/title"; }}

                    return null;
                }})()
            """)
        except Exception:
            result = None
        clicked_detail = str(result or "").strip()
        return bool(clicked_detail), clicked_detail

    before = await _get_page_signature()
    log.info(
        f"  ℹ️ Estado antes da navegação: selectedTab='{before.get('selected_tab','')}', "
        f"visuals={before.get('visual_count', 0)}"
    )

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
            log.info(
                f"  ✅ Navegação confirmada para '{target_page}': "
                f"selectedTab='{after.get('selected_tab','')}', visuals={after.get('visual_count', 0)}"
            )
            return True

        log.warning(
            f"  ⚠️ Clique ocorreu, mas sem confirmação sólida da troca de página. "
            f"selectedTab='{after.get('selected_tab','')}', visuals={after.get('visual_count', 0)}"
        )
        await asyncio.sleep(1.5)

    log.error(f"❌ Não foi possível navegar para a página '{target_page}'.")
    return False

async def wait_for_visual_containers(tab, retries: int = 10, wait_seconds: int = 3) -> bool:
    """
    Aguarda os visual-containers aparecerem na página atual do relatório.
    """
    for attempt in range(1, retries + 1):
        try:
            count = await tab.evaluate("document.querySelectorAll('visual-container').length")
            count = int(count or 0)
        except Exception:
            count = 0

        if count > 0:
            log.info(f"✅ Visual-containers detectados: {count}")
            return True

        log.info(f"⏳ Aguardando visuais renderizarem... tentativa {attempt}/{retries}")
        await asyncio.sleep(wait_seconds)

    return False


async def wait_for_visuals_or_abort(
    tab,
    stage_label: str,
    retries: int = 10,
    wait_seconds: int = 3,
    allow_reload: bool = True,
) -> bool:
    """
    Aguarda renderização real dos visuais antes de seguir o fluxo.

    Se não encontrar visual-container:
    - registra aviso
    - tenta um reload controlado (quando permitido)
    - aborta com erro claro se continuar zerado
    """
    log.info(f"⏳ Aguardando visual-containers ({stage_label})...")
    loaded = await wait_for_visual_containers(tab, retries=retries, wait_seconds=wait_seconds)
    if loaded:
        return True

    log.warning(
        f"⚠️ Nenhum visual-container detectado em '{stage_label}' após {retries} tentativas."
    )

    if allow_reload:
        log.warning(f"🔄 Tentando reload controlado em '{stage_label}'...")
        try:
            await tab.evaluate("window.location.reload()")
            await asyncio.sleep(PAGE_LOAD_WAIT)
        except Exception as e:
            log.warning(f"⚠️ Falha ao recarregar página em '{stage_label}': {e}")

        loaded_after_reload = await wait_for_visual_containers(
            tab, retries=max(6, retries // 2), wait_seconds=wait_seconds
        )
        if loaded_after_reload:
            return True

    log.error(
        f"❌ Abortando: página sem visual-containers em '{stage_label}' mesmo após nova tentativa."
    )
    return False




async def run_export(url: str, browser_path: str, target_page: str):
    """Executa o fluxo completo."""
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
        # Abre somente a aba do relatório que será controlada pelo script
        tab = await open_report_tab(browser, url, owned_tab_refs)
        await safe_focus_tab(tab)

        # Mantém apenas a aba criada pelo script dentro da instância controlada
        await close_extra_tabs_created_by_script(browser, tab, owned_tab_refs)

        # Permite múltiplos downloads
        await allow_multiple_downloads(tab)

        log.info(f"⏳ Aguardando carregamento inicial ({PAGE_LOAD_WAIT}s)...")
        await asyncio.sleep(PAGE_LOAD_WAIT)

        # Proteções gerais
        await block_microsoft_learn_and_external_links(tab)
        await cleanup_residual_ui(tab, stage_label="preparação geral após carregamento", aggressive=True)
        tab = await ensure_report_tab_still_valid(tab, url)

        # ── Carregamento robusto: combate redirecionamento por homepage ──
        try:
            current_url = await get_tab_url(tab)
            if current_url and "powerbi.com" not in current_url.lower():
                log.warning("⚠️ URL atual não parece ser do relatório. Reabrindo...")
                await tab.get(url)
                await asyncio.sleep(LONG_WAIT)
        except Exception:
            pass

        # Só segue após renderização real da página inicial do relatório
        initial_loaded = await wait_for_visuals_or_abort(
            tab,
            stage_label="carregamento inicial do relatório",
            retries=10,
            wait_seconds=3,
            allow_reload=True,
        )
        if not initial_loaded:
            return False

        # Navegação opcional para página específica do relatório
        if target_page:
            page_ok = await navigate_to_report_page(tab, target_page)
            if not page_ok:
                log.error(
                    f"❌ Fluxo interrompido: a navegação para '{target_page}' não foi confirmada."
                )
                return False

            target_loaded = await wait_for_visuals_or_abort(
                tab,
                stage_label=f"navegação para TARGET_PAGE='{target_page}'",
                retries=8,
                wait_seconds=3,
                allow_reload=True,
            )
            if not target_loaded:
                return False
        else:
            # Mesmo sem TARGET_PAGE, garante renderização antes de escanear.
            ready_without_target = await wait_for_visuals_or_abort(
                tab,
                stage_label="página inicial (sem TARGET_PAGE)",
                retries=8,
                wait_seconds=3,
                allow_reload=True,
            )
            if not ready_without_target:
                return False

        await close_open_menus_and_overlays(tab, aggressive=True)

        try:
            current_visual_count = await tab.evaluate("document.querySelectorAll('visual-container').length")
            current_visual_count = int(current_visual_count or 0)
        except Exception:
            current_visual_count = 0

        if current_visual_count == 0:
            log.error("❌ Nenhum visual-container encontrado antes do scan. Abortando para evitar diagnóstico falso.")
            return False

        # Escaneia filtros/slicers
        print("\n" + "=" * 70)
        print("📌 ETAPA 1 - LEITURA DE FILTROS")
        print("=" * 70)

        await cleanup_residual_ui(tab, stage_label="antes do scan de slicers", aggressive=True)
        slicers = await scan_slicers(tab)
        display_slicers(slicers)
        await cleanup_residual_ui(tab, stage_label="após scan de slicers", aggressive=True)

        # Escaneia visuais
        log.info("======================================================================")
        log.info("📌 Escaneando visuais disponíveis")
        log.info("======================================================================")
        await cleanup_residual_ui(tab, stage_label="antes do scan de visuais", aggressive=True)
        visuals = await scan_visuals(tab)
        display_visuals(visuals)

        exportable = [i for i, v in enumerate(visuals) if v.get("hasExportData")]
        if not exportable:
            log.warning("⚠️ Nenhum visual com 'Exportar dados' confirmado.")
            return False

        # Pergunta ao usuário
        selected_indexes = await ask_user_visual_selection(visuals)
        if not selected_indexes:
            log.warning("⛔ Exportação cancelada pelo usuário.")
            return False

        log.info(f"📥 {len(selected_indexes)} visuais selecionados para exportação")

        # Exportação
        results = await export_selected_visuals(tab, visuals, selected_indexes)
        display_export_summary(results)

        return True

    finally:
        await graceful_browser_shutdown(tab, browser)


async def graceful_browser_shutdown(tab, browser):
    """
    Encerra recursos do nodriver de forma suave para reduzir erros de pipe/transport,
    comuns no fim da execução (ex.: Windows + Python 3.14).

    Regras:
    - limpa overlays/popups antes de fechar;
    - tenta fechar a aba controlada de forma não agressiva;
    - finaliza a instância do browser apenas uma vez, com suppress de erros esperados.
    """
    # 1) limpeza de UI antes do fechamento
    if tab is not None:
        with contextlib.suppress(Exception):
            await dismiss_sensitive_data_popup(tab)
        with contextlib.suppress(Exception):
            await close_open_menus_and_overlays(tab, aggressive=True)
        with contextlib.suppress(Exception):
            await asyncio.sleep(0.3)

        # 2) tenta fechar somente a aba controlada (suave)
        with contextlib.suppress(BrokenPipeError, ConnectionResetError, RuntimeError, OSError, Exception):
            await tab.close()
        with contextlib.suppress(Exception):
            await asyncio.sleep(0.2)

    # 3) encerra browser da automação (sem matar processos externos)
    if browser is not None:
        with contextlib.suppress(BrokenPipeError, ConnectionResetError, RuntimeError, OSError, Exception):
            browser.stop()
        with contextlib.suppress(Exception):
            await asyncio.sleep(0.2)

# ---------------------------------------------------------------------------
# Entrada
# ---------------------------------------------------------------------------

def main():
    if not validate_runtime_config():
        sys.exit(1)

    try:
        success = asyncio.run(run_export(POWERBI_URL, BROWSER_PATH, TARGET_PAGE))
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        log.warning("⛔ Execução interrompida pelo usuário")
        sys.exit(130)


if __name__ == "__main__":
    main()
