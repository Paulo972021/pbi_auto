#pbi_auto_v05.py
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
import subprocess
import time
import json
import re
import configpbi
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
# Constantes de timeout (segundos)
# ---------------------------------------------------------------------------
PAGE_LOAD_WAIT = 20
SHORT_WAIT = 3
MEDIUM_WAIT = 5
LONG_WAIT = 10
DOWNLOAD_WAIT = 15


# ---------------------------------------------------------------------------
# Funções auxiliares
# ---------------------------------------------------------------------------

def kill_browser_processes(browser_path: str):
    """Mata processos residuais do navegador para evitar erro WebSocket 500."""
    exe_name = os.path.basename(browser_path).lower()
    try:
        if sys.platform == "win32":
            subprocess.run(["taskkill", "/F", "/IM", exe_name],
                           capture_output=True, timeout=5)
        else:
            subprocess.run(["pkill", "-f", exe_name.replace(".exe", "")],
                           capture_output=True, timeout=5)
        time.sleep(2)
        log.info(f"🧹 Processos residuais de '{exe_name}' encerrados")
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


async def scroll_visual_into_view(tab, idx: int) -> bool:
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
    await scroll_visual_into_view(tab, idx)
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


async def scan_visuals(tab) -> list:
    """
    Escaneia os visual-containers e confirma com retry quais suportam
    'Exportar dados'. Faz scroll antes do probe para evitar falhas com
    visuais mais abaixo da página.
    """
    log.info("🔎 Escaneando visuais na página...")

    await hover_all_visuals(tab)

    visuals = await eval_json(tab, """
        (() => {
            const results = [];
            const containers = document.querySelectorAll('visual-container');

            containers.forEach((vc, index) => {
                const el = vc.querySelector('transform') || vc;
                const rect = el.getBoundingClientRect();
                if (rect.width < 20 || rect.height < 20) return;

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
                    const textContent = el.textContent?.replace(/\s+/g, ' ').trim()?.substring(0, 100) || '';
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
                    hasOptionsBtn: !!optionsBtn,
                    hasHeader: hasHeader,
                    ariaLabel: ariaLabel.substring(0, 60)
                });
            });

            return JSON.stringify(results);
        })()
    """, default=[])

    if not visuals:
        log.error("❌ Nenhum visual encontrado no DOM")
        return []

    exportable = [v for v in visuals if v.get('hasHeader') or v.get('hasOptionsBtn')]
    exportable.sort(key=lambda v: (0 if v.get('hasOptionsBtn') else 1, v.get('y', 0), v.get('x', 0)))

    log.info(f"📊 Total de visual-containers: {len(visuals)}")
    log.info(f"📋 Visuais com header de opções: {len(exportable)}")
    log.info("🔬 Verificando quais visuais suportam 'Exportar dados'...")

    for v in exportable:
        idx = v['index']
        v['hasExportData'] = await probe_visual_export(tab, idx)

    export_count = sum(1 for v in exportable if v.get('hasExportData'))
    log.info(f"✅ {export_count} visuais com 'Exportar dados' confirmado")
    return exportable


def display_visuals_menu(visuals: list):
    """Exibe a lista de visuais no terminal de forma organizada."""
    print("\n" + "=" * 90)
    print("📋 VISUAIS ENCONTRADOS NA PÁGINA")
    print("=" * 90)
    print(f"{'#':<4} {'Tipo':<12} {'Título':<40} {'Tam.':<12} {'Export?':<8} {'Btn?':<5}")
    print("-" * 90)

    for i, v in enumerate(visuals):
        btn_status = "✅" if v.get('hasOptionsBtn') else "🔍"
        export_status = "✅" if v.get('hasExportData') else "❌"
        size = f"{v['width']}x{v['height']}"
        title = v['title'][:38]
        vtype = v.get('type', '?')[:10]
        print(f"{i:<4} {vtype:<12} {title:<40} {size:<12} {export_status:<8} {btn_status:<5}")

    print("-" * 90)
    print("Export? = suporta 'Exportar dados' | Btn? = botão 'Mais opções' detectado")
    print("⚠️  Visuais com Export? = ❌ não possuem opção de exportação de dados")
    print()


def ask_user_selection(visuals: list) -> list:
    """Pergunta ao usuário quais visuais exportar."""
    print("📥 Quais visuais deseja exportar?")
    print("   Opções:")
    print("   • Digite os números separados por vírgula: 0,2,5")
    print("   • Digite 'todos' para exportar todos")
    print("   • Digite 'sair' para cancelar")
    print()

    while True:
        choice = input("   Sua escolha: ").strip().lower()

        if choice == 'sair' or choice == 'exit' or choice == 'q':
            return []

        if choice == 'todos' or choice == 'all' or choice == '*':
            return list(range(len(visuals)))

        try:
            indices = [int(x.strip()) for x in choice.split(',')]
            # Valida
            invalid = [i for i in indices if i < 0 or i >= len(visuals)]
            if invalid:
                print(f"   ❌ Índices inválidos: {invalid}. Tente novamente.")
                continue
            return indices
        except ValueError:
            print("   ❌ Formato inválido. Use números separados por vírgula (ex: 0,2,5)")


# ---------------------------------------------------------------------------
# Exportação de um visual específico
# ---------------------------------------------------------------------------

async def export_single_visual(tab, visual: dict, visual_number: int, total: int) -> bool:
    """
    Exporta os dados de um visual específico com tratamento mais robusto para:
    - visuais fora da dobra
    - menus que só aparecem após scroll/hover real
    - aviso de confidencialidade com fluxo de Copiar/Fechar
    """
    idx = visual['index']
    title = visual['title'][:40]
    log.info(f"📦 [{visual_number}/{total}] Exportando: {title} (container #{idx})")

    await scroll_visual_into_view(tab, idx)
    if not await hover_visual(tab, idx):
        log.warning(f"  ⚠️ Não foi possível fazer hover no visual #{idx}")
        return False

    if not await open_visual_more_options(tab, idx, attempts=5):
        log.warning(f"  ⚠️ Botão 'Mais opções' não encontrado para visual #{idx}")
        return False

    menu_debug = await eval_json(tab, """
        (() => {
            const results = [];
            const menuSelectors = [
                '[role="menu"]', '[role="menubar"]', '[role="listbox"]',
                '.context-menu', '.dropdown-menu', '.cdk-overlay-pane',
                'mat-menu-panel', '[class*="menu"]', '[class*="dropdown"]',
                '[class*="popup"]', '[class*="overlay"]'
            ];
            for (const sel of menuSelectors) {
                document.querySelectorAll(sel).forEach(menu => {
                    const rect = menu.getBoundingClientRect();
                    if (rect.width > 10 && rect.height > 10) {
                        menu.querySelectorAll('button, [role="menuitem"], li, a, div[tabindex], span[tabindex]').forEach(item => {
                            const text = item.textContent?.trim();
                            if (text && text.length < 80) {
                                results.push({
                                    text: text.substring(0, 60),
                                    tag: item.tagName,
                                    role: item.getAttribute('role') || ''
                                });
                            }
                        });
                    }
                });
            }
            return JSON.stringify(results);
        })()
    """, default=[])
    if menu_debug:
        log.info(f"  📋 Menu aberto com {len(menu_debug)} itens")
        for item in menu_debug[:12]:
            log.info(f"    tag={item.get('tag','')} role='{item.get('role','')}' text='{item.get('text','')}'")

    clicked_label = await click_export_data_menu(tab)
    if not clicked_label:
        log.warning(f"  ⚠️ 'Exportar dados' não encontrada para visual #{idx}")
        await close_menu(tab)
        return False

    log.info(f"  ✅ Menu acionado: {clicked_label}")
    await asyncio.sleep(MEDIUM_WAIT)

    await dismiss_sensitive_data_warning(tab)
    await asyncio.sleep(1.5)

    log.info("  🔘 Selecionando tipo de exportação...")
    radio_selected = False
    for js_expr in [
        'document.querySelector("#pbi-radio-button-1 > label > section > div")',
        'document.querySelector("#pbi-radio-button-1 label")',
        'document.querySelector("#pbi-radio-button-1")',
        'document.querySelector("input[type=radio]:not([disabled])")'
    ]:
        if await js_click(tab, js_expr, "Radio button"):
            radio_selected = True
            break
    if not radio_selected:
        log.info("  ℹ️ Nenhum radio específico detectado; seguindo com o diálogo atual")

    await asyncio.sleep(SHORT_WAIT)

    log.info("  📤 Confirmando exportação...")
    export_confirmed = False
    for js_expr in [
        'document.querySelector("export-data-dialog button.exportButton")',
        'document.querySelector("button.exportButton")',
        'document.querySelector("button.primaryBtn.exportButton")',
        'document.querySelector("mat-dialog-actions button.primaryBtn")',
        'document.querySelector("mat-dialog-actions button:first-child")'
    ]:
        if await js_click(tab, js_expr, "Botão Exportar"):
            export_confirmed = True
            break

    if not export_confirmed:
        if await js_click_xpath(
            tab,
            '//*[@id="mat-mdc-dialog-0"]/div/div/export-data-dialog/mat-dialog-actions/button[1]',
            "Exportar (XPath)"
        ):
            export_confirmed = True

    if not export_confirmed:
        result = await tab.evaluate("""
            (() => {
                for (const btn of document.querySelectorAll('button, [role="button"]')) {
                    const text = btn.textContent?.trim()?.toLowerCase() || '';
                    if (text === 'exportar' || text === 'export' || text.includes('exportar')) {
                        const r = btn.getBoundingClientRect();
                        if (r.width > 0 && r.height > 0) {
                            btn.click();
                            return true;
                        }
                    }
                }
                return false;
            })()
        """)
        if result:
            export_confirmed = True
            log.info("  ✅ Exportar (texto)")

    if not export_confirmed:
        log.warning(f"  ⚠️ Não foi possível confirmar exportação do visual #{idx}")
        await close_any_dialog(tab)
        return False

    log.info(f"  ⏳ Aguardando download ({DOWNLOAD_WAIT}s)...")
    await asyncio.sleep(3)
    await accept_download_permission(tab)
    await asyncio.sleep(max(1, DOWNLOAD_WAIT - 3))
    await dismiss_sensitive_data_warning(tab)
    await close_any_dialog(tab)

    log.info(f"  🎉 Visual #{idx} exportado com sucesso!")
    return True


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

async def scan_slicers(tab) -> list:
    """
    Escaneia slicers com leitura passiva mais forte do estado visível.
    Melhora a leitura de filtros aplicados e reduz falsos "Ainda não aplicado".
    """
    log.info("🎚️  Escaneando filtros/slicers na página...")

    slicers = await eval_json(tab, """
        (() => {
            const results = [];
            const containers = document.querySelectorAll('visual-container');
            const norm = txt => (txt || '').replace(/\s+/g, ' ').trim();

            containers.forEach((vc, index) => {
                const el = vc.querySelector('transform') || vc;
                const rect = el.getBoundingClientRect();
                if (rect.width < 15 || rect.height < 15) return;

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

                let name = '';
                const headerEl = vc.querySelector(
                    '.slicer-header-text, [class*="header-text"], h3, h4, .visual-title, .visualTitle, [class*="title"]'
                );
                if (headerEl) name = norm(headerEl.textContent);
                if (!name) {
                    const ariaLabel = vc.getAttribute('aria-label') || el.getAttribute('aria-label') || '';
                    name = norm(ariaLabel.replace(/\(.*?\)/g, ''));
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
                    name: name || `Slicer #${index + 1}`,
                    type: slicerType,
                    allValues: allValues.slice(0, 50),
                    selectedValues: selectedValues.slice(0, 50),
                    totalValues: allValues.length,
                    totalSelected: selectedValues.length,
                    hasPending: hasPending,
                    isFiltered: inferredFiltered && !hasPending,
                    width: Math.round(rect.width),
                    height: Math.round(rect.height)
                });
            });

            return JSON.stringify(results);
        })()
    """, default=[])

    log.info(f"🎚️  {len(slicers)} slicers encontrados")
    return slicers


def display_slicers(slicers: list):
    """Exibe os filtros/slicers encontrados com seus valores."""
    if not slicers:
        print("\n  Nenhum filtro/slicer encontrado na página.\n")
        return

    print("\n" + "=" * 100)
    print("🎚️  FILTROS / SLICERS ENCONTRADOS")
    print("=" * 100)

    for i, s in enumerate(slicers):
        pending = " ⚠️PENDENTE" if s.get('hasPending') else ""
        filtered = " ✅FILTRADO" if s.get('isFiltered') else ""
        print(f"\n  [{i}] {s['name']} (tipo: {s['type']}, container #{s['index']}){pending}{filtered}")
        print(f"      Tamanho: {s['width']}x{s['height']}")

        all_vals = s.get('allValues', [])
        sel_vals = s.get('selectedValues', [])
        total = s.get('totalValues', 0)
        total_sel = s.get('totalSelected', 0)

        if all_vals:
            print(f"      Valores disponíveis ({total}):")
            for v in all_vals[:15]:
                marker = " ✅" if v in sel_vals else "   "
                print(f"        {marker} {v}")
            if total > 15:
                print(f"        ... e mais {total - 15} valores")

            if sel_vals:
                print(f"      Selecionados ({total_sel}): {', '.join(sel_vals[:10])}")
                if total_sel > 10:
                    print(f"        ... e mais {total_sel - 10}")
            else:
                print(f"      Selecionados: Todos / Nenhuma seleção específica")
        else:
            if sel_vals:
                print(f"      Valor atual: {', '.join(sel_vals)}")
            else:
                print(f"      Valores: não detectados (dropdown fechado ou formato especial)")

    print("\n" + "-" * 100)
    print(f"  Total: {len(slicers)} filtros")
    print("=" * 100)
    print()


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


async def ensure_correct_page(browser, tab, target_url: str):
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
                await asyncio.sleep(PAGE_LOAD_WAIT)
            except Exception as e:
                log.error(f"  ❌ Erro ao abrir nova aba: {e}")

    # Última chance: tenta seguir mesmo sem confirmação total
    log.warning("⚠️ Não foi possível confirmar carregamento completo. Tentando prosseguir...")
    return tab


# ---------------------------------------------------------------------------
# Fluxo principal
# ---------------------------------------------------------------------------

async def run_export(url: str, browser_path: str, target_page: str):
    """Executa o fluxo completo."""
    log.info("🚀 Iniciando automação Power BI")
    log.info(f"📎 URL: {url}")
    log.info(f"🌐 Navegador: {browser_path}")
    if target_page:
        log.info(f"📄 Página alvo: {target_page}")

    kill_browser_processes(browser_path)

    browser = await uc.start(
        headless=False,
        browser_executable_path=browser_path,
        browser_args=[
            "--no-sandbox",
            "--disable-blink-features=AutomationControlled",
            "--disable-infobars",
            "--start-maximized",
            "--lang=pt-BR",
            "--disable-popup-blocking",
        ],
    )

    try:
        # ── Carregamento robusto: combate redirecionamento por homepage ──
        tab = await browser.get(url)

        # Permite múltiplos downloads automáticos via CDP
        await allow_multiple_downloads(tab)

        log.info(f"⏳ Aguardando carregamento inicial ({PAGE_LOAD_WAIT}s)...")
        await asyncio.sleep(PAGE_LOAD_WAIT)

        tab = await ensure_correct_page(browser, tab, url)

        # ── Navegar para página alvo ──
        if target_page:
            log.info("=" * 70)
            log.info(f"📌 Navegando para: {target_page}")
            log.info("=" * 70)
            if not await navigate_to_page(tab, target_page):
                log.error("💥 Abortando: página não encontrada")
                return False
            log.info(f"⏳ Aguardando página carregar ({LONG_WAIT}s)...")
            await asyncio.sleep(LONG_WAIT)

        # ── Scan de filtros/slicers ──
        log.info("=" * 70)
        log.info("📌 Escaneando filtros/slicers")
        log.info("=" * 70)

        slicers = await scan_slicers(tab)
        display_slicers(slicers)

        # ── Scan de visuais ──
        log.info("=" * 70)
        log.info("📌 Escaneando visuais disponíveis")
        log.info("=" * 70)

        visuals = await scan_visuals(tab)

        if not visuals:
            log.error("❌ Nenhum visual exportável encontrado na página")
            return False

        # ── Exibir lista e pedir seleção ──
        display_visuals_menu(visuals)
        selected_indices = ask_user_selection(visuals)

        if not selected_indices:
            log.info("👋 Nenhum visual selecionado. Encerrando.")
            return True

        selected_visuals = [visuals[i] for i in selected_indices]
        log.info(f"📥 {len(selected_visuals)} visuais selecionados para exportação")

        # ── Exportar cada visual selecionado ──
        log.info("=" * 70)
        log.info("📌 Iniciando exportação")
        log.info("=" * 70)

        results = []
        for num, visual in enumerate(selected_visuals, 1):
            success = await export_single_visual(tab, visual, num, len(selected_visuals))
            results.append({
                'title': visual['title'],
                'index': visual['index'],
                'success': success,
            })
            # Pausa entre exportações para não sobrecarregar
            if num < len(selected_visuals):
                await asyncio.sleep(SHORT_WAIT)

        # ── Resumo final ──
        print("\n" + "=" * 70)
        print("📊 RESUMO DA EXPORTAÇÃO")
        print("=" * 70)
        ok = sum(1 for r in results if r['success'])
        fail = sum(1 for r in results if not r['success'])
        for r in results:
            status = "✅" if r['success'] else "❌"
            print(f"  {status} {r['title'][:60]}")
        print(f"\n  Total: {ok} sucesso, {fail} falha")
        print(f"  📂 Arquivos salvos na pasta Downloads padrão")
        print("=" * 70)

        await asyncio.sleep(5)
        return fail == 0

    except Exception as e:
        log.error(f"💥 Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        try:
            browser.stop()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Entrada
# ---------------------------------------------------------------------------

def main():
    # ╔══════════════════════════════════════════════════════════════╗
    # ║  COLE O LINK DO POWER BI AQUI EMBAIXO (entre as aspas):    ║
    # ╚══════════════════════════════════════════════════════════════╝
    POWERBI_URL = configpbi.url
    

    # ╔══════════════════════════════════════════════════════════════╗
    # ║  COLE O CAMINHO DO EXECUTÁVEL DO NAVEGADOR (entre aspas):  ║
    # ║                                                              ║
    # ║  Edge:                                                       ║
    # ║  C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe║
    # ║                                                              ║
    # ║  Chrome:                                                     ║
    # ║  C:\Program Files\Google\Chrome\Application\chrome.exe      ║
    # ╚══════════════════════════════════════════════════════════════╝
    BROWSER_PATH = configpbi.browser

    # ╔══════════════════════════════════════════════════════════════╗
    # ║  NOME DA ABA/PÁGINA PARA NAVEGAR (entre aspas):            ║
    # ║  Deixe vazio "" para ficar na página inicial                ║
    # ╚══════════════════════════════════════════════════════════════╝
    TARGET_PAGE = "COMPARATIVO"

    # ── Validações ──
    if "COLE_SEU_LINK" in POWERBI_URL:
        log.error("❌ Cole o link do Power BI na variável POWERBI_URL!")
        sys.exit(1)

    if "COLE_O_CAMINHO" in BROWSER_PATH:
        log.error("❌ Cole o caminho do navegador na variável BROWSER_PATH!")
        sys.exit(1)

    if not os.path.exists(BROWSER_PATH):
        log.error(f"❌ Navegador não encontrado: {BROWSER_PATH}")
        sys.exit(1)

    success = uc.loop().run_until_complete(
        run_export(POWERBI_URL, BROWSER_PATH, TARGET_PAGE)
    )
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
