#!/usr/bin/env python3
"""
Скрипт для моніторингу залишків
Шукає товари заданого бренду (та категорії) у Сільпо, Novus та Varus.
Експортує результати у багатосторінковий Excel файл.
"""

import math
import requests
import subprocess
import json
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("ПОМИЛКА: Відсутня бібліотека bs4. Встановіть її через pip install beautifulsoup4")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ПОМИЛКА: Відсутня бібліотека openpyxl. Встановіть її через pip install openpyxl")
    sys.exit(1)

# ─────────────────────────────────────────────────────────
#  НАЛАШТУВАННЯ
# ─────────────────────────────────────────────────────────
DEFAULT_BRAND = "Торчин"
MAX_PAGES     = 10

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "uk-UA,uk;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
}

XLSX_COLS = [
    ("#",               5),
    ("Товар",          55),
    ("Артикул / SKU",  18),
    ("Ціна (грн)",     14),
    ("Знижка",         12),
    ("Ціна зі знижкою", 18),
    ("Наявність",      12),
    ("Мережа",         20),
    ("Посилання",      65),
]

C_HEADER    = "10B981" # Смарагдовий (як у вашому дизайні)
C_IN_STOCK  = "D1FAE5"
C_OUT_STOCK = "FEE2E2"
C_UNKNOWN   = "F1F5F9"
C_TITLE     = "047857"
C_DISCOUNT  = "FFEDD5"

# ─────────────────────────────────────────────────────────
#  ДОПОМІЖНІ ФУНКЦІЇ
# ─────────────────────────────────────────────────────────

def fetch(url, session):
    try:
        r = session.get(url, headers=REQUEST_HEADERS, timeout=20, allow_redirects=True)
        r.raise_for_status()
        r.encoding = 'utf-8'
        return BeautifulSoup(r.text, "html.parser"), r.text
    except requests.RequestException as e:
        print(f"  ПОПЕРЕДЖЕННЯ: {e}")
        return None, None

def _node_cmd():
    for cmd in ("node", "nodejs"):
        try:
            if subprocess.run([cmd, "--version"], capture_output=True, timeout=5).returncode == 0:
                return cmd
        except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
            pass
    return None

def node_available():
    return _node_cmd() is not None

def decode_nuxt(script_text, log_fn=print):
    cmd = _node_cmd()
    if not cmd: return None
    import tempfile, os
    js = "const window={};\n" + script_text + "\nprocess.stdout.write(JSON.stringify(window.__NUXT__||null));"
    tmp = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False) as f:
            f.write(js)
            tmp = f.name
        result = subprocess.run([cmd, tmp], capture_output=True, text=True, timeout=15)
        if result.returncode == 0 and result.stdout:
            return json.loads(result.stdout)
    except Exception: pass
    finally:
        if tmp:
            try: os.unlink(tmp)
            except OSError: pass
    return None

# ─────────────────────────────────────────────────────────
#  ПАРСЕРИ (СІЛЬПО, NOVUS, VARUS)
# ─────────────────────────────────────────────────────────

def _scrape_zakaz(retailer, query, session, log_fn=print, meta=None):
    """Спільний парсер для Novus та Varus, які працюють/працювали на платформі zakaz.ua"""
    log_fn(f"{retailer.title()}: шукаємо '{query}'...")
    query_enc = requests.utils.quote(query)
    base_url = f"https://{retailer}.zakaz.ua/uk/search/?q={query_enc}"
    
    soup, raw = fetch(base_url, session)
    if not soup:
        log_fn(f"  {retailer.title()}: не вдалося завантажити сторінку.")
        return []

    products = []
    
    # Спроба 1: Витягнути дані через Next.js стан (найточніший метод для zakaz.ua)
    script = soup.find("script", id="__NEXT_DATA__")
    if script:
        try:
            data = json.loads(script.string)
            items = data.get('props', {}).get('pageProps', {}).get('initialState', {}).get('products', {}).get('products', [])
            
            for p in items:
                if not isinstance(p, dict): continue
                name = p.get('title') or p.get('name') or ""
                if not name: continue
                
                sku = str(p.get('ean') or p.get('id') or "")
                
                # Ціни в zakaz.ua часто в копійках
                price_raw = float(p.get('price', 0) or 0)
                if price_raw > 1000: price_raw = price_raw / 100.0
                
                old_price_raw = float(p.get('old_price', 0) or 0)
                if old_price_raw > 1000: old_price_raw = old_price_raw / 100.0
                
                on_discount = bool(old_price_raw and old_price_raw > price_raw)
                
                in_stock = p.get('in_stock', True)
                url_p = f"https://{retailer}.zakaz.ua/uk/products/{sku}/" if sku else base_url
                
                products.append({
                    "name": name, 
                    "sku": sku,
                    "price": str(old_price_raw if on_discount else price_raw), 
                    "on_discount": on_discount,
                    "discount_price": str(price_raw) if on_discount else "",
                    "url": url_p, 
                    "in_stock": in_stock, 
                    "seller": retailer.title()
                })
        except Exception as e:
            log_fn(f"  {retailer.title()}: Помилка розбору JSON: {e}")

    # Спроба 2: Якщо Next.js не знайдено, парсимо HTML напряму
    if not products:
        for card in soup.select('[data-testid="product-tile"]'):
            name_el = card.select_one('[data-testid="product-tile-title"]')
            name = name_el.text.strip() if name_el else ""
            
            link_el = card.select_one('a')
            url_p = f"https://{retailer}.zakaz.ua" + link_el['href'] if link_el and link_el.has_attr('href') else base_url
            
            price_el = card.select_one('[data-testid="product-tile-price"]')
            price_txt = price_el.text.strip() if price_el else "0"
            price = float(re.sub(r'[^\d.]', '', price_txt.replace(',', '.')) or 0)
            
            old_price_el = card.select_one('[data-testid="product-tile-old-price"]')
            old_price = 0
            if old_price_el:
                old_price = float(re.sub(r'[^\d.]', '', old_price_el.text.replace(',', '.')) or 0)
                
            on_discount = old_price > price
            in_stock = bool(card.select_one('[data-testid="add-to-cart-button"]'))
            
            if name:
                products.append({
                    "name": name, "sku": "",
                    "price": str(old_price if on_discount else price),
                    "on_discount": on_discount,
                    "discount_price": str(price) if on_discount else "",
                    "url": url_p, "in_stock": in_stock, "seller": retailer.title()
                })

    log_fn(f"  {retailer.title()}: знайдено {len(products)} товарів.")
    if meta is not None: meta["scraped_total"] = len(products)
    return products

def scrape_novus(query, session, log_fn=print, meta=None):
    return _scrape_zakaz("novus", query, session, log_fn, meta)

def scrape_varus(query, session, log_fn=print, meta=None):
    return _scrape_zakaz("varus", query, session, log_fn, meta)

def scrape_silpo(query, session, has_node, log_fn=print, meta=None):
    log_fn(f"Сільпо: шукаємо '{query}'...")
    query_enc = requests.utils.quote(query)
    url = f"https://silpo.ua/catalog/search?search={query_enc}"
    
    soup, raw = fetch(url, session)
    if not soup:
        log_fn("  Сільпо: не вдалося завантажити сторінку.")
        return []

    products = []
    
    # Парсинг карток товарів Сільпо напряму з HTML (надійний fallback)
    # Сільпо використовує різноманітні класи, додаємо найпоширеніші
    for card in soup.select('.product-card, .product-list-item, li[data-product-id]'):
        name_el = card.select_one('.product-title, .product-card__title, .name')
        name = name_el.text.strip() if name_el else ""
        if not name: continue
        
        sku = card.get('data-product-id') or ""
        
        link_el = card.select_one('a')
        url_p = "https://silpo.ua" + link_el['href'] if link_el and link_el.has_attr('href') else url
        if not url_p.startswith('http'): url_p = url
        
        price_el = card.select_one('.product-price__current, .price, .current-price')
        price_txt = price_el.text.strip() if price_el else "0"
        price = float(re.sub(r'[^\d.]', '', price_txt.replace(',', '.')) or 0)
        
        old_price_el = card.select_one('.product-price__old, .old-price, .strike')
        old_price = 0
        if old_price_el:
            old_price = float(re.sub(r'[^\d.]', '', old_price_el.text.replace(',', '.')) or 0)
            
        on_discount = old_price > price
        
        # Перевірка наявності (кнопка "До кошика")
        btn = card.select_one('button.add-to-cart, .btn-buy')
        out_of_stock_badge = card.select_one('.out-of-stock, .not-available')
        
        in_stock = bool(btn) and not bool(out_of_stock_badge)
        
        products.append({
            "name": name, "sku": sku,
            "price": str(old_price if on_discount else price),
            "on_discount": on_discount,
            "discount_price": str(price) if on_discount else "",
            "url": url_p, "in_stock": in_stock, "seller": "Сільпо"
        })

    # Спроба отримати дані з Nuxt, якщо HTML пустий (що часто буває в SPA)
    if not products and has_node:
        nuxt_data = None
        for script in soup.find_all("script"):
            txt = script.string or ""
            if "window.__NUXT__" in txt:
                nuxt_data = decode_nuxt(txt, log_fn=log_fn)
                break
        
        if nuxt_data:
            try:
                # Навігація по структурі Nuxt Сільпо може відрізнятися, 
                # шукаємо ключі схожі на items або products
                state = nuxt_data.get('state', {})
                items = []
                for k, v in state.items():
                    if isinstance(v, dict) and 'items' in v:
                        items = v['items']
                        break
                
                for p in items:
                    name = p.get('title') or p.get('name') or ""
                    if not name: continue
                    sku = str(p.get('id') or "")
                    price = float(p.get('price', 0))
                    old_price = float(p.get('oldPrice', 0))
                    on_discount = old_price > price
                    in_stock = p.get('status') != 'out_of_stock'
                    url_p = f"https://silpo.ua/product/{p.get('slug')}" if p.get('slug') else url
                    
                    products.append({
                        "name": name, "sku": sku,
                        "price": str(old_price if on_discount else price),
                        "on_discount": on_discount,
                        "discount_price": str(price) if on_discount else "",
                        "url": url_p, "in_stock": in_stock, "seller": "Сільпо"
                    })
            except Exception as e:
                pass

    log_fn(f"  Сільпо: знайдено {len(products)} товарів.")
    if meta is not None: meta["scraped_total"] = len(products)
    return products

# ─────────────────────────────────────────────────────────
#  ПЕРЕВІРКА ЯКОСТІ ТА ЕКСПОРТ
# ─────────────────────────────────────────────────────────

def check_data_quality(products):
    warnings = []
    for p in products:
        name = p.get("name", "")[:50]
        url  = p.get("url", "")
        try: price = float(str(p.get("price", 0)).replace(" ", "").replace(",", "."))
        except (ValueError, TypeError): price = 0
        try: disc = float(str(p.get("discount_price", 0) or 0).replace(" ", "").replace(",", "."))
        except (ValueError, TypeError): disc = 0

        if price == 0 and p.get("in_stock") is not False:
            warnings.append((f"Ціна 0 грн: {name}", url))
        if p.get("on_discount") and disc >= price:
            warnings.append((f"Знижка >= звичайної ціни: {name}", url))
    return warnings

def _fmt_price(raw):
    digits = re.sub(r"[^\d.]", "", str(raw).replace(",", "."))
    try: return "{:,.2f}".format(float(digits)).replace(",", " ")
    except ValueError: return str(raw)

def write_store_sheet(ws, products, store_name, query, checked_at):
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    hdr_fill = PatternFill("solid", fgColor=C_HEADER)
    in_fill  = PatternFill("solid", fgColor=C_IN_STOCK)
    out_fill = PatternFill("solid", fgColor=C_OUT_STOCK)
    unk_fill = PatternFill("solid", fgColor=C_UNKNOWN)
    center   = Alignment(horizontal="center", vertical="center")
    left     = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    last_col = get_column_letter(len(XLSX_COLS))
    ws.merge_cells(f"A1:{last_col}1")
    t = ws["A1"]
    t.value     = f"{store_name}  ·  Пошук: {query}  ·  Перевірено: {checked_at}"
    t.font      = Font(bold=True, size=13, color=C_TITLE)
    t.alignment = center
    ws.row_dimensions[1].height = 26

    for ci, (col_name, col_width) in enumerate(XLSX_COLS, 1):
        c = ws.cell(row=2, column=ci, value=col_name)
        c.font, c.fill, c.alignment = hdr_font, hdr_fill, center
        ws.column_dimensions[get_column_letter(ci)].width = col_width
    ws.freeze_panes = "A3"

    disc_fill = PatternFill("solid", fgColor=C_DISCOUNT)
    aligns = [center, left, center, center, center, center, center, center, left]
    for i, p in enumerate(products, 1):
        row = i + 2
        status, stock_fill = ("Є", in_fill) if p["in_stock"] else ("Немає", out_fill)
        on_disc = p.get("on_discount", False)
        
        vals = [
            i, p.get("name", ""), p.get("sku", ""),
            _fmt_price(p.get("price", "")), "Так" if on_disc else "Ні",
            _fmt_price(p.get("discount_price", "")) if on_disc else "",
            status, p.get("seller", ""), p.get("url", "")
        ]
        
        for ci, (val, aln) in enumerate(zip(vals, aligns), 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.alignment = aln
            if ci in (5, 6) and on_disc: c.fill = disc_fill
            elif ci == 7: c.fill = stock_fill

def write_summary_sheet(ws, results, query, checked_at):
    hdr_font, hdr_fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=C_HEADER)
    center = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = f"Зведення  ·  Пошук: {query}  ·  {checked_at}"
    t.font, t.alignment = Font(bold=True, size=13, color=C_TITLE), center
    ws.row_dimensions[1].height = 26

    headers = ["Мережа", "Всього", "В наявності", "Відсутні", "Зі знижкою", "Невідомо"]
    widths  = [20, 14, 14, 14, 14, 12]
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font, c.fill, c.alignment = hdr_font, hdr_fill, center
        ws.column_dimensions[get_column_letter(ci)].width = w

    for ri, (store, products) in enumerate(results.items(), 3):
        in_n   = sum(1 for p in products if p["in_stock"] is True)
        out_n  = sum(1 for p in products if p["in_stock"] is False)
        disc_n = sum(1 for p in products if p.get("on_discount"))
        unk_n  = sum(1 for p in products if p["in_stock"] is None)
        
        for ci, val in enumerate([store, len(products), in_n, out_n, disc_n, unk_n], 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.alignment = center
            if ci == 3 and in_n > 0: c.fill = PatternFill("solid", fgColor=C_IN_STOCK)
            if ci == 4 and out_n > 0: c.fill = PatternFill("solid", fgColor=C_OUT_STOCK)
            if ci == 5 and disc_n > 0: c.fill = PatternFill("solid", fgColor=C_DISCOUNT)