#!/usr/bin/env python3
"""
Скрипт для моніторингу залишків (Сільпо, Novus, Varus)
Повністю побудований на оригінальній архітектурі (Node.js Nuxt Decoder, Fetch Wrapper)
"""

import math
import requests
import subprocess
import time
import json
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("ERROR: Missing package bs4.")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    pass

# ─────────────────────────────────────────────────────────
#  НАЛАШТУВАННЯ ТА СТИЛІ
# ─────────────────────────────────────────────────────────
MAX_PAGES = 10

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8,en;q=0.6",
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

C_HEADER    = "10B981" 
C_IN_STOCK  = "D1FAE5"
C_OUT_STOCK = "FEE2E2"
C_UNKNOWN   = "F1F5F9"
C_TITLE     = "047857"
C_DISCOUNT  = "FFEDD5"

# ─────────────────────────────────────────────────────────
#  ОРИГІНАЛЬНІ ФУНКЦІЇ З ВАШОГО КОДУ (Node.js, Fetch, Bypass)
# ─────────────────────────────────────────────────────────

def fetch(url, session):
    """Оригінальна функція fetch з обробкою помилок"""
    try:
        r = session.get(url, headers=REQUEST_HEADERS, timeout=25, allow_redirects=True)
        # Додаємо обхід простого блокування, як було в Organic Market
        if len(r.text) < 2000 and "challenge_passed" in r.text:
            m = re.search(r'defaultHash\s*=\s*"([a-f0-9]+)"', r.text)
            if m:
                session.cookies.set("challenge_passed", m.group(1), domain=urllib.parse.urlparse(url).netloc)
                r = session.get(url, headers=REQUEST_HEADERS, timeout=25, allow_redirects=True)
                
        r.raise_for_status()
        r.encoding = 'utf-8'
        return BeautifulSoup(r.text, "html.parser"), r.text
    except requests.RequestException as e:
        return None, str(e)

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
    """Оригінальний декодер Nuxt через Node.js"""
    import tempfile, os
    cmd = _node_cmd()
    if not cmd:
        return None
    js = "const window={};\n" + script_text + "\nprocess.stdout.write(JSON.stringify(window.__NUXT__||null));"
    tmp = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False) as f:
            f.write(js)
            tmp = f.name
        result = subprocess.run([cmd, tmp], capture_output=True, text=True, timeout=30)
        if result.returncode == 0 and result.stdout:
            return json.loads(result.stdout)
    except Exception:
        pass
    finally:
        if tmp:
            try: os.unlink(tmp)
            except OSError: pass
    return None

def _resolve_nuxt(arr, idx, depth=0):
    """Оригінальний дереференсер для Nuxt 3 масивів (як у Eva)"""
    if depth > 30 or not isinstance(idx, int) or idx < 0 or idx >= len(arr):
        return None
    val = arr[idx]
    if isinstance(val, list) and len(val) == 2 and isinstance(val[0], str) and isinstance(val[1], int):
        return _resolve_nuxt(arr, val[1], depth + 1)
    if isinstance(val, dict):
        return {k: (_resolve_nuxt(arr, v, depth + 1) if isinstance(v, int) else v) for k, v in val.items()}
    if isinstance(val, list):
        return [_resolve_nuxt(arr, v, depth + 1) if isinstance(v, int) else v for v in val]
    return val

# Універсальний пошук товарів у будь-якому складному JSON
def extract_products_from_json(node, found_products):
    if isinstance(node, dict):
        # Ознаки товару: є ціна і назва
        if "price" in node and ("title" in node or "name" in node or "name_ua" in node):
            if "sku" in node or "id" in node or "ean" in node or "slug" in node:
                # Уникаємо додавання дублікатів
                sku = str(node.get("id") or node.get("sku") or node.get("ean") or node.get("slug") or "")
                if sku and sku not in found_products:
                    found_products[sku] = node
        for v in node.values():
            extract_products_from_json(v, found_products)
    elif isinstance(node, list):
        for v in node:
            extract_products_from_json(v, found_products)

# ─────────────────────────────────────────────────────────
#  ПАРСЕР СІЛЬПО (Використовує Node.js + _resolve_nuxt)
# ─────────────────────────────────────────────────────────

def scrape_silpo(query, session, has_node, log_fn=print, meta=None):
    if not has_node:
        log_fn("  Сільпо: Node.js не знайдено! Парсинг може бути неповним.")

    log_fn(f"Сільпо: шукаємо '{query}'...")
    query_enc = requests.utils.quote(query)
    
    # Спробуємо два різні ендпоінти, які використовує Сільпо
    urls_to_try = [
        f"https://silpo.ua/search?find={query_enc}",
        f"https://silpo.ua/catalog/search?search={query_enc}"
    ]
    
    products_data = {}
    
    for url in urls_to_try:
        soup, raw_text = fetch(url, session)
        if not soup:
            log_fn(f"  Сільпо: Помилка доступу до {url} ({raw_text})")
            continue
            
        # Спосіб 1: Шукаємо Nuxt 3 дані (як у Eva)
        tag = soup.find("script", id="__NUXT_DATA__")
        if tag:
            try:
                arr = json.loads(tag.string)
                resolved = [_resolve_nuxt(arr, i) for i in range(min(15, len(arr)))]
                extract_products_from_json(resolved, products_data)
            except Exception as e:
                log_fn(f"  Сільпо: Помилка __NUXT_DATA__ - {e}")
                
        # Спосіб 2: Якщо сайт віддав класичний Nuxt 2 (через Node.js)
        for script in soup.find_all("script"):
            txt = script.string or ""
            if "window.__NUXT__" in txt:
                nuxt_data = decode_nuxt(txt, log_fn)
                if nuxt_data:
                    extract_products_from_json(nuxt_data, products_data)

        # Спосіб 3: Прямий парсинг HTML карток, якщо JSON заблоковано
        for card in soup.select('.product-card, li[data-product-id]'):
            name_el = card.select_one('.product-title, .name')
            if not name_el: continue
            
            sku = card.get('data-product-id') or ""
            price_el = card.select_one('.product-price__current, .price')
            price_txt = price_el.text if price_el else "0"
            price = float(re.sub(r'[^\d.]', '', price_txt.replace(',', '.')) or 0)
            
            if sku and sku not in products_data:
                products_data[sku] = {"title": name_el.text.strip(), "price": price, "slug": sku, "html_parsed": True}

    # Форматуємо знайдені товари
    final_products = []
    for p in products_data.values():
        name = p.get("title") or p.get("name") or p.get("name_ua") or ""
        sku = str(p.get("id") or p.get("sku") or p.get("slug") or p.get("ean") or "")
        
        price = float(p.get("price", 0))
        old_price = float(p.get("oldPrice") or p.get("old_price") or p.get("price_old") or 0)
        
        on_discount = old_price > price
        reg_price = old_price if on_discount else price
        disc_price = price if on_discount else 0
        
        in_stock = str(p.get("status")) != "out_of_stock"
        url_p = f"https://silpo.ua/product/{p.get('slug')}" if p.get("slug") else f"https://silpo.ua/search?find={query_enc}"
        
        final_products.append({
            "name": name, "sku": sku,
            "price": str(reg_price), "on_discount": on_discount,
            "discount_price": str(disc_price) if on_discount else "",
            "url": url_p, "in_stock": in_stock, "seller": "Сільпо"
        })

    log_fn(f"  Сільпо: знайдено {len(final_products)} товарів.")
    return final_products

# ─────────────────────────────────────────────────────────
#  ПАРСЕР ZAKAZ (NOVUS / VARUS)
# ─────────────────────────────────────────────────────────

def scrape_zakaz(retailer, query, session, log_fn=print, meta=None):
    log_fn(f"{retailer.title()}: шукаємо '{query}'...")
    query_enc = requests.utils.quote(query)
    base_url = f"https://{retailer}.zakaz.ua/uk/search/?q={query_enc}"
    
    soup, raw_text = fetch(base_url, session)
    if not soup:
        log_fn(f"  {retailer.title()}: Помилка завантаження сторінки ({raw_text})")
        return []
        
    products_data = {}
    
    # Zakaz використовує Next.js
    script = soup.find("script", id="__NEXT_DATA__")
    if script:
        try:
            data = json.loads(script.string)
            extract_products_from_json(data, products_data)
        except Exception as e:
            log_fn(f"  {retailer.title()}: Помилка розбору JSON: {e}")
            
    # Fallback: HTML парсинг
    if not products_data:
        for card in soup.select('[data-testid="product-tile"]'):
            name_el = card.select_one('[data-testid="product-tile-title"]')
            if not name_el: continue
            
            price_el = card.select_one('[data-testid="product-tile-price"]')
            price_txt = price_el.text if price_el else "0"
            price = float(re.sub(r'[^\d.]', '', price_txt.replace(',', '.')) or 0)
            
            name = name_el.text.strip()
            if name not in products_data:
                products_data[name] = {"title": name, "price": price * 100} # Закладаємо формат копійок

    final_products = []
    for p in products_data.values():
        name = p.get("title") or p.get("name") or ""
        sku = str(p.get("ean") or p.get("id") or "")
        
        # Ціни в Zakaz часто в копійках
        price_raw = float(p.get("price", 0))
        if price_raw > 1000: price_raw = price_raw / 100.0
        
        old_price_raw = float(p.get("old_price") or p.get("discount_price") or 0)
        if old_price_raw > 1000: old_price_raw = old_price_raw / 100.0
        
        on_discount = bool(old_price_raw and old_price_raw > price_raw)
        reg_price = old_price_raw if on_discount else price_raw
        disc_price = price_raw if on_discount else 0
        
        in_stock = p.get("in_stock", True)
        url_p = f"https://{retailer}.zakaz.ua/uk/products/{sku}/" if sku else base_url
        
        final_products.append({
            "name": name, "sku": sku,
            "price": str(reg_price), "on_discount": on_discount,
            "discount_price": str(disc_price) if on_discount else "",
            "url": url_p, "in_stock": in_stock, "seller": retailer.title()
        })
            
    log_fn(f"  {retailer.title()}: знайдено {len(final_products)} товарів.")
    return final_products

def scrape_novus(query, session, log_fn=print, meta=None):
    return scrape_zakaz("novus", query, session, log_fn, meta)

def scrape_varus(query, session, log_fn=print, meta=None):
    return scrape_zakaz("varus", query, session, log_fn, meta)

# ─────────────────────────────────────────────────────────
#  ЕКСПОРТ (Як було у вас)
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
