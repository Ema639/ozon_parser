#!/usr/bin/env python3
"""
Ozon Price Parser — Тестовое задание аналитик данных

Pipeline: Excel -> Ozon парсинг через camoufox -> SQLite -> Статистика

Подход:
  - Ozon использует DataDome (антибот). Прямые requests/headless Chromium блокируются.
  - Решение: camoufox — Firefox с антифингерпринт-патчами. DataDome не отличает его
    от реального браузера (корректный JA3/TLS, нет navigator.webdriver, нормальный canvas).
  - Один экземпляр браузера на весь прогон — быстрее и экономит ресурсы.
  - Из HTML извлекаем зелёную цену (с Ozon Банком) из виджета webPrice.
"""

import re
import sqlite3
from pathlib import Path

import openpyxl
import numpy as np
from scipy import stats

# ──────────────────────────────────────────────────────────────────────────────
EXCEL_PATH = Path(__file__).parent / "ИНДЕКС ЦЕН - ТЕСТОВОЕ ЗАДАНИЕ АНАЛИТИК ДАННЫХ.xlsx"
DB_PATH    = Path(__file__).parent / "prices.db"

# Пауза после загрузки страницы (мс) — ждём рендера JS-виджетов
PAGE_LOAD_MS = 7000


# ──────────────────────────────────────────────────────────────────────────────
# ШАГ 1: ЧТЕНИЕ EXCEL
# ──────────────────────────────────────────────────────────────────────────────
def read_excel(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rows.append({
            "ozon_id":    str(row[1]),   # B: Rezon ID
            "barcode":    row[2],        # C: Штрихкод (формула -> None при values_only)
            "wb_sku":     row[3],        # D: WB SKU
            "ozon_link":  row[4],        # E: RezonLink (https://www.ozon.ru/product/{id})
            "comp_price": row[15],       # P: Competitor ItemPrice (WB некорректная ссылка)
            "true_price": row[16],       # Q: TRUE ItemPrice (WB корректная ссылка)
            "ozon_price": None,          # R: Ozon ItemPrice — нужно спарсить
            "out_of_stock": False,
        })
    print(f"[Excel] Прочитано {len(rows)} строк")
    return rows


# ──────────────────────────────────────────────────────────────────────────────
# ШАГ 2: БРАУЗЕР CAMOUFOX
# ──────────────────────────────────────────────────────────────────────────────
def make_browser():
    """Запускает camoufox headless и возвращает (browser, page)."""
    from camoufox.sync_api import Camoufox
    browser = Camoufox(headless=True).__enter__()
    page = browser.new_page()
    # Прогреваем сессию на главной — получаем нужные cookies
    page.goto("https://www.ozon.ru/", timeout=30000, wait_until="commit")
    page.wait_for_timeout(3000)
    return browser, page


def get_page_source(page, url: str) -> str | None:
    """Переходит по URL и возвращает HTML после рендера JS."""
    try:
        page.goto(url, timeout=30000, wait_until="commit")
        page.wait_for_timeout(PAGE_LOAD_MS)
        html = page.content()
        if len(html) < 5000:
            return None
        return html
    except Exception as e:
        print(f"  [camoufox] Ошибка: {e}")
        return None


# ──────────────────────────────────────────────────────────────────────────────
# ШАГ 3: ПАРСИНГ ЦЕНЫ ИЗ HTML
# ──────────────────────────────────────────────────────────────────────────────
def _strip_tags(html: str) -> str:
    return re.sub(r'<[^>]+>', ' ', html)


def parse_out_of_stock_price(html: str) -> float | None:
    """
    Если товар закончился, Ozon показывает блок «Этот товар закончился»
    с последней ценой. Ищем цену в окне ±600 символов вокруг этой фразы.
    """
    text = _strip_tags(html)
    m = re.search(r'товар\s+закончился', text, re.I)
    if not m:
        return None
    window = text[max(0, m.start() - 100): m.end() + 600]
    prices = re.findall(r'([\d\u2009\s]+)\s*₽', window)
    for p in prices:
        raw = re.sub(r'[^\d]', '', p)
        if raw and 50 <= int(raw) <= 999999:
            return float(raw)
    return None


def parse_ozon_bank_price(html: str) -> float | None:
    """
    Извлекает зелёную цену «с Ozon Банком» из исходника страницы.

    Структура в HTML (виджет webPrice):
      <span class="tsHeadline600Large">3 074 ₽</span> ... с Ozon Банком

    Если цена с Ozon Банком не найдена — берём первую разумную цену из webPrice.
    """
    # Вырезаем секцию виджета webPrice
    wp_match = re.search(r'webPrice.*?(?=webInstallment|webAddToCart|<footer|$)', html, re.S)
    if not wp_match:
        return _fallback_price(html)

    section_html = wp_match.group(0)[:5000]
    section_text = _strip_tags(section_html)  # plain text без тегов

    # Приоритет 1: число перед ₽, после которого в тексте стоит "с Ozon Банком"
    m = re.search(
        r'([\d\u2009\s]+)\s*₽.{0,120}с\s+Ozon\s+Банком',
        section_text,
        re.S,
    )
    if m:
        raw = re.sub(r'[^\d]', '', m.group(1))
        if raw:
            return float(raw)

    # Приоритет 2: любая цена в секции webPrice (например, товар не в наличии)
    prices = re.findall(r'([\d\u2009\s]+)\s*₽', section_text)
    for p in prices:
        raw = re.sub(r'[^\d]', '', p)
        if raw and 50 <= int(raw) <= 999999:
            return float(raw)

    return _fallback_price(html)


def _fallback_price(html: str) -> float | None:
    """Запасной поиск по всему HTML — берём наименьшую разумную цену."""
    prices = re.findall(r'([\d\u2009\s]+)\s*₽', html)
    candidates = []
    for p in prices:
        raw = re.sub(r'[^\d]', '', p)
        if raw and 50 <= int(raw) <= 999999:
            candidates.append(int(raw))
    if candidates:
        # Ozon Bank price обычно меньше обычной цены
        return float(min(candidates))
    return None


# ──────────────────────────────────────────────────────────────────────────────
# ШАГ 4: ОСНОВНОЙ ЦИКЛ ПАРСИНГА
# ──────────────────────────────────────────────────────────────────────────────
def fetch_all_prices(rows: list[dict]) -> list[dict]:
    from camoufox.sync_api import Camoufox

    total = len(rows)
    print("[Browser] Запускаем camoufox...", flush=True)

    with Camoufox(headless=True) as browser:
        page = browser.new_page()

        # Прогрев: главная страница — DataDome выдаёт сессионные cookies
        print("[Browser] Прогрев на главной ozon.ru...", flush=True)
        page.goto("https://www.ozon.ru/", timeout=30000, wait_until="commit")
        page.wait_for_timeout(3000)

        for i, row in enumerate(rows):
            url = row["ozon_link"]
            print(f"[{i+1}/{total}] {url} ...", end=" ", flush=True)

            html = get_page_source(page, url)
            if html is None:
                print("✗ страница не загрузилась")
                continue

            oos_price = parse_out_of_stock_price(html)
            if oos_price is not None:
                row["ozon_price"] = oos_price
                row["out_of_stock"] = True
                print(f"~ {oos_price:.0f} ₽ (нет в наличии, последняя цена)")
                continue

            price = parse_ozon_bank_price(html)
            if price is not None:
                row["ozon_price"] = price
                print(f"✓ {price:.0f} ₽")
            else:
                print("✗ цена не найдена")

    return rows


# ──────────────────────────────────────────────────────────────────────────────
# ШАГ 5: СОХРАНЕНИЕ В SQLITE
# ──────────────────────────────────────────────────────────────────────────────
def save_to_db(rows: list[dict], db_path: Path):
    """
    Загружает в SQLite три нужных столбца: Штрихкод, TRUE ItemPrice, Ozon ItemPrice.
    Таблица каждый раз пересоздаётся (по условию задания).
    """
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # Пересоздаём таблицу
    cur.execute("DROP TABLE IF EXISTS prices")
    cur.execute("""
        CREATE TABLE prices (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            ozon_id        TEXT    NOT NULL,
            barcode        TEXT,
            wb_sku         TEXT,
            comp_price     REAL,
            true_price     REAL,
            ozon_price     REAL,
            out_of_stock   INTEGER NOT NULL DEFAULT 0
        )
    """)

    for row in rows:
        cur.execute(
            """INSERT INTO prices
               (ozon_id, barcode, wb_sku, comp_price, true_price, ozon_price, out_of_stock)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (
                row["ozon_id"],
                str(row["barcode"]) if row["barcode"] else None,
                str(row["wb_sku"])  if row["wb_sku"]  else None,
                row["comp_price"],
                row["true_price"],
                row["ozon_price"],
                1 if row["out_of_stock"] else 0,
            ),
        )

    conn.commit()

    # Проверка
    count = cur.execute("SELECT COUNT(*) FROM prices").fetchone()[0]
    conn.close()
    print(f"\n[DB] Сохранено {count} строк → {db_path}")


# ──────────────────────────────────────────────────────────────────────────────
# ШАГ 6: СТАТИСТИЧЕСКИЙ АНАЛИЗ
# ──────────────────────────────────────────────────────────────────────────────
def statistical_analysis(rows: list[dict]) -> None:
    """
    Проверяет: есть ли статистически значимое отличие между
      A) Competitor ItemPrice и Ozon ItemPrice
      B) TRUE ItemPrice и Ozon ItemPrice

    Используем:
      - Shapiro-Wilk для проверки нормальности разностей
      - Парный t-тест (если нормальные) + критерий Вилкоксона (непараметрический)
      - Cohen's d как мера размера эффекта
    """
    valid = [r for r in rows if r["ozon_price"] is not None]
    n = len(valid)

    if n < 3:
        print(f"\n[Stats] Слишком мало данных: {n} строк с ценой Ozon. Нужно хотя бы 3.")
        return

    comp  = np.array([r["comp_price"]  for r in valid], dtype=float)
    true_ = np.array([r["true_price"]  for r in valid], dtype=float)
    ozon  = np.array([r["ozon_price"]  for r in valid], dtype=float)

    alpha = 0.05

    print("\n" + "=" * 65)
    print("  СТАТИСТИЧЕСКИЙ АНАЛИЗ")
    print("=" * 65)
    print(f"  N = {n}  (строк с известной ценой Ozon)")

    for label, arr_a, arr_b in [
        ("Competitor ItemPrice  vs  Ozon ItemPrice", comp,  ozon),
        ("TRUE ItemPrice        vs  Ozon ItemPrice", true_, ozon),
    ]:
        diff = arr_a - arr_b

        mean_a  = arr_a.mean()
        mean_b  = arr_b.mean()
        mean_d  = diff.mean()
        std_d   = diff.std(ddof=1)

        # Нормальность разностей
        if n >= 3:
            sw_stat, sw_p = stats.shapiro(diff)
            normal = sw_p > alpha
        else:
            sw_stat, sw_p, normal = 0, 1, True

        # Парный t-тест
        t_stat, t_p = stats.ttest_rel(arr_a, arr_b)

        # Тест Вилкоксона (непараметрический, знаковых рангов)
        try:
            _w = stats.wilcoxon(arr_a, arr_b, alternative="two-sided")
            w_stat, w_p = float(_w[0]), float(_w[1])
        except ValueError:
            w_stat, w_p = float("nan"), float("nan")

        # Размер эффекта Cohen's d
        cohens_d = mean_d / std_d if std_d > 0 else 0.0

        sig_t = t_p < alpha
        sig_w = w_p < alpha if not np.isnan(w_p) else False

        print(f"\n--- {label} ---")
        print(f"  Среднее A = {mean_a:,.0f} ₽    Среднее B = {mean_b:,.0f} ₽")
        print(f"  Средняя разница (A−B) = {mean_d:+,.0f} ₽  (σ = {std_d:,.0f})")
        print(f"  Shapiro-Wilk: W={sw_stat:.4f}, p={sw_p:.4f}  "
              f"({'нормальное' if normal else 'НЕ нормальное'})")
        print(f"  Парный t-тест:     t={t_stat:+.3f}, p={t_p:.6f}  "
              f"{'*** ЗНАЧИМО' if sig_t else 'не значимо'}")
        if not np.isnan(w_p):
            print(f"  Тест Вилкоксона:   W={w_stat:.1f},  p={w_p:.6f}  "
                  f"{'*** ЗНАЧИМО' if sig_w else 'не значимо'}")
        print(f"  Cohen's d = {cohens_d:.3f}  "
              f"({'большой' if abs(cohens_d)>=0.8 else 'средний' if abs(cohens_d)>=0.5 else 'малый'})")

        # Вывод
        significant = sig_t or sig_w
        print(f"\n  ВЫВОД: {'Статистически значимое отличие ЕСТЬ (p < 0.05)' if significant else 'Статистически значимого отличия НЕТ (p ≥ 0.05)'}")

    print("\n" + "=" * 65)


# ──────────────────────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ──────────────────────────────────────────────────────────────────────────────
def main():
    print("=" * 65)
    print("  OZON PRICE PARSER  —  старт pipeline")
    print("=" * 65)

    # 1. Читаем Excel
    rows = read_excel(EXCEL_PATH)

    # 2. Парсим цены Ozon через camoufox
    rows = fetch_all_prices(rows)

    # 3. Сохраняем в SQLite
    save_to_db(rows, DB_PATH)

    # 4. Статистический анализ
    statistical_analysis(rows)

    # 5. Итоговая сводка
    print("\nИТОГОВАЯ ТАБЛИЦА:")
    print(f"{'Ozon ID':<15} {'Comp ₽':>9} {'TRUE ₽':>9} {'Ozon ₽':>16}")
    print("-" * 54)
    for r in rows:
        if r["ozon_price"] is not None:
            oz = f"{r['ozon_price']:.0f}" + (" (нет в нал.)" if r["out_of_stock"] else "")
        else:
            oz = "N/A"
        print(f"{r['ozon_id']:<15} {str(r['comp_price'] or ''):>9} {str(r['true_price'] or ''):>9} {oz:>16}")

    found = sum(1 for r in rows if r["ozon_price"] is not None)
    print(f"\nНайдено цен Ozon: {found}/{len(rows)}")
    print(f"База данных:      {DB_PATH}")


if __name__ == "__main__":
    main()
