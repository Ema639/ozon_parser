#!/usr/bin/env python3
# noinspection SpellCheckingInspection
"""
Генерирует Word-отчёт по результатам тестового задания.
"""

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import sqlite3
from pathlib import Path

DB_PATH     = Path(__file__).parent / "prices.db"
REPORT_PATH = Path(__file__).parent / "Отчёт_тестовое_задание.docx"


# ── Helpers ──────────────────────────────────────────────────────────────────
def disable_spell_check(doc):
    """
    Отключает проверку орфографии:
    - w:lang zxx на каждом run (работает в Word Online)
    - w:noProof на каждом run
    - hideSpellingErrors / hideGrammaticalErrors в настройках документа (desktop Word)
    """
    def _mark_run(run):
        rPr = run._r.get_or_add_rPr()
        lang = OxmlElement("w:lang")
        lang.set(qn("w:val"), "zxx")
        lang.set(qn("w:eastAsia"), "zxx")
        lang.set(qn("w:bidi"), "zxx")
        rPr.append(lang)
        rPr.append(OxmlElement("w:noProof"))

    for para in doc.paragraphs:
        for run in para.runs:
            _mark_run(run)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        _mark_run(run)

    settings = doc.settings.element
    settings.append(OxmlElement("w:hideSpellingErrors"))
    settings.append(OxmlElement("w:hideGrammaticalErrors"))


def set_cell_bg(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def heading(doc, text: str, level: int = 1):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p


def add_table_row(table, values, bold=False, bg=None):
    row = table.add_row()
    for i, val in enumerate(values):
        cell = row.cells[i]
        cell.text = str(val)
        run = cell.paragraphs[0].runs[0]
        run.font.size = Pt(10)
        run.font.bold = bold
        if bg:
            set_cell_bg(cell, bg)
    return row


# ── Main ──────────────────────────────────────────────────────────────────────
def build_report():
    doc = Document()

    # Заголовок документа
    title = doc.add_heading("Тестовое задание: Аналитик данных", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(
        "Восстановление цен Ozon и статистический анализ отклонений"
    ).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # ── 1. Описание задачи ────────────────────────────────────────────────────
    heading(doc, "1. Описание задачи")
    doc.add_paragraph(
        "По условиям договора с Ozon цены на товары на Ozon и Wildberries (WB) "
        "должны быть приблизительно равны. Ozon прислал список товаров, "
        "утверждая, что они стоят дешевле на WB.\n\n"
        "Однако большинство ссылок в файле указывают на похожие, а не идентичные "
        "товары. В столбце TRUE ItemURL приведены ссылки на корректные аналоги.\n\n"
        "Задача: спарсить утерянные цены из столбца Ozon ItemPrice и провести "
        "статистический анализ."
    )

    # ── 2. Подход к парсингу ──────────────────────────────────────────────────
    heading(doc, "2. Подход к парсингу")

    heading(doc, "2.1 Почему не requests / headless браузер", level=2)
    doc.add_paragraph(
        "Ozon использует систему защиты от ботов DataDome. При попытке обратиться к "
        "страницам через requests или headless Chromium/Playwright сервер возвращает "
        "редирект и страницу-вызов «Antibot Challenge Page», которую "
        "headless-браузер не может пройти (DataDome проверяет canvas fingerprint, "
        "WebGL, navigator.webdriver и т.д.)."
    )

    heading(doc, "2.2 Решение: camoufox (headless Firefox с антифингерпринт-патчами)", level=2)
    doc.add_paragraph(
        "Задание прямо допускает использование любых подходов, включая «python кликер "
        "по координатам экрана». В данном случае использован headless-браузер, "
        "который DataDome не отличает от настоящего:\n\n"
        "• camoufox — форк Firefox с патчами против bot-detection.\n"
        "• Корректный TLS/JA3 fingerprint (не отличается от обычного Firefox).\n"
        "• navigator.webdriver = false, нормальный canvas/WebGL fingerprint.\n"
        "• Один экземпляр браузера на весь прогон: прогрев на главной странице Ozon "
        "(DataDome выдаёт сессионные cookies), затем последовательный обход всех URL.\n"
        "• После загрузки ждём 7 секунд (достаточно для рендера JS-виджетов).\n"
        "• Полный HTML читается через page.content(), цена парсится regex-ом."
    )

    heading(doc, "2.3 Как извлекается зелёная цена (Ozon Банк)", level=2)
    doc.add_paragraph(
        "На странице товара Ozon цена с Ozon Банком отображается в виджете "
        "webPrice. В HTML это выглядит так:\n\n"
        "    <span class=\"tsHeadline600Large\">3 074 ₽</span>\n"
        "    <span class=\"...\">с Ozon Банком</span>\n\n"
        "Парсер ищет секцию webPrice и извлекает число перед подписью "
        "«с Ozon Банком». Если товара нет в наличии и эта подпись отсутствует — "
        "берётся первая разумная цена (50–999 999 ₽) из секции webPrice."
    )

    heading(doc, "2.4 База данных", level=2)
    doc.add_paragraph(
        "Pipeline непрерывный: входной Excel-файл → парсинг → SQLite-база данных.\n\n"
        "В таблицу prices сохраняются поля:\n"
        "  • ozon_id  — ID товара на Ozon\n"
        "  • barcode  — Штрихкод (из Excel; в исходном файле хранится как формула VLOOKUP,\n"
        "               поэтому при чтении через openpyxl значение None — это ожидаемо)\n"
        "  • true_price  — TRUE ItemPrice (цена WB по корректной ссылке)\n"
        "  • ozon_price  — Ozon ItemPrice (спарсенная зелёная цена)\n\n"
        "SQLite выбран вместо корпоративной СУБД, т.к. корпоративный VPN недоступен. "
        "Логика загрузки идентична — таблица пересоздаётся при каждом запуске."
    )

    # ── 3. Результаты парсинга ────────────────────────────────────────────────
    heading(doc, "3. Результаты парсинга")

    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT ozon_id, comp_price, true_price, ozon_price FROM prices ORDER BY rowid"
    ).fetchall()
    conn.close()

    headers = ["Ozon ID", "Competitor ₽\n(WB некорр.)", "TRUE ₽\n(WB корр.)", "Ozon ₽\n(спарсено)"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"

    hdr_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.text = h
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.size = Pt(9)
        set_cell_bg(cell, "D0E4FF")

    for r in rows:
        add_table_row(table, [r[0], r[1], r[2], r[3] if r[3] else "N/A"])

    doc.add_paragraph()

    # ── 4. Статистический анализ ──────────────────────────────────────────────
    heading(doc, "4. Статистический анализ")

    heading(doc, "4.1 Методология", level=2)
    doc.add_paragraph(
        "Каждая строка в таблице — один товар, для которого есть цена на WB и на Ozon. "
        "Это парные (зависимые) наблюдения, поэтому применяется:\n\n"
        "• Критерий Шапиро–Уилка — проверяем нормальность разностей цен.\n"
        "• Парный t-тест — основной тест при нормальном распределении разностей.\n"
        "• Критерий Вилкоксона (знаковых рангов) — непараметрический аналог, "
        "не требует нормальности (применяется как дополнительная проверка или "
        "основной тест при ненормальных данных).\n"
        "• Cohen's d — размер эффекта: малый (<0.5), средний (0.5–0.8), большой (>0.8).\n\n"
        "Уровень значимости α = 0.05. N = 20."
    )

    heading(doc, "4.2 Сравнение A: Competitor ItemPrice vs Ozon ItemPrice", level=2)
    doc.add_paragraph(
        "Competitor ItemPrice — цены с WB по НЕКОРРЕКТНЫМ ссылкам "
        "(похожие, но не идентичные товары — как правило, дешевле).\n\n"
        "Результаты:\n"
        "  • Среднее WB (некорр.) = 998 ₽    Среднее Ozon = 3 107 ₽\n"
        "  • Средняя разница = −2 109 ₽  (σ = 1 739)\n"
        "  • Shapiro–Wilk: W = 0.9015, p = 0.044 → разности НЕ нормальные\n"
        "  • Парный t-тест: t = −5.425, p = 0.000031  *** ЗНАЧИМО\n"
        "  • Тест Вилкоксона: W = 0.0, p = 0.000088   *** ЗНАЧИМО\n"
        "  • Cohen's d = −1.21 (большой эффект)\n\n"
        "ВЫВОД: Статистически значимое отличие ЕСТЬ (p < 0.05).\n"
        "Ozon значимо дороже похожих (но некорректных) товаров на WB — "
        "однако это сравнение некорректно по условию задачи, "
        "поскольку ссылки ведут на разные товары."
    )

    heading(doc, "4.3 Сравнение B: TRUE ItemPrice vs Ozon ItemPrice", level=2)
    doc.add_paragraph(
        "TRUE ItemPrice — цены с WB по КОРРЕКТНЫМ ссылкам (абсолютные аналоги).\n\n"
        "Результаты:\n"
        "  • Среднее WB (корр.) = 2 950 ₽    Среднее Ozon = 3 107 ₽\n"
        "  • Средняя разница = −158 ₽  (σ = 483)\n"
        "  • Shapiro–Wilk: p < 0.001 → разности НЕ нормальные\n"
        "  • Парный t-тест: t = −1.460, p = 0.161\n"
        "  • Тест Вилкоксона: W = 73.0, p = 0.868\n"
        "  • Cohen's d = −0.33 (малый эффект)\n\n"
        "ВЫВОД: Статистически значимого отличия НЕТ (p > 0.05).\n"
        "При сравнении с КОРРЕКТНЫМИ аналогами цены на Ozon и WB статистически равны — "
        "претензии Ozon не подтверждаются."
    )

    heading(doc, "4.4 Интерпретация и оговорки", level=2)
    doc.add_paragraph(
        "1. Цены на маркетплейсах динамичны — они менялись с момента создания файла "
        "(февраль) до момента парсинга. Абсолютные числа не имеют значения; "
        "важна логика рассуждений.\n\n"
        "2. Сравнение A показывает значимую разницу, но оно методологически некорректно: "
        "некорректные ссылки ведут на другие (более дешёвые) товары, "
        "поэтому этот результат нельзя использовать как аргумент.\n\n"
        "3. N=20 — небольшая выборка. Тем не менее оба теста в сравнении B "
        "дают p > 0.05, что говорит об устойчивости результата.\n\n"
        "4. Главный вывод: претензия Ozon основана на некорректных ссылках. "
        "Корректное сравнение (сравнение B) показывает, что цены статистически равны."
    )

    # ── 5. Выводы ─────────────────────────────────────────────────────────────
    heading(doc, "5. Итоговые выводы")
    doc.add_paragraph(
        "• Цены Ozon восстановлены для всех 20 товаров (100% успех).\n"
        "• Подход: camoufox (headless Firefox) обходит DataDome без сторонних прокси.\n"
        "• Сравнение A (некорр. WB vs Ozon): отличие ЗНАЧИМО (p < 0.001), "
        "но методологически некорректно — ссылки ведут на разные товары.\n"
        "• Сравнение B (корр. WB vs Ozon):   отличие НЕ значимо (p = 0.87) — "
        "цены на корректных аналогах статистически равны.\n"
        "• Претензии Ozon не подтверждаются: при корректном сравнении разницы нет."
    )

    disable_spell_check(doc)
    doc.save(str(REPORT_PATH))
    print(f"Отчёт сохранён: {REPORT_PATH}")


if __name__ == "__main__":
    build_report()
