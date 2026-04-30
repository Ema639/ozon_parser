# Тестовое задание: Аналитик данных

Восстановление цен Ozon и статистический анализ отклонений.

## Описание

Pipeline: Excel-файл → парсинг цен Ozon через camoufox → SQLite → статистический анализ → Word-отчёт.

## Установка

### 1. Клонировать репозиторий

```bash
git clone git@github.com:Ema639/ozon_parser.git
cd ozon_parser
```

### 2. Создать виртуальное окружение и активировать его

```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
```

### 3. Установить зависимости

```bash
pip install -r requirements.txt
```

### 4. Скачать браузер для camoufox (обязательный шаг)

```bash
python -m camoufox fetch
```

> Без этого шага парсер не запустится — camoufox скачивает модифицированный Firefox (~100 МБ).

## Запуск

### Шаг 1 — Парсинг цен и сохранение в БД

```bash
python ozon_parser.py
```

Скрипт читает `ИНДЕКС ЦЕН - ТЕСТОВОЕ ЗАДАНИЕ АНАЛИТИК ДАННЫХ.xlsx`, парсит цены с Ozon и сохраняет результаты в `prices.db`.

Время выполнения: ~3–5 минут (20 товаров × 7 сек ожидания рендера).

### Шаг 2 — Генерация Word-отчёта

```bash
python generate_report.py
```

Читает данные из `prices.db` и создаёт `Отчёт_тестовое_задание.docx`.

## Структура проекта

```
ozon_parser/
├── ozon_parser.py      # Парсер цен (Excel -> camoufox -> SQLite -> статистика)
├── generate_report.py  # Генератор Word-отчёта
├── requirements.txt    # Зависимости Python
└── ИНДЕКС ЦЕН - ТЕСТОВОЕ ЗАДАНИЕ АНАЛИТИК ДАННЫХ.xlsx  # Входные данные
```

> `prices.db` и `Отчёт_тестовое_задание.docx` генерируются при запуске и не включены в репозиторий.

## Требования

- Python 3.10+
- Интернет-соединение (для парсинга Ozon)
