import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).parent

COLORS = {
    'header_bg':   '1F4E79',
    'header_font': 'FFFFFF',
    'city_font':   'A0B4C8',
    'phone_font':  'D4E4F7',
    'tg_font':     '64B5F6',
    'group_bg':    'BDD7EE',
    'group_font':  '1F4E79',
    'alt_row':     'EBF3FB',
}

COL_WIDTHS = {'A': 7.57, 'B': 13.14, 'C': 14.43, 'D': 77.29, 'E': 16.0, 'F': 15.57}


def load_config():
    with open(BASE_DIR / 'config.json', encoding='utf-8') as f:
        return json.load(f)


def _side():
    return Side(style='thin', color='B8CCE4')


def _border():
    s = _side()
    return Border(left=s, right=s, top=s, bottom=s)


def _apply(cell, style):
    for attr, val in style.items():
        setattr(cell, attr, val)


# ─── Чтение порядка групп из файла ────────────────────────────────────────────

def load_group_order(config):
    order_file = BASE_DIR / config.get('group_order_file', 'Места_по_группам_товара_исправлено.xlsx')
    if not order_file.exists():
        return {}

    df = pd.read_excel(order_file, skiprows=1, header=None)
    order = {}
    for _, row in df.iterrows():
        group_name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        position = row.iloc[2] if len(row) > 2 and pd.notna(row.iloc[2]) else None
        if group_name and group_name != 'Группа товара' and position is not None:
            try:
                order[group_name] = int(float(position))
            except (ValueError, TypeError):
                pass
    return order


# ─── Чтение наценок из файла ──────────────────────────────────────────────────

def load_markups(config):
    markups_path = BASE_DIR / config.get('markups_file', 'наценки по группам товара.xlsx')
    if not markups_path.exists():
        return {}, {}

    df = pd.read_excel(markups_path, skiprows=2, header=None)

    group_markups = {}
    homeline_prices = {}
    in_homeline = False

    for _, row in df.iterrows():
        col_b = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        col_c = row.iloc[2] if len(row) > 2 else None

        if col_b == 'Артикул':
            in_homeline = True
            continue

        if not in_homeline:
            if col_b and col_b not in ('Группа товара', 'Наценка (%)', 'nan'):
                try:
                    markup_val = int(float(col_c)) if pd.notna(col_c) else 0
                    group_markups[col_b] = markup_val
                except (ValueError, TypeError):
                    pass
        else:
            article = col_b
            col_d = row.iloc[3] if len(row) > 3 else None
            if article and article != 'nan' and col_d is not None and pd.notna(col_d):
                try:
                    homeline_prices[article] = float(col_d)
                except (ValueError, TypeError):
                    pass

    return group_markups, homeline_prices


# ─── Извлечение бренда из названия товара ────────────────────────────────────

def extract_brand(name: str) -> str:
    """Возвращает бренд из названия товара.
    Бренд — первое слово из заглавных букв (≥4 символов), не являющееся
    русским прилагательным (окончания ОЙ, АЯ, ОЕ, ЫЙ, ИЙ).
    Пример: 'Аккумуляторная дрель EINHELL PXC...' → 'EINHELL'
    Если не найден — возвращает 'Noname'.
    """
    _adj_endings = re.compile(r'(ОЙ|АЯ|ОЕ|ЫЙ|ИЙ|УЮ|ЫЕ|ИЕ)$')
    for word in name.split():
        letters = re.sub(r'[^а-яёА-ЯЁa-zA-Z]', '', word)
        if len(letters) >= 4 and letters == letters.upper():
            if _adj_endings.search(letters):
                continue
            return word.rstrip(',').rstrip('.')
    return 'Noname'


# ─── Чтение прайса поставщика ─────────────────────────────────────────────────

def read_supplier_price(filepath, config):
    """Парсит прайс Аксенова: колонка A (0) — название, колонка N (13) — цена.
    Строки без цены используются как заголовки групп.
    Артикулы генерируются автоматически: UT-000001, UT-000002, ...
    """
    skiprows = config.get('read_options', {}).get('skiprows', 8)
    df_raw = pd.read_excel(filepath, header=None, skiprows=skiprows, usecols=[0, 13])
    df_raw.columns = ['name', 'price_in']

    rows = []
    current_group = ''
    counter = 1

    for _, row in df_raw.iterrows():
        name = str(row['name']).strip() if pd.notna(row['name']) else ''
        if not name or name == 'nan':
            continue

        price_num = pd.to_numeric(row['price_in'], errors='coerce')

        if pd.isna(price_num):
            current_group = name
        else:
            rows.append({
                'article':  f'UT-{counter:06d}',
                'group':    current_group,
                'brand':    extract_brand(name),
                'name':     name,
                'price_in': price_num,
            })
            counter += 1

    return pd.DataFrame(rows)


# ─── Применение наценок ───────────────────────────────────────────────────────

def apply_pricing(df, group_markups, homeline_prices, default_markup):
    df = df.copy()
    prices_out = []

    for _, row in df.iterrows():
        article = row['article']

        if article in homeline_prices:
            prices_out.append(homeline_prices[article])
            continue

        group = row['group']
        markup_pct = group_markups.get(group)
        if markup_pct is None:
            markup_pct = group_markups.get(group.rstrip())
        if markup_pct is None:
            markup_pct = default_markup

        price = row['price_in'] * (1 + markup_pct / 100)
        prices_out.append(round(price, 0))

    df['price_out'] = prices_out
    return df


# ─── Лист «Прайс клиента» ────────────────────────────────────────────────────

def build_pricelist(df, wb, config, group_order):
    ws = wb.active
    ws.title = 'Прайс клиента'
    company = config['company']

    for letter, w in COL_WIDTHS.items():
        ws.column_dimensions[letter].width = w

    dark_fill = PatternFill('solid', start_color=COLORS['header_bg'])

    # ─── Строка 1: Название компании ───
    ws.merge_cells('A1:F1')
    c = ws['A1']
    c.value = company['name']
    c.font = Font(bold=True, size=22, name='Arial', color=COLORS['header_font'])
    c.fill = dark_fill
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 38.1
    for col in range(2, 7):
        ws.cell(1, col).fill = dark_fill

    # ─── Строка 2: Город ───
    ws.merge_cells('A2:F2')
    c = ws['A2']
    c.value = company['city']
    c.font = Font(size=11, name='Arial', color=COLORS['city_font'])
    c.fill = dark_fill
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 21.95
    for col in range(2, 7):
        ws.cell(2, col).fill = dark_fill

    # ─── Строка 3: Телефоны ───
    ws.merge_cells('A3:F3')
    c = ws['A3']
    c.value = company['phones']
    c.font = Font(bold=True, size=10, name='Arial', color=COLORS['phone_font'])
    c.fill = dark_fill
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[3].height = 21.95
    for col in range(2, 7):
        ws.cell(3, col).fill = dark_fill

    # ─── Строка 4: Telegram с гиперссылками ───
    tg_links = company.get('telegram_links', [])

    ws.merge_cells('A4:C4')
    c = ws['A4']
    if tg_links:
        c.value = tg_links[0]['text']
        c.hyperlink = tg_links[0]['url']
    c.font = Font(bold=True, size=12, name='Arial', color=COLORS['tg_font'], underline='single')
    c.fill = dark_fill
    c.alignment = Alignment(horizontal='left', vertical='center')
    for col in range(2, 4):
        ws.cell(4, col).fill = dark_fill

    ws.merge_cells('D4:F4')
    c = ws['D4']
    if len(tg_links) > 1:
        c.value = tg_links[1]['text']
        c.hyperlink = tg_links[1]['url']
    c.font = Font(bold=True, size=12, name='Arial', color=COLORS['tg_font'], underline='single')
    c.fill = dark_fill
    c.alignment = Alignment(horizontal='right', vertical='center')
    for col in range(5, 7):
        ws.cell(4, col).fill = dark_fill
    ws.row_dimensions[4].height = 21.95

    # ─── Строка 5: Заголовки таблицы ───
    headers = ['№', 'Артикул', 'Бренд', 'Наименование', 'Цена (руб.)', 'Заказ (шт.)']
    header_style = {
        'font': Font(bold=True, size=10, name='Arial', color=COLORS['header_font']),
        'fill': PatternFill('solid', start_color=COLORS['header_bg']),
        'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'border': _border(),
    }
    for ci, h in enumerate(headers, 1):
        _apply(ws.cell(row=5, column=ci, value=h), header_style)

    ws.freeze_panes = 'A6'
    ws.auto_filter.ref = 'A5:F5'

    # ─── Сортировка групп: сначала приоритетные, потом по заданному порядку ───
    max_order = 9999
    unique_groups = df['group'].unique().tolist()
    priority_keys = [p.lower() for p in config.get('priority_groups', [])]

    def group_sort_key(g):
        g_lower = g.lower()
        for i, p in enumerate(priority_keys):
            if p in g_lower:
                return (0, i, g)
        pos = group_order.get(g) or group_order.get(g.rstrip())
        return (1, pos if pos is not None else max_order, g)

    sorted_groups = sorted(unique_groups, key=group_sort_key)

    # ─── Данные ───
    row = 6
    num = 1

    for group in sorted_groups:
        gdf = df[df['group'] == group]

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        gc = ws.cell(row=row, column=1, value=group)
        gc.font = Font(bold=True, size=10, name='Arial', color=COLORS['group_font'])
        gc.fill = PatternFill('solid', start_color=COLORS['group_bg'])
        gc.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        row += 1

        for _, item in gdf.iterrows():
            is_alt = (num % 2 == 0)
            alt_fill = PatternFill('solid', start_color=COLORS['alt_row']) if is_alt else None

            vals = [num, item['article'], item['brand'], item['name'], item['price_out'], '']

            for ci, v in enumerate(vals, 1):
                cell = ws.cell(row=row, column=ci, value=v)
                cell.font = Font(size=10, name='Arial')
                cell.border = _border()
                cell.alignment = Alignment(vertical='center', wrap_text=(ci == 4))
                if ci == 5:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.number_format = '#,##0'
                if alt_fill:
                    cell.fill = alt_fill

            row += 1
            num += 1


# ─── Главная функция ──────────────────────────────────────────────────────────

def transform(input_path, output_dir=None):
    config = load_config()

    print(f"  Читаю файл: {Path(input_path).name}")
    df = read_supplier_price(input_path, config)
    print(f"  Загружено позиций: {len(df)}")

    default_markup = config.get('default_markup', 13)
    print(f"  Наценка: +{default_markup}%")

    df = apply_pricing(df, {}, {}, default_markup)

    wb = openpyxl.Workbook()
    build_pricelist(df, wb, config, {})
    # Лист «Наценки» НЕ добавляется — только 1 лист «Прайс клиента»

    if output_dir is None:
        output_dir = str(BASE_DIR / 'output')
    os.makedirs(output_dir, exist_ok=True)

    prefix = config.get('output_prefix', 'БытТехОпт')
    ts = datetime.today().strftime('%Y%m%d')
    output_path = str(Path(output_dir) / f"{prefix}_{ts}.xlsx")
    wb.save(output_path)
    print(f"  Сохранено: {output_path}")
    return output_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Использование: python transform.py <прайс_поставщика.xlsx>")
        sys.exit(1)
    transform(sys.argv[1])
