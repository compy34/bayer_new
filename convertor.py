import pandas as pd
from datetime import datetime, timedelta
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from act_generator import generate_act_sheet


def format_date_ukrainian(date_str):
    """Format date as 'ДД ММММ РРРР року' in Ukrainian (e.g., '28 лютого 2026 року')."""
    month_names = [
        'січня', 'лютого', 'березня', 'квітня', 'травня', 'червня',
        'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня'
    ]
    dt = datetime.strptime(date_str, '%Y-%m-%d')
    return f'{dt.day:02d} {month_names[dt.month - 1]} {dt.year} року'


def get_last_day_of_month(date_str):
    """Get last day of month for a date in YYYY-MM-DD format, returns YYYY-MM-DD."""
    dt = datetime.strptime(date_str, '%Y-%m-%d')
    if dt.month == 12:
        next_month = datetime(dt.year + 1, 1, 1)
    else:
        next_month = datetime(dt.year, dt.month + 1, 1)
    last_day = next_month - timedelta(days=1)
    return last_day.strftime('%Y-%m-%d')

# ============================================================
# BRAND LISTS FROM TEMPLATE
# ============================================================

# Template brands for LIC (лікарські засоби) - from звіт викладка+MHL(ЛІК) row 10
LIK_BRANDS_TEMPLATE = [
    'Антифлу',
    'Бепантен',
    'Бепантен Плюс',
    'Елевіт-3',
    'Еріус',
    'Реліф',
    'Супрадин-10',
    'Ренні',
    'Терафлекс + Тералів',
]

# Template brands for DD (не лікарські засоби) - from звіт викладка+MHL(ДД) row 10
DD_BRANDS_TEMPLATE = [
    'Елевіт',
    'Супрадин-13',
    'Супрадин-10',
    'Бепантен Сенсідерм-4',
    'Бепантен Сенсідерм-2',
]

# Database brand mapping (from calc_Merch_KA BRAND column to template brand names)
DB_BRAND_TO_TEMPLATE = {
    # LIC brands
    'Антифлу': 'Антифлу',
    'Бепантен': 'Бепантен',
    'Бепантен Плюс': 'Бепантен Плюс',
    'Елевіт-3': 'Елевіт-3',
    'Еріус': 'Еріус',
    'Реліф': 'Реліф',
    'Супрадин-10': 'Супрадин-10',
    'Супрадин-10__': 'Супрадин-10',
    'Ренні': 'Ренні',
    'Терафлекс + Тералів': 'Терафлекс + Тералів',
    'Терафлекс': 'Терафлекс + Тералів',  # map to combined
    'Тералів': 'Терафлекс + Тералів',    # map to combined
    'Кальцемін': 'Кальцемін',
    'Бепантен Сенсідерм-4': 'Бепантен Сенсідерм-4',
    # DD brands
    'Елевіт': 'Елевіт',
    'Супрадин-13': 'Супрадин-13',
    'Супрадин-10': 'Супрадин-10',
    'Бепантен Сенсідерм-4': 'Бепантен Сенсідерм-4',
    'Бепантен Сенсідерм-2': 'Бепантен Сенсідерм-2',
    'Бепантен Плюс': 'Бепантен Плюс',
}

# Service prices (from LA_Price sheet in template)
SERVICE_PRICES = {
    'Бепантен': 800.0,
    'Антифлу': 300.0,
    'Еріус': 600.0,
    'Реліф': 800.0,
    'Ренні': 600.0,
    'Терафлекс + Тералів': 800.0,
    'Бепантен Плюс': 300.0,
    'Тералів': 300.0,
    'Кальцемін': 300.0,
    'Терафлекс': 300.0,
    'Елевіт-3': 600.0,
    'Супрадин-13': 800.0,
    'Елевіт': 800.0,
    'Супрадин-10': 800.0,
    'Бепантен Сенсідерм-4': 800.0,
    'Бепантен Сенсідерм-2': 800.0,
}

# ============================================================
# SKU CLASSIFICATION CONFIG (loaded from CSV)
# ============================================================

from pathlib import Path

SKU_CLASSIFICATION_FILE = Path(__file__).resolve().parent / 'sku_classification.csv'

_sku_classification_cache = None
_sku_to_category = None
_sku_to_brand = None


def load_sku_classification():
    """Load SKU classification from CSV file."""
    global _sku_classification_cache, _sku_to_category, _sku_to_brand

    if _sku_classification_cache is not None:
        return _sku_classification_cache

    import pandas as pd

    if not SKU_CLASSIFICATION_FILE.exists():
        # Return empty classification if file doesn't exist
        _sku_classification_cache = pd.DataFrame(columns=['SKU_UA', 'BRAND', 'CATEGORY', 'active'])
        _sku_to_category = {}
        _sku_to_brand = {}
        return _sku_classification_cache

    try:
        df = pd.read_csv(SKU_CLASSIFICATION_FILE, dtype=str)
        df.columns = df.columns.str.strip()

        # Ensure 'active' column exists (migration for existing CSV)
        if 'active' not in df.columns:
            df['active'] = 'true'

        # Normalize active column to boolean
        df['active'] = df['active'].fillna('true').astype(str).str.strip().str.lower().isin(['true', '1', 'yes', 'так'])

        # Normalize SKU names for matching
        df['SKU_NORMALIZED'] = df['SKU_UA'].fillna('').astype(str).str.strip().str.casefold()
        _sku_classification_cache = df

        # Build lookup dictionaries (only for active SKUs)
        active_df = df[df['active']]
        _sku_to_category = dict(zip(active_df['SKU_NORMALIZED'], active_df['CATEGORY']))
        _sku_to_brand = dict(zip(active_df['SKU_NORMALIZED'], active_df['BRAND']))

        return df
    except Exception as e:
        print(f"Error loading SKU classification: {e}")
        _sku_classification_cache = pd.DataFrame(columns=['SKU_UA', 'BRAND', 'CATEGORY', 'active'])
        _sku_to_category = {}
        _sku_to_brand = {}
        return _sku_classification_cache


def get_sku_category(sku_name):
    """Get category (ЛІК/ДД) for a SKU name."""
    if _sku_to_category is None:
        load_sku_classification()
    normalized = str(sku_name).strip().casefold()
    return _sku_to_category.get(normalized)


def get_sku_brand(sku_name):
    """Get template brand for a SKU name."""
    if _sku_to_brand is None:
        load_sku_classification()
    normalized = str(sku_name).strip().casefold()
    return _sku_to_brand.get(normalized)


def get_mhl_lik_sku():
    """Get list of LIC SKUs from classification (only active)."""
    df = load_sku_classification()
    if df.empty:
        return []
    return df[(df['CATEGORY'] == 'ЛІК') & (df['active'])]['SKU_UA'].tolist()


def get_mhl_dd_sku():
    """Get list of DD SKUs from classification (only active)."""
    df = load_sku_classification()
    if df.empty:
        return []
    return df[(df['CATEGORY'] == 'ДД') & (df['active'])]['SKU_UA'].tolist()


# For backward compatibility - these are now dynamically loaded
def get_mhl_lik_sku_list():
    """Backward compatible function returning LIC SKU list."""
    return get_mhl_lik_sku()


def get_mhl_dd_sku_list():
    """Backward compatible function returning DD SKU list."""
    return get_mhl_dd_sku()


# Legacy constants (populated on first access)
MHL_LIK_SKU = []
MHL_DD_SKU = []


def _init_legacy_sku_lists():
    """Initialize legacy SKU lists for backward compatibility."""
    global MHL_LIK_SKU, MHL_DD_SKU
    if not MHL_LIK_SKU:
        MHL_LIK_SKU = get_mhl_lik_sku()
    if not MHL_DD_SKU:
        MHL_DD_SKU = get_mhl_dd_sku()


def save_sku_classification(sku_name, brand, category, active=True):
    """Add or update a SKU classification in the CSV file."""
    import pandas as pd
    import csv

    global _sku_classification_cache, _sku_to_category, _sku_to_brand

    load_sku_classification()  # Ensure cache is loaded

    normalized = str(sku_name).strip().casefold()
    new_row = {
        'SKU_UA': str(sku_name).strip(),
        'BRAND': str(brand).strip(),
        'CATEGORY': str(category).strip(),
        'active': bool(active)
    }

    # Update cache
    if normalized in _sku_to_category:
        # Update existing
        idx = _sku_classification_cache[_sku_classification_cache['SKU_NORMALIZED'] == normalized].index
        if len(idx) > 0:
            _sku_classification_cache.loc[idx[0], 'SKU_UA'] = new_row['SKU_UA']
            _sku_classification_cache.loc[idx[0], 'BRAND'] = new_row['BRAND']
            _sku_classification_cache.loc[idx[0], 'CATEGORY'] = new_row['CATEGORY']
            _sku_classification_cache.loc[idx[0], 'active'] = new_row['active']
    else:
        # Add new
        new_row['SKU_NORMALIZED'] = normalized
        _sku_classification_cache = pd.concat([_sku_classification_cache, pd.DataFrame([new_row])], ignore_index=True)

    # Rebuild lookup dictionaries (only active SKUs)
    active_df = _sku_classification_cache[_sku_classification_cache['active']]
    _sku_to_category = dict(zip(active_df['SKU_NORMALIZED'], active_df['CATEGORY']))
    _sku_to_brand = dict(zip(active_df['SKU_NORMALIZED'], active_df['BRAND']))

    # Save to CSV with quoting for consistency
    try:
        # Save without the normalized column
        save_df = _sku_classification_cache[['SKU_UA', 'BRAND', 'CATEGORY', 'active']].copy()
        save_df.to_csv(SKU_CLASSIFICATION_FILE, index=False, encoding='utf-8', quoting=csv.QUOTE_ALL)
        return True
    except Exception as e:
        print(f"Error saving SKU classification: {e}")
        return False


def set_sku_active(sku_name, active=True):
    """Set the active status of a SKU."""
    import pandas as pd
    import csv

    global _sku_classification_cache, _sku_to_category, _sku_to_brand

    load_sku_classification()

    normalized = str(sku_name).strip().casefold()

    # Find and update in cache
    idx = _sku_classification_cache[_sku_classification_cache['SKU_NORMALIZED'] == normalized].index
    if len(idx) == 0:
        return False

    _sku_classification_cache.loc[idx[0], 'active'] = bool(active)

    # Rebuild lookup dictionaries (only active SKUs)
    active_df = _sku_classification_cache[_sku_classification_cache['active']]
    _sku_to_category = dict(zip(active_df['SKU_NORMALIZED'], active_df['CATEGORY']))
    _sku_to_brand = dict(zip(active_df['SKU_NORMALIZED'], active_df['BRAND']))

    # Save to CSV
    try:
        save_df = _sku_classification_cache[['SKU_UA', 'BRAND', 'CATEGORY', 'active']].copy()
        save_df.to_csv(SKU_CLASSIFICATION_FILE, index=False, encoding='utf-8', quoting=csv.QUOTE_ALL)
        return True
    except Exception as e:
        print(f"Error updating SKU active status: {e}")
        return False


def get_all_skus_with_status():
    """Get all SKUs with their classification and active status for management dialog."""
    df = load_sku_classification()
    if df.empty:
        return pd.DataFrame(columns=['SKU_UA', 'BRAND', 'CATEGORY', 'active'])
    return df[['SKU_UA', 'BRAND', 'CATEGORY', 'active']].copy()


def get_unknown_skus(mhl_df, sku_col='SKU_UA'):
    """Get unknown SKUs from MHL data that are not in classification (only checks active SKUs)."""
    load_sku_classification()

    if mhl_df is None or mhl_df.empty or sku_col not in mhl_df.columns:
        return []

    known_skus = set(_sku_to_category.keys())
    mhl_skus = mhl_df[sku_col].fillna('').astype(str).str.strip().str.casefold()
    unknown_mask = ~mhl_skus.isin(known_skus) & (mhl_skus != '')

    unknown_df = mhl_df[unknown_mask].drop_duplicates(subset=[sku_col])
    return unknown_df[sku_col].unique().tolist()


# Initialize legacy lists on module import
_init_legacy_sku_lists()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def map_db_brands_to_template(df, template_brands):
    """Map database BRAND values to template brand names."""
    if df is None or df.empty:
        return df
    work_df = df.copy()
    # Map BRAND column
    work_df['TEMPLATE_BRAND'] = work_df['BRAND'].map(DB_BRAND_TO_TEMPLATE)
    # Filter to only template brands
    work_df = work_df[work_df['TEMPLATE_BRAND'].isin(template_brands)]
    return work_df


def build_pivot_for_report(df, template_brands):
    """Build pivot table matching template structure: pharmacies x template brands."""
    if df is None or df.empty:
        return None

    work_df = map_db_brands_to_template(df, template_brands)
    if work_df.empty:
        return None

    # Pivot: pharmacy x template brand, sum of MERCH
    pivot_df = work_df.pivot_table(
        index='PHARMACY_NAME',
        columns='TEMPLATE_BRAND',
        values='MERCH',
        aggfunc='sum',
        fill_value=0,
    ).reset_index()

    # Ensure all template brands are present as columns
    for brand in template_brands:
        if brand not in pivot_df.columns:
            pivot_df[brand] = 0

    # Reorder columns: PHARMACY_NAME first, then template brands in order
    cols = ['PHARMACY_NAME'] + template_brands
    pivot_df = pivot_df[cols]

    # Add total sum column
    price_dict = {b: SERVICE_PRICES.get(b, 0) for b in template_brands}
    total_sum = 0
    for brand in template_brands:
        total_sum += pivot_df[brand] * price_dict.get(brand, 0)
    pivot_df['Загальна сума'] = total_sum.round(2)

    return pivot_df


def get_mhl_availability_sum(mhl_df, template_brands):
    """Calculate MHL availability service sum (Service 2) from MHL data.
    This corresponds to the 'наявність' service - sum of purchased goods where penetration >= target.
    """
    if mhl_df is None or mhl_df.empty:
        return 0.0

    work_df = mhl_df.copy()
    sku_col = 'SKU_UA' if 'SKU_UA' in work_df.columns else 'Назва товару'

    # Filter MHL data for relevant SKUs using dynamic classification
    is_lic = template_brands == LIK_BRANDS_TEMPLATE
    allowed_skus = get_mhl_lik_sku() if is_lic else get_mhl_dd_sku()
    allowed_values = {str(v).strip().casefold() for v in allowed_skus}
    sku_values = work_df[sku_col].fillna('').astype(str).str.strip().str.casefold()
    work_df = work_df[sku_values.isin(allowed_values)].copy()

    if work_df.empty:
        return 0.0

    # Convert numeric columns
    for column in ['PENETRATION', 'PRICE_CP', 'SUPPLY', 'PENETRATION_TARGET']:
        work_df[column] = pd.to_numeric(work_df[column], errors='coerce')

    # Calculate sum of purchased goods
    work_df['Сума закупленого'] = work_df['SUPPLY'] * work_df['PRICE_CP']

    # Eligible: PENETRATION >= PENETRATION_TARGET
    # If PENETRATION is NULL -> treat as 0 (not eligible unless target is also 0/NULL)
    # If PENETRATION_TARGET is NULL -> treat as 0 (eligible if penetration >= 0)
    penetration = work_df['PENETRATION'].fillna(0)
    target = work_df['PENETRATION_TARGET'].fillna(0)
    work_df['eligible'] = penetration >= target

    total = work_df.loc[work_df['eligible'], 'Сума закупленого'].sum()
    return round(total, 2)


# ============================================================
# ZVIT SHEET GENERATOR (звіт викладка+MHL) - matches template exactly
# ============================================================

def fill_zvit_worksheet(ws, pivot_df, template_brands, period_str, period_end_ukrainian, executor_name, contract_no, contract_date,
                         service1_total, service2_total, is_lic=True, mhl_df=None):
    """Fill звіт викладка+MHL sheet matching template structure exactly.
    Includes MHL data section at the bottom.
    """
    # Handle None or empty pivot_df - create sheet with headers only
    if pivot_df is None or pivot_df.empty:
        pivot_df = pd.DataFrame(columns=['PHARMACY_NAME'] + template_brands)

    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right = Alignment(horizontal='right', vertical='center', wrap_text=True)
    bold_font = Font(bold=True, size=10)
    normal_font = Font(size=10)
    header_font = Font(bold=True, size=9)
    title_font = Font(bold=True, size=12)

    num_brands = len(template_brands)
    # Columns: A=№, B=Pharmacy, C=Responsible, D... = brands, last = Total
    brand_start_col = 4
    brand_end_col = brand_start_col + num_brands - 1
    total_col = brand_end_col + 1
    end_col = total_col

    # --- Row 1: Title ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    act_type = "ЛІК" if is_lic else "ДД"
    ws.cell(1, 1, f'Додаток № 1 до Акта {act_type} № _______ про надані послуги')
    ws.cell(1, 1).font = title_font
    ws.cell(1, 1).alignment = center

    # --- Row 2: Contract reference ---
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
    ws.cell(2, 1, f'відповідно до Договору про маркетингові послуги № {contract_no} від {contract_date} (надалі – «Договір»)')
    ws.cell(2, 1).font = normal_font
    ws.cell(2, 1).alignment = center

    # --- Row 3: Report title ---
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=end_col)
    ws.cell(3, 1, 'Звіт щодо Послуг, передбачених Договором за період')
    ws.cell(3, 1).font = title_font
    ws.cell(3, 1).alignment = center

    # --- Row 4: Period ---
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=end_col)
    ws.cell(4, 1, period_str)
    ws.cell(4, 1).font = normal_font
    ws.cell(4, 1).alignment = center

    # --- Row 6: Location + date ---
    ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=end_col)
    ws.cell(6, 1, f'м. Київ                                                                                                                                                                                                                               {period_end_ukrainian}')
    ws.cell(6, 1).font = normal_font
    ws.cell(6, 1).alignment = left

    # --- Row 7: Confirmation text ---
    ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=end_col)
    confirmation = (
        f'{executor_name} підтверджує, що відповідно до Договору Замовнику були надані маркетингові послуги.\n\n'
        '1. Розміщення Товару Замовника на полицях аптечних закладів Аптечної мережі '
        'відповідно до встановленого Замовником плану викладки згідно з Планограмою та адресною програмою:'
    )
    ws.cell(7, 1, confirmation)
    ws.cell(7, 1).font = normal_font
    ws.cell(7, 1).alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[7].height = 60

    # --- Row 9: Table headers ---
    headers = ['№ п/п', 'Назва та адреса Аптеки/Аптечного пункту, де були надані маркетингові послуги',
               'Відповідальна особа', 'Бренди Товару, щодо яких надавалися маркетингові послуги, кількість викладок',
               'Сума до виплати, без ПДВ']

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(9, col_idx, header)
        cell.font = header_font
        cell.alignment = center
        cell.border = thin_border

    # --- Row 10: Brand sub-headers ---
    for i, brand in enumerate(template_brands):
        cell = ws.cell(10, brand_start_col + i, brand)
        cell.font = Font(bold=True, size=9)
        cell.alignment = center
        cell.border = thin_border

    # Total column header
    cell = ws.cell(10, total_col, 'Сума до виплати, без ПДВ')
    cell.font = header_font
    cell.alignment = center
    cell.border = thin_border

    ws.row_dimensions[9].height = 30
    ws.row_dimensions[10].height = 40

    # --- Data rows ---
    start_data_row = 11

    for idx, row_data in enumerate(pivot_df.iterrows(), start=start_data_row):
        row_values = row_data[1]
        pharmacy_name = row_values['PHARMACY_NAME']
        responsible = ''  # Could be extracted from data if available

        # № п/п
        cell = ws.cell(idx, 1, idx - start_data_row + 1)
        cell.font = normal_font
        cell.alignment = center
        cell.border = thin_border

        # Pharmacy name
        cell = ws.cell(idx, 2, pharmacy_name)
        cell.font = normal_font
        cell.alignment = left
        cell.border = thin_border

        # Responsible person
        cell = ws.cell(idx, 3, responsible)
        cell.font = normal_font
        cell.alignment = left
        cell.border = thin_border

        # Brand quantities
        row_total = 0
        for i, brand in enumerate(template_brands):
            qty = row_values.get(brand, 0)
            cell = ws.cell(idx, brand_start_col + i, int(qty) if qty else 0)
            cell.font = normal_font
            cell.alignment = center
            cell.border = thin_border
            cell.number_format = '#,##0'

            price = SERVICE_PRICES.get(brand, 0)
            row_total += qty * price

        # Total amount
        cell = ws.cell(idx, total_col, round(row_total, 2))
        cell.font = normal_font
        cell.alignment = right
        cell.border = thin_border
        cell.number_format = '#,##0.00'

    # --- Total row (Service 1) ---
    last_data_row = start_data_row + len(pivot_df) - 1
    total_row = last_data_row + 1

    ws.cell(total_row, 1, '').font = header_font
    ws.cell(total_row, 1).border = thin_border
    ws.cell(total_row, 2, 'ВСЬОГО:').font = header_font
    ws.cell(total_row, 2).alignment = left
    ws.cell(total_row, 2).border = thin_border
    ws.cell(total_row, 3, '').font = header_font
    ws.cell(total_row, 3).border = thin_border

    for i in range(num_brands):
        ws.cell(total_row, brand_start_col + i, '').font = header_font
        ws.cell(total_row, brand_start_col + i).border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='double')
        )

    cell = ws.cell(total_row, total_col, round(service1_total, 2))
    cell.font = header_font
    cell.alignment = right
    cell.number_format = '#,##0.00'
    cell.border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )

    # ============================================================
    # MHL SECTION (Service 2 - Availability) at the bottom
    # ============================================================
    mhl_start_row = total_row + 2

    # MHL section title
    ws.merge_cells(start_row=mhl_start_row, start_column=1, end_row=mhl_start_row, end_column=end_col)
    ws.cell(mhl_start_row, 1, '2. Маркетингові послуги з забезпечення наявності у аптечному закладі Аптечної мережі Товарів із Рекомендованого переліку Товарів Замовника відповідно до Договору.')
    ws.cell(mhl_start_row, 1).font = normal_font
    ws.cell(mhl_start_row, 1).alignment = left
    ws.row_dimensions[mhl_start_row].height = 40

    # MHL headers
    mhl_header_row = mhl_start_row + 1
    mhl_headers = ['№ п/п', 'Назва товару', 'Бренд',
                   'Відсоток аптек мережі, які мають товар',
                   'Розрахункова ціна за одиницю товару, грн',
                   'Кількість упаковок товару, закуплена аптечною мережею',
                   'Сума закупленого товару, грн']

    for col_idx, header in enumerate(mhl_headers, 1):
        cell = ws.cell(mhl_header_row, col_idx, header)
        cell.font = header_font
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[mhl_header_row].height = 30

    # Filter and prepare MHL data for this category
    mhl_data_rows = []
    if mhl_df is not None and not mhl_df.empty:
        work_df = mhl_df.copy()
        sku_col = 'SKU_UA' if 'SKU_UA' in work_df.columns else 'Назва товару'

        is_lic_category = template_brands == LIK_BRANDS_TEMPLATE
        allowed_skus = get_mhl_lik_sku() if is_lic_category else get_mhl_dd_sku()
        allowed_values = {str(v).strip().casefold() for v in allowed_skus}
        sku_values = work_df[sku_col].fillna('').astype(str).str.strip().str.casefold()
        work_df = work_df[sku_values.isin(allowed_values)].copy()

        if not work_df.empty:
            # Convert numeric columns, keeping track of which were originally NULL
            for column in ['PENETRATION', 'PRICE_CP', 'SUPPLY', 'PENETRATION_TARGET']:
                work_df[column] = pd.to_numeric(work_df[column], errors='coerce')

            # Calculate sum of purchased goods
            work_df['Сума закупленого'] = work_df['SUPPLY'] * work_df['PRICE_CP']

            # Eligible: PENETRATION >= PENETRATION_TARGET
            # If PENETRATION is NULL -> treat as 0 (not eligible unless target is also 0/NULL)
            # If PENETRATION_TARGET is NULL -> treat as 0 (eligible if penetration >= 0)
            penetration = work_df['PENETRATION'].fillna(0)
            target = work_df['PENETRATION_TARGET'].fillna(0)
            work_df['eligible'] = penetration >= target

            work_df = work_df[work_df['eligible']].copy()

            if not work_df.empty:
                # Group by product for display
                group_cols = ['DRUG_ID', 'SKU_UA', 'BRAND']
                agg_dict = {
                    'PENETRATION': 'last',
                    'PRICE_CP': 'last',
                    'SUPPLY': 'sum',
                    'PENETRATION_TARGET': 'last',
                    'Сума закупленого': 'sum',
                }
                if all(c in work_df.columns for c in group_cols):
                    work_df = work_df.groupby(group_cols, dropna=False).agg(agg_dict).reset_index()

                for _, row in work_df.iterrows():
                    mhl_data_rows.append({
                        'Назва товару': row.get('SKU_UA', ''),
                        'Бренд': row.get('BRAND', ''),
                        'PENETRATION': row.get('PENETRATION', 0),
                        'PRICE_CP': row.get('PRICE_CP', 0),
                        'SUPPLY': row.get('SUPPLY', 0),
                        'Сума закупленого': row.get('Сума закупленого', 0),
                    })

    # Write MHL data rows
    mhl_first_data_row = mhl_header_row + 1
    for idx, mhl_row in enumerate(mhl_data_rows, start=mhl_first_data_row):
        ws.cell(idx, 1, idx - mhl_first_data_row + 1).font = normal_font
        ws.cell(idx, 1).alignment = center
        ws.cell(idx, 1).border = thin_border

        ws.cell(idx, 2, mhl_row['Назва товару']).font = normal_font
        ws.cell(idx, 2).alignment = left
        ws.cell(idx, 2).border = thin_border

        ws.cell(idx, 3, mhl_row['Бренд']).font = normal_font
        ws.cell(idx, 3).alignment = left
        ws.cell(idx, 3).border = thin_border

        # PENETRATION as integer percentage with % sign
        penetration_pct = int(round(mhl_row['PENETRATION'] * 100))
        ws.cell(idx, 4, f'{penetration_pct}%').font = normal_font
        ws.cell(idx, 4).alignment = center
        ws.cell(idx, 4).border = thin_border

        ws.cell(idx, 5, round(mhl_row['PRICE_CP'], 2)).font = normal_font
        ws.cell(idx, 5).alignment = right
        ws.cell(idx, 5).border = thin_border
        ws.cell(idx, 5).number_format = '#,##0.00'

        ws.cell(idx, 6, int(mhl_row['SUPPLY'])).font = normal_font
        ws.cell(idx, 6).alignment = center
        ws.cell(idx, 6).border = thin_border

        ws.cell(idx, 7, round(mhl_row['Сума закупленого'], 2)).font = normal_font
        ws.cell(idx, 7).alignment = right
        ws.cell(idx, 7).border = thin_border
        ws.cell(idx, 7).number_format = '#,##0.00'

    mhl_last_data_row = max(mhl_first_data_row, mhl_first_data_row + len(mhl_data_rows) - 1)

    # MHL Total row (Total purchased)
    mhl_total_row = mhl_last_data_row + 1
    ws.cell(mhl_total_row, 1, '').font = header_font
    ws.cell(mhl_total_row, 1).border = thin_border
    ws.cell(mhl_total_row, 2, 'ВСЬОГО закуплено товару, грн.').font = header_font
    ws.cell(mhl_total_row, 2).alignment = right
    ws.cell(mhl_total_row, 2).border = thin_border

    for col in range(3, 7):
        ws.cell(mhl_total_row, col, '').font = header_font
        ws.cell(mhl_total_row, col).border = thin_border

    total_cell = ws.cell(mhl_total_row, 7, round(service2_total, 2))
    total_cell.font = header_font
    total_cell.alignment = right
    total_cell.number_format = '#,##0.00'
    total_cell.border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )

    # Reward row (10% of total purchased)
    mhl_reward_row = mhl_total_row + 1
    # Hardcoded date as requested
    reward_text = (
        f'Виногорода за забезпечення наявності у аптечному закладі '
        f'Товарів із Рекомендованого переліку Товару Замовника, '
        f'згідно пункту 2.2.2 Договору про маркетингові послуги №{contract_no} '
        f'від 27 лютого 2026 року'
    )
    ws.merge_cells(start_row=mhl_reward_row, start_column=1, end_row=mhl_reward_row, end_column=6)
    cell = ws.cell(mhl_reward_row, 1, reward_text)
    cell.font = header_font
    cell.alignment = left
    cell.border = thin_border
    # Apply border to all cells in merged range (only first cell gets value)
    for col in range(2, 7):
        ws.cell(mhl_reward_row, col).border = thin_border

    reward_amount = round(service2_total * 0.1, 2)
    reward_cell = ws.cell(mhl_reward_row, 7, reward_amount)
    reward_cell.font = header_font
    reward_cell.alignment = right
    reward_cell.number_format = '#,##0.00'
    reward_cell.border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )

    # VAT row (20% of reward)
    mhl_vat_row = mhl_reward_row + 1
    ws.cell(mhl_vat_row, 2, 'ПДВ 20%').font = header_font
    ws.cell(mhl_vat_row, 2).alignment = right
    ws.cell(mhl_vat_row, 2).border = thin_border
    vat_amount = round(reward_amount * 0.2, 2)
    vat_cell = ws.cell(mhl_vat_row, 7, vat_amount)
    vat_cell.font = header_font
    vat_cell.alignment = right
    vat_cell.number_format = '#,##0.00'
    vat_cell.border = thin_border

    # Total with VAT row (reward + VAT)
    mhl_total_vat_row = mhl_vat_row + 1
    ws.cell(mhl_total_vat_row, 2, 'Всього з ПДВ').font = header_font
    ws.cell(mhl_total_vat_row, 2).alignment = right
    ws.cell(mhl_total_vat_row, 2).border = thin_border
    total_vat_amount = round(reward_amount + vat_amount, 2)
    total_vat_cell = ws.cell(mhl_total_vat_row, 7, total_vat_amount)
    total_vat_cell.font = header_font
    total_vat_cell.alignment = right
    total_vat_cell.number_format = '#,##0.00'
    total_vat_cell.border = thin_border

    # --- Column widths ---
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 30
    for i in range(num_brands):
        ws.column_dimensions[get_column_letter(brand_start_col + i)].width = 12
    ws.column_dimensions[get_column_letter(total_col)].width = 18

    # MHL section column widths
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 22
    ws.column_dimensions['F'].width = 25
    ws.column_dimensions['G'].width = 22


# ============================================================
# EMPTY TMA SHEETS (під час акції)
# ============================================================

def create_empty_tma_sheet(ws, is_lic=True):
    """Create empty під час акції sheet with just the name."""
    act_type = "ЛІК" if is_lic else "ДД"
    ws.title = f'під час акції ({act_type})'
    # Sheet left empty as requested


# ============================================================
# MAIN GENERATE FUNCTION - exact template sheet names
# ============================================================

def generate_excel_report(
        df,
        df_req,
        output_filename='marketing_report.xlsx',
        period_start=None,
        period_end=None,
        selected_pharmacy='Усі аптеки',
        mhl_df=None,
):
    """Generate Excel report with 6 sheets matching template exactly:
    1. акт_ЛІК
    2. акт_ДД
    3. звіт викладка+MHL(ЛІК)
    4. звіт викладка+MHL(ДД)
    5. під час акції (ЛІК) - empty
    6. під час акції (ДД) - empty
    """
    if (df is None or df.empty) and (mhl_df is None or mhl_df.empty):
        return

    # Build pivot data for LIC and DD using template brand lists
    if df is not None and not df.empty:
        pivot_lik = build_pivot_for_report(df, LIK_BRANDS_TEMPLATE)
        pivot_dd = build_pivot_for_report(df, DD_BRANDS_TEMPLATE)
    else:
        pivot_lik = None
        pivot_dd = None

    # Calculate Service 1 totals (shelf placement / викладка)
    service1_lik = 0.0
    service1_dd = 0.0
    if pivot_lik is not None and not pivot_lik.empty:
        service1_lik = pivot_lik['Загальна сума'].sum()
    if pivot_dd is not None and not pivot_dd.empty:
        service1_dd = pivot_dd['Загальна сума'].sum()

    # Calculate Service 2 totals (availability / наявність) from MHL data
    service2_lik = get_mhl_availability_sum(mhl_df, LIK_BRANDS_TEMPLATE)
    service2_dd = get_mhl_availability_sum(mhl_df, DD_BRANDS_TEMPLATE)

    # Period formatting
    period_start = period_start or 'No date'
    period_end = period_end or 'No date'

    # Parse dates and format properly
    month_names = [
        'січня', 'лютого', 'березня', 'квітня', 'травня', 'червня',
        'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня'
    ]

    if period_start != 'No date' and period_end != 'No date':
        start_dt = datetime.strptime(period_start, '%Y-%m-%d')
        end_dt = datetime.strptime(period_end, '%Y-%m-%d')
        # Last day of month for period end
        if end_dt.month == 12:
            next_month = datetime(end_dt.year + 1, 1, 1)
        else:
            next_month = datetime(end_dt.year, end_dt.month + 1, 1)
        end_last_day = next_month - timedelta(days=1)

        # Format: "з "01" жовтня по "31" жовтня 2026 року."
        period_str = f'з «{start_dt.day:02d}» {month_names[start_dt.month - 1]} по «{end_last_day.day:02d}» {month_names[end_last_day.month - 1]} {end_last_day.year} року'
        period_end_ukrainian = f'{end_last_day.day:02d} {month_names[end_last_day.month - 1]} {end_last_day.year} року'
    else:
        period_str = period_start
        period_end_ukrainian = period_end

    # Contractor details
    contract_no = '260326'
    contract_date = '01 березня 2026 року'
    executor_name = 'ПрАТ «Аптеки Запоріжжя»'

    if df_req is not None and not df_req.empty:
        matched_row = None
        clean_selected = str(selected_pharmacy).strip()
        target_col = None
        for col_candidate in ['chain_name', 'Клиент из базы данных', 'Мережа', 'Назва']:
            if col_candidate in df_req.columns:
                target_col = col_candidate
                break

        if clean_selected and clean_selected != 'Усі аптеки' and target_col:
            filtered_req = df_req[
                df_req[target_col].astype(str).str.strip().str.lower()
                == clean_selected.lower()
            ]
            if not filtered_req.empty:
                matched_row = filtered_req.iloc[0]

        if matched_row is None:
            matched_row = df_req.iloc[0]

        executor_name = str(
            matched_row.get('Виконавець', matched_row.get('Юридическое лицо', executor_name))
        )
        contract_no = str(matched_row.get('№ договора', contract_no))
        contract_date = str(matched_row.get('Дата договора', contract_date))

    import openpyxl
    wb = openpyxl.Workbook()
    default_sheet = wb.active

    # ============================================================
    # SHEET 1: акт_ЛІК (uses act_generator for correct act structure)
    # ============================================================
    ws_akt_lik = wb.create_sheet(title='акт_ЛІК')
    # Pass total = service1 + service2 for both services
    generate_act_sheet(
        ws=ws_akt_lik,
        df_req=df_req,
        total_sum_without_vat=service1_lik + service2_lik,
        period_start=period_start,
        period_end=period_end,
        selected_pharmacy=selected_pharmacy,
        act_number='1',
        contract_no=contract_no,
        contract_date=contract_date,
        executor_name=executor_name
    )

    # ============================================================
    # SHEET 2: акт_ДД
    # ============================================================
    ws_akt_dd = wb.create_sheet(title='акт_ДД')
    generate_act_sheet(
        ws=ws_akt_dd,
        df_req=df_req,
        total_sum_without_vat=service1_dd + service2_dd,
        period_start=period_start,
        period_end=period_end,
        selected_pharmacy=selected_pharmacy,
        act_number='2',
        contract_no=contract_no,
        contract_date=contract_date,
        executor_name=executor_name
    )

    # ============================================================
    # SHEET 3: звіт викладка+MHL(ЛІК)
    # ============================================================
    ws_zvit_lik = wb.create_sheet(title='звіт викладка+MHL(ЛІК)')
    fill_zvit_worksheet(ws_zvit_lik, pivot_lik, LIK_BRANDS_TEMPLATE, period_str,
                        period_end_ukrainian,
                        executor_name, contract_no, contract_date,
                        service1_lik, service2_lik, is_lic=True, mhl_df=mhl_df)

    # ============================================================
    # SHEET 4: звіт викладка+MHL(ДД)
    # ============================================================
    ws_zvit_dd = wb.create_sheet(title='звіт викладка+MHL(ДД)')
    fill_zvit_worksheet(ws_zvit_dd, pivot_dd, DD_BRANDS_TEMPLATE, period_str,
                        period_end_ukrainian,
                        executor_name, contract_no, contract_date,
                        service1_dd, service2_dd, is_lic=False, mhl_df=mhl_df)

    # ============================================================
    # SHEET 5: під час акції (ЛІК) - empty
    # ============================================================
    ws_tma_lik = wb.create_sheet(title='під час акції (ЛІК)')
    create_empty_tma_sheet(ws_tma_lik, is_lic=True)

    # ============================================================
    # SHEET 6: під час акції (ДД) - empty
    # ============================================================
    ws_tma_dd = wb.create_sheet(title='під час акції (ДД)')
    create_empty_tma_sheet(ws_tma_dd, is_lic=False)

    # Remove default sheet
    if default_sheet.title in wb.sheetnames and len(wb.sheetnames) > 1:
        if default_sheet.title == 'Sheet':
            wb.remove(default_sheet)

    wb.save(output_filename)
    return output_filename


# ============================================================
# DIAGNOSTIC FUNCTIONS
# ============================================================

def build_report_dataframe(df):
    """Build combined preview dataframe for GUI tree view."""
    if df is None or df.empty:
        return pd.DataFrame()

    pivot_lik = build_pivot_for_report(df, LIK_BRANDS_TEMPLATE)
    pivot_dd = build_pivot_for_report(df, DD_BRANDS_TEMPLATE)

    result_rows = []

    for pivot_df, brand_type in [(pivot_lik, 'ЛІК'), (pivot_dd, 'ДД')]:
        if pivot_df is not None and not pivot_df.empty:
            template_brands = LIK_BRANDS_TEMPLATE if brand_type == 'ЛІК' else DD_BRANDS_TEMPLATE
            for _, row in pivot_df.iterrows():
                pharmacy = row['PHARMACY_NAME']
                for brand in template_brands:
                    qty = row.get(brand, 0)
                    if qty > 0:
                        price = SERVICE_PRICES.get(brand, 0)
                        amount = round(qty * price, 2)
                        result_rows.append({
                            'Тип': brand_type,
                            'Аптека': pharmacy,
                            'Бренд': brand,
                            'Кількість': int(qty),
                            'Ціна': price,
                            'Сума': amount,
                        })

    return pd.DataFrame(result_rows)


def mhl_brand_counts(df):
    """Return category counts for diagnostics before Excel generation."""
    if df is None or df.empty or 'SKU_UA' not in df.columns:
        return 0, 0, 0

    load_sku_classification()
    sku_names = df['SKU_UA'].fillna('').astype(str).str.strip().str.casefold()

    lik_skus = {str(v).strip().casefold() for v in get_mhl_lik_sku()}
    dd_skus = {str(v).strip().casefold() for v in get_mhl_dd_sku()}

    lik_match = sku_names.isin(lik_skus)
    dd_match = sku_names.isin(dd_skus)
    return int(lik_match.sum()), int(dd_match.sum()), int(
        (~(lik_match | dd_match)).sum()
    )


def mhl_unknown_samples(df, limit=8):
    """Return real database values that are not classified by the lists."""
    if df is None or df.empty:
        return []

    load_sku_classification()
    sku_names = df['SKU_UA'].fillna('').astype(str).str.strip()

    known = set(_sku_to_category.keys())
    unknown = df.loc[
        ~sku_names.str.casefold().isin(known),
        ['BRAND', 'SKU_UA'],
    ].drop_duplicates()
    return [
        f'{brand} / {sku}'
        for brand, sku in unknown.head(limit).itertuples(index=False, name=None)
    ]