import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import threading
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from act_generator import generate_act_sheet, generate_tma_act_sheet
from config import (
    SERVICE_PRICES,
    DB_BRAND_TO_TEMPLATE,
    LIK_BRANDS_TEMPLATE,
    DD_BRANDS_TEMPLATE,
    DEFAULT_CONTRACT_NO,
    DEFAULT_CONTRACT_DATE,
    DEFAULT_TMA_CONTRACT_DATE,
    DEFAULT_EXECUTOR_NAME,
    MHL_REWARD_MULTIPLIER,
    TMA_ACTION_PERCENT_FILE,
    MHL_ACTION_PERCENT_FILE,
    LA_PRICE_FILE,
)
from chain_normalizer import normalize_chain_name, find_matching_chain


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
# SIMPLE SKU CLASSIFICATION (brand-based ЛІК/ДД)
# ============================================================
# No CSV file needed - classification is derived from brand mapping

from config import DB_BRAND_TO_TEMPLATE, LIK_BRANDS_TEMPLATE, DD_BRANDS_TEMPLATE


def get_sku_category(sku_name):
    """Get category (ЛІК/ДД) for a SKU name based on its brand.

    Uses DB_BRAND_TO_TEMPLATE mapping to find template brand,
    then checks if template brand is in LIK or DD list.
    """
    # This function is kept for compatibility but doesn't use SKU name directly
    # Category is determined by brand, not SKU
    return None  # Not used - use get_brand_category instead


def get_sku_brand(sku_name):
    """Get template brand for a SKU name - not used in new system."""
    return None  # Not used


def get_brand_category(brand):
    """Get category (ЛІК/ДД) for a brand name.

    Maps brand through DB_BRAND_TO_TEMPLATE, then checks
    if template brand is in LIK or DD template lists.
    """
    if not brand:
        return None

    brand_clean = str(brand).strip()

    # Try direct mapping
    template_brand = DB_BRAND_TO_TEMPLATE.get(brand_clean)
    if not template_brand:
        # Try case-insensitive lookup
        for k, v in DB_BRAND_TO_TEMPLATE.items():
            if k.lower() == brand_clean.lower():
                template_brand = v
                break

    if template_brand:
        if template_brand in LIK_BRANDS_TEMPLATE:
            return 'ЛІК'
        elif template_brand in DD_BRANDS_TEMPLATE:
            return 'ДД'

    # If brand is already a template brand (direct match)
    if brand_clean in LIK_BRANDS_TEMPLATE:
        return 'ЛІК'
    if brand_clean in DD_BRANDS_TEMPLATE:
        return 'ДД'

    return None


# Backward compatibility - these now return brand lists instead of SKU lists
def get_mhl_lik_sku():
    """Get list of LIC template brands (for backward compatibility)."""
    return list(LIK_BRANDS_TEMPLATE)


def get_mhl_dd_sku():
    """Get list of DD template brands (for backward compatibility)."""
    return list(DD_BRANDS_TEMPLATE)


def get_mhl_lik_sku_list():
    """Backward compatible function returning LIC brand list."""
    return get_mhl_lik_sku()


def get_mhl_dd_sku_list():
    """Backward compatible function returning DD brand list."""
    return get_mhl_dd_sku()


def get_db_brands_for_category(template_brands):
    """
    Get database brand names for a category (ЛІК or ДД).

    Args:
        template_brands: Either LIK_BRANDS_TEMPLATE or DD_BRANDS_TEMPLATE

    Returns:
        Set of database brand names that map to the given template brands
    """
    # Build reverse mapping: template brand -> set of db brands
    template_to_db = {}
    for db_brand, template_brand in DB_BRAND_TO_TEMPLATE.items():
        if template_brand not in template_to_db:
            template_to_db[template_brand] = set()
        template_to_db[template_brand].add(db_brand)

    db_brands = set()
    for template_brand in template_brands:
        db_brands.update(template_to_db.get(template_brand, set()))
    return db_brands


def get_mhl_lik_brands():
    """Get database brand names for ЛІК category."""
    return get_db_brands_for_category(LIK_BRANDS_TEMPLATE)


def get_mhl_dd_brands():
    """Get database brand names for ДД category."""
    return get_db_brands_for_category(DD_BRANDS_TEMPLATE)


# ============================================================
# PENETRATION TARGET CONFIG (local overrides)
# ============================================================
# Allows overriding database PENETRATION_TARGET values per SKU_UA

PENETRATION_TARGET_FILE = Path(__file__).resolve().parent / 'penetration_targets.json'
_penetration_target_cache = None
_penetration_target_lock = threading.Lock()


def load_penetration_targets():
    """Load penetration targets from JSON file."""
    global _penetration_target_cache

    if _penetration_target_cache is not None:
        return _penetration_target_cache

    import json

    if not PENETRATION_TARGET_FILE.exists():
        _penetration_target_cache = {}
        return _penetration_target_cache

    try:
        with open(PENETRATION_TARGET_FILE, 'r', encoding='utf-8') as f:
            _penetration_target_cache = json.load(f)
        # Normalize keys to lowercase for case-insensitive matching
        _penetration_target_cache = {
            str(k).strip().casefold(): float(v) for k, v in _penetration_target_cache.items()
        }
        return _penetration_target_cache
    except Exception as e:
        print(f"Error loading penetration targets: {e}")
        _penetration_target_cache = {}
        return _penetration_target_cache


def _save_penetration_targets(targets):
    """Thread-safe save of penetration targets to JSON."""
    with _penetration_target_lock:
        import json
        try:
            with open(PENETRATION_TARGET_FILE, 'w', encoding='utf-8') as f:
                json.dump(targets, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"Error saving penetration targets: {e}")
            return False


def get_penetration_target(sku_ua):
    """Get local penetration target override for a SKU.
    Returns None if no override exists (use database value).
    """
    load_penetration_targets()
    if _penetration_target_cache is None:
        return None
    normalized = str(sku_ua).strip().casefold()
    return _penetration_target_cache.get(normalized)


def set_penetration_target(sku_ua, target):
    """Set local penetration target override for a SKU.
    target: float value (e.g., 0.8 for 80%) or None to remove override
    """
    load_penetration_targets()
    normalized = str(sku_ua).strip().casefold()

    if target is None:
        # Remove override
        if normalized in _penetration_target_cache:
            del _penetration_target_cache[normalized]
    else:
        # Set override
        _penetration_target_cache[normalized] = float(target)

    return _save_penetration_targets(_penetration_target_cache)


def get_all_penetration_targets():
    """Get all local penetration target overrides for management dialog."""
    load_penetration_targets()
    if not _penetration_target_cache:
        return pd.DataFrame(columns=['SKU_UA', 'TARGET'])
    df = pd.DataFrame([
        {'SKU_UA': k, 'TARGET': v} for k, v in _penetration_target_cache.items()
    ])
    return df.sort_values('SKU_UA')


# ============================================================
# MHL ACTION PERCENTAGE LOADER (for "під час акції" sheets)
# ============================================================

MHL_ACTION_PERCENT_FILE = Path(__file__).resolve().parent / MHL_ACTION_PERCENT_FILE
_mhl_action_percent_cache = None
_mhl_action_percent_lock = threading.Lock()


def load_mhl_action_percents(filepath=None):
    """Load action percentages from Excel file.

    Expected file: MHL_Q3'2026.xlsx (configurable)
    Sheet: "для договору"
    Columns: №, SKU_ENG, SKU_UKR, Q3'26, %

    Returns dict mapping SKU_UKR (normalized) -> percentage (float, 0-100)
    """
    global _mhl_action_percent_cache

    if _mhl_action_percent_cache is not None:
        return _mhl_action_percent_cache

    file_to_load = Path(filepath) if filepath else MHL_ACTION_PERCENT_FILE

    if not file_to_load.exists():
        _mhl_action_percent_cache = {}
        return _mhl_action_percent_cache

    try:
        df = pd.read_excel(file_to_load, sheet_name='для договору', dtype=str)
        # Normalize column names
        df.columns = df.columns.astype(str).str.strip()

        # Find the percentage column (Q3'26, % or similar)
        pct_col = None
        for col in df.columns:
            col_lower = col.strip().lower()
            if 'q3' in col_lower and ('%' in col_lower or 'percent' in col_lower or 'відсоток' in col_lower):
                pct_col = col
                break

        # Find SKU_UKR column
        sku_col = None
        for col in df.columns:
            col_lower = col.strip().lower()
            if 'sku_ukr' in col_lower or 'sku_укр' in col_lower or 'sku укр' in col_lower:
                sku_col = col
                break

        if not pct_col or not sku_col:
            print(f"Warning: Required columns not found in MHL action percent file. Columns: {list(df.columns)}")
            _mhl_action_percent_cache = {}
            return _mhl_action_percent_cache

        # Build mapping: SKU_UKR (normalized) -> percentage
        _mhl_action_percent_cache = {}
        for _, row in df.iterrows():
            sku = str(row.get(sku_col, '')).strip()
            pct_str = str(row.get(pct_col, '')).strip()

            if sku and pct_str:
                try:
                    # Handle percentage strings like "10%", "10.5%", "0.1", "10"
                    pct_clean = pct_str.replace('%', '').replace(',', '.').strip()
                    pct_val = float(pct_clean)
                    # If value > 1, assume it's already in percent (e.g., 10 = 10%)
                    # If value <= 1, assume it's a fraction (e.g., 0.1 = 10%)
                    if pct_val > 1:
                        pct_val = pct_val  # already percent
                    else:
                        pct_val = pct_val * 100  # convert fraction to percent

                    _mhl_action_percent_cache[sku.casefold()] = pct_val
                except ValueError:
                    continue

        return _mhl_action_percent_cache

    except Exception as e:
        print(f"Error loading MHL action percentages: {e}")
        _mhl_action_percent_cache = {}
        return _mhl_action_percent_cache


def get_mhl_action_percent(sku_ukr):
    """Get action percentage for a SKU.
    Returns percentage as float (e.g., 10.0 for 10%), or 0.0 if not found.
    """
    load_mhl_action_percents()
    if _mhl_action_percent_cache is None:
        return 0.0
    normalized = str(sku_ukr).strip().casefold()
    return _mhl_action_percent_cache.get(normalized, 0.0)


def set_mhl_action_percent_file(filepath):
    """Update the MHL action percent file path and reload."""
    global MHL_ACTION_PERCENT_FILE, _mhl_action_percent_cache
    MHL_ACTION_PERCENT_FILE = Path(filepath)
    _mhl_action_percent_cache = None  # Force reload
    return load_mhl_action_percents()


# ============================================================
# TMA ACTION PERCENTAGE LOADER (for "під час акції" sheets - TMA)
# ============================================================

TMA_ACTION_PERCENT_FILE = Path(__file__).resolve().parent / TMA_ACTION_PERCENT_FILE
_tma_action_percent_cache = None
_tma_action_percent_lock = threading.Lock()


def load_tma_action_percents(filepath=None):
    """Load TMA action percentages from Excel file.

    Expected file: TMA.xlsx (configurable)
    Sheet: first sheet
    Columns: Назва мережі, Назва СКЮ, Початок промо, Кінець промо, знижка

    Returns list of dicts with keys: chain, sku, promo_start, promo_end, discount_pct
    If file doesn't exist, returns empty list (will result in 0% for all SKUs).
    """
    global _tma_action_percent_cache

    with _tma_action_percent_lock:
        if _tma_action_percent_cache is not None:
            return _tma_action_percent_cache

        file_to_load = Path(filepath) if filepath else TMA_ACTION_PERCENT_FILE

        if not file_to_load.exists():
            _tma_action_percent_cache = []
            return _tma_action_percent_cache

        try:
            df = pd.read_excel(file_to_load, sheet_name=0, dtype=str)  # first sheet
            # Normalize column names
            df.columns = df.columns.astype(str).str.strip()

            # Map expected columns (Ukrainian names)
            col_map = {}
            for col in df.columns:
                col_lower = col.strip().lower()
                if 'назва мереж' in col_lower or 'chain' in col_lower:
                    col_map['chain'] = col
                elif 'назва ску' in col_lower or 'sku' in col_lower:
                    col_map['sku'] = col
                elif 'початок промо' in col_lower or 'promo start' in col_lower or 'start' in col_lower:
                    col_map['promo_start'] = col
                elif 'кінець промо' in col_lower or 'promo end' in col_lower or 'end' in col_lower:
                    col_map['promo_end'] = col
                elif 'знижк' in col_lower or 'discount' in col_lower or '%' in col_lower:
                    col_map['discount'] = col

            required = ['chain', 'sku', 'promo_start', 'promo_end', 'discount']
            if not all(k in col_map for k in required):
                print(f"Warning: Required columns not found in TMA file. Found: {col_map}, Expected: {required}")
                _tma_action_percent_cache = []
                return _tma_action_percent_cache

            # Parse rows
            records = []
            for _, row in df.iterrows():
                chain = str(row.get(col_map['chain'], '')).strip()
                sku = str(row.get(col_map['sku'], '')).strip()
                promo_start = str(row.get(col_map['promo_start'], '')).strip()
                promo_end = str(row.get(col_map['promo_end'], '')).strip()
                discount_str = str(row.get(col_map['discount'], '')).strip()

                if not (chain and sku and discount_str):
                    continue

                try:
                    pct_clean = discount_str.replace('%', '').replace(',', '.').strip()
                    discount_pct = float(pct_clean)
                    if discount_pct <= 1:
                        discount_pct = discount_pct * 100

                    # Parse dates (try multiple formats)
                    start_date = _parse_date_flexible(promo_start)
                    end_date = _parse_date_flexible(promo_end)

                    records.append({
                        'chain': chain.casefold(),
                        'sku': sku.casefold(),
                        'promo_start': start_date,
                        'promo_end': end_date,
                        'discount_pct': discount_pct
                    })
                except ValueError:
                    continue

            _tma_action_percent_cache = records
            return _tma_action_percent_cache

        except Exception as e:
            print(f"Error loading TMA action percentages: {e}")
            _tma_action_percent_cache = []
            return _tma_action_percent_cache


def _parse_date_flexible(date_str):
    """Parse date from various formats: DD.MM.YYYY, YYYY-MM-DD, DD/MM/YYYY, etc."""
    if not date_str:
        return None
    for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%d.%m.%y', '%Y/%m/%d'):
        try:
            return datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            continue
    return None


def get_tma_action_percent(sku_ukr, period_start=None, period_end=None, chain_name=None):
    """Get TMA action percentage for a SKU, filtered by period and chain.

    Args:
        sku_ukr: SKU name
        period_start: Report period start (date or 'YYYY-MM-DD' string)
        period_end: Report period end (date or 'YYYY-MM-DD' string)
        chain_name: Pharmacy chain name (will be normalized)

    Returns:
        Percentage as float (e.g., 10.0 for 10%), or 0.0 if not found / file missing.
    """
    load_tma_action_percents()
    if not _tma_action_percent_cache:
        return 0.0

    normalized_sku = str(sku_ukr).strip().casefold()
    normalized_chain = chain_normalizer.normalize_chain_name(chain_name).casefold() if chain_name else None

    # Parse period dates
    if isinstance(period_start, str):
        period_start = _parse_date_flexible(period_start)
    if isinstance(period_end, str):
        period_end = _parse_date_flexible(period_end)

    # Find matching record: same SKU, chain matches (or no chain filter), promo overlaps with period
    for rec in _tma_action_percent_cache:
        if rec['sku'] != normalized_sku:
            continue
        if normalized_chain and rec['chain'] != normalized_chain:
            continue
        # Check promo period overlaps with report period
        if period_start and rec['promo_end'] and rec['promo_end'] < period_start:
            continue
        if period_end and rec['promo_start'] and rec['promo_start'] > period_end:
            continue
        return rec['discount_pct']

    return 0.0


def set_tma_action_percent_file(filepath):
    """Update the TMA action percent file path and reload."""
    global TMA_ACTION_PERCENT_FILE, _tma_action_percent_cache
    TMA_ACTION_PERCENT_FILE = Path(filepath)
    _tma_action_percent_cache = None  # Force reload
    return load_tma_action_percents()


# ============================================================
# LA_PRICE LOADER (service prices per brand)
# ============================================================

LA_PRICE_FILE = Path(__file__).resolve().parent / LA_PRICE_FILE
_la_price_cache = None
_la_price_lock = threading.Lock()


def load_la_price(filepath=None):
    """Load service prices from Excel file (LA_Price).

    Expected file: LA_Price.xlsx (configurable)
    Sheet: "LA_Price" or first sheet
    Columns: Brand, Price (or Brand, Ціна)

    Returns dict mapping Brand -> price (float)
    """
    global _la_price_cache

    if _la_price_cache is not None:
        return _la_price_cache

    file_to_load = Path(filepath) if filepath else LA_PRICE_FILE

    if not file_to_load.exists():
        # Return default prices from config
        _la_price_cache = SERVICE_PRICES.copy()
        return _la_price_cache

    try:
        # Try to read "LA_Price" sheet first, then first sheet
        try:
            df = pd.read_excel(file_to_load, sheet_name='LA_Price', dtype=str)
        except ValueError:
            df = pd.read_excel(file_to_load, sheet_name=0, dtype=str)

        # Normalize column names
        df.columns = df.columns.astype(str).str.strip()

        # Find Brand column
        brand_col = None
        for col in df.columns:
            col_lower = col.strip().lower()
            if col_lower in ['brand', 'бренд', 'назва', 'name']:
                brand_col = col
                break

        # Find Price column
        price_col = None
        for col in df.columns:
            col_lower = col.strip().lower()
            if col_lower in ['price', 'ціна', 'price_uah', 'cost', 'вартість']:
                price_col = col
                break

        if not brand_col or not price_col:
            print(f"Warning: Required columns not found in LA_Price file. Columns: {list(df.columns)}")
            _la_price_cache = SERVICE_PRICES.copy()
            return _la_price_cache

        # Build mapping: Brand -> price
        _la_price_cache = {}
        for _, row in df.iterrows():
            brand = str(row.get(brand_col, '')).strip()
            price_str = str(row.get(price_col, '')).strip()

            if brand and price_str:
                try:
                    price_clean = price_str.replace(',', '.').replace(' ', '').strip()
                    price_val = float(price_clean)
                    _la_price_cache[brand] = price_val
                except ValueError:
                    continue

        return _la_price_cache

    except Exception as e:
        print(f"Error loading LA_Price: {e}")
        _la_price_cache = SERVICE_PRICES.copy()
        return _la_price_cache


def get_la_price(brand):
    """Get service price for a brand.
    Returns price as float, or default from SERVICE_PRICES if not found.
    """
    load_la_price()
    if _la_price_cache is None:
        return SERVICE_PRICES.get(brand, 0.0)
    return _la_price_cache.get(brand, SERVICE_PRICES.get(brand, 0.0))


def get_all_la_prices():
    """Get all loaded service prices."""
    load_la_price()
    if _la_price_cache is None:
        return SERVICE_PRICES.copy()
    return _la_price_cache.copy()


def set_la_price_file(filepath):
    """Update the LA_Price file path and reload."""
    global LA_PRICE_FILE, _la_price_cache
    LA_PRICE_FILE = Path(filepath)
    _la_price_cache = None  # Force reload
    return load_la_price()


def set_la_price(brand, price):
    """Set/update a service price in memory (not persisted to file).
    For persistence, edit the Excel file directly and reload.
    """
    load_la_price()
    _la_price_cache[brand] = float(price)


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
    price_dict = {b: get_la_price(b) for b in template_brands}
    total_sum = 0
    for brand in template_brands:
        total_sum += pivot_df[brand] * price_dict.get(brand, 0)
    pivot_df['Загальна сума'] = total_sum.round(2)

    return pivot_df


def get_mhl_availability_sum(mhl_df, template_brands):
    """Calculate MHL availability service sum (Service 2) from MHL data.
    This corresponds to the 'наявність' service - sum of purchased goods where penetration >= target.
    Local penetration target overrides are applied from penetration_targets.json.
    """
    if mhl_df is None or mhl_df.empty:
        return 0.0

    work_df = mhl_df.copy()

    # Filter MHL data by brand using dynamic classification
    is_lic = template_brands == LIK_BRANDS_TEMPLATE
    allowed_brands = get_mhl_lik_brands() if is_lic else get_mhl_dd_brands()
    # BRAND column in MHL data contains database brand names
    brand_col = 'BRAND' if 'BRAND' in work_df.columns else 'brand'
    # Case-insensitive matching with whitespace handling
    brand_values = work_df[brand_col].fillna('').astype(str).str.strip()
    allowed_brands_lower = {b.lower() for b in allowed_brands}
    work_df = work_df[brand_values.str.lower().isin(allowed_brands_lower)].copy()

    if work_df.empty:
        return 0.0

    # Convert numeric columns
    for column in ['PENETRATION', 'PRICE_CP', 'SUPPLY', 'PENETRATION_TARGET']:
        work_df[column] = pd.to_numeric(work_df[column], errors='coerce')

    # Apply local penetration target overrides
    local_targets = load_penetration_targets()
    if local_targets:
        # Create a mapping from normalized SKU_UA to local target
        work_df['sku_normalized'] = work_df[sku_col].fillna('').astype(str).str.strip().str.casefold()
        work_df['local_target'] = work_df['sku_normalized'].map(local_targets)
        # Use local target where available, otherwise fall back to database PENETRATION_TARGET
        work_df['effective_target'] = work_df['local_target'].combine_first(work_df['PENETRATION_TARGET'])
    else:
        work_df['effective_target'] = work_df['PENETRATION_TARGET']

    # Calculate sum of purchased goods
    work_df['Сума закупленого'] = work_df['SUPPLY'] * work_df['PRICE_CP']

    # Eligible: PENETRATION >= effective_target
    # If PENETRATION is NULL -> treat as 0
    # If effective_target is NULL -> treat as 0
    penetration = work_df['PENETRATION'].fillna(0)
    target = work_df['effective_target'].fillna(0)
    work_df['eligible'] = penetration >= target

    total = work_df.loc[work_df['eligible'], 'Сума закупленого'].sum()
    return round(total, 2)


# ============================================================
# ZVIT SHEET GENERATOR (звіт викладка+MHL) - matches template exactly
# ============================================================

def fill_zvit_worksheet(ws, pivot_df, template_brands, period_str, period_end_ukrainian, executor_name, contract_no, contract_date,
                         service1_total, service2_raw, service2_reward, is_lic=True, mhl_df=None, df_req=None, selected_pharmacy=None):
    """Fill звіт викладка+MHL sheet matching template structure exactly.
    Includes MHL data section at the bottom.

    service2_raw: total purchased goods where penetration >= target (for "Всього без ПДВ")
    service2_reward: reward amount = service2_raw * MHL_REWARD_MULTIPLIER (for "Винагорода")
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
    # All text at 12pt; header (шапка) in bold
    header_font = Font(bold=True, size=12)       # For header rows (шапка) - bold 12pt
    normal_font = Font(size=12)                   # For data rows - regular 12pt
    title_font = Font(bold=True, size=12)         # For titles - bold 12pt

    num_brands = len(template_brands)
    # Columns: A=№, B=Pharmacy, C=first brand, ... = brands, last = Total
    # "Відповідальна особа" column completely removed - brands start at column 3
    brand_start_col = 3
    brand_end_col = brand_start_col + num_brands - 1
    total_col = brand_end_col + 1
    end_col = total_col

    # --- Row 1: Title ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    act_type = "ЛІК" if is_lic else "ДД"
    ws.cell(1, 1, f'Додаток № 1 до Акта {act_type} № _______ про надані послуги')
    ws.cell(1, 1).font = header_font
    ws.cell(1, 1).alignment = center

    # --- Row 2: Contract reference ---
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
    ws.cell(2, 1, f'відповідно до Договору про маркетингові послуги № {contract_no} від {contract_date} (надалі – «Договір»)')
    ws.cell(2, 1).font = header_font
    ws.cell(2, 1).alignment = center

    # --- Row 3: Report title ---
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=end_col)
    ws.cell(3, 1, 'Звіт щодо Послуг, передбачених Договором за період')
    ws.cell(3, 1).font = header_font
    ws.cell(3, 1).alignment = center

    # --- Row 4: Period ---
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=end_col)
    ws.cell(4, 1, period_str)
    ws.cell(4, 1).font = header_font
    ws.cell(4, 1).alignment = center

    # --- Row 6: Location + date ---
    ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=end_col)
    ws.cell(6, 1, f'м. Київ                                                                                                                                                                                                                               {period_end_ukrainian}')
    ws.cell(6, 1).font = header_font
    ws.cell(6, 1).alignment = left

    # --- Row 7: Confirmation text ---
    ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=end_col)
    confirmation = (
        f'{executor_name} підтверджує, що відповідно до Договору Замовнику були надані маркетингові послуги.\n\n'
        '1. Розміщення Товару Замовника на полицях аптечних закладів Аптечної мережі '
        'відповідно до встановленого Замовником плану викладки згідно з Планограмою та адресною програмою:'
    )
    ws.cell(7, 1, confirmation)
    ws.cell(7, 1).font = header_font
    ws.cell(7, 1).alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[7].height = 60

    # --- Row 9: Table headers ---
    # Headers: № п/п, Pharmacy name, (hidden Відповідальна особа), Brand section header
    headers = ['№ п/п', 'Назва та адреса Аптеки/Аптечного пункту, де були надані маркетингові послуги']

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(9, col_idx, header)
        cell.font = header_font
        cell.alignment = center
        cell.border = thin_border

    # Brand section header - spans across all brand columns (from brand_start_col to brand_end_col)
    ws.merge_cells(start_row=9, start_column=brand_start_col, end_row=9, end_column=brand_end_col)
    brand_header_cell = ws.cell(9, brand_start_col, 'Бренди Товару, щодо яких надавалися маркетингові послуги, кількість викладок')
    brand_header_cell.font = header_font
    brand_header_cell.alignment = center
    brand_header_cell.border = thin_border
    # Apply border to all merged cells
    for col in range(brand_start_col, brand_end_col + 1):
        ws.cell(9, col).border = thin_border

    # --- Row 10: Brand sub-headers ---
    for i, brand in enumerate(template_brands):
        cell = ws.cell(10, brand_start_col + i, brand)
        cell.font = header_font
        cell.alignment = center
        cell.border = thin_border

    # Total column header (only in row 10)
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

        # Brand quantities
        row_total = 0
        for i, brand in enumerate(template_brands):
            qty = row_values.get(brand, 0)
            cell = ws.cell(idx, brand_start_col + i, int(qty) if qty else 0)
            cell.font = normal_font
            cell.alignment = center
            cell.border = thin_border
            cell.number_format = '#,##0'

            price = get_la_price(brand)
            row_total += qty * price

        # Total amount
        cell = ws.cell(idx, total_col, round(row_total, 2))
        cell.font = normal_font
        cell.alignment = right
        cell.border = thin_border
        cell.number_format = '#,##0.00'

    # --- VAT breakdown for Service 1 (only if there's data) ---
    has_service1_data = service1_total > 0 or (pivot_df is not None and not pivot_df.empty)
    if has_service1_data:
        # "Всього без ПДВ" row (replaces the old "ВСЬОГО:" row)
        last_data_row = start_data_row + len(pivot_df) - 1
        subtotal_row = last_data_row + 1
        ws.cell(subtotal_row, total_col - 1, 'Всього без ПДВ').font = header_font
        ws.cell(subtotal_row, total_col - 1).alignment = right
        ws.cell(subtotal_row, total_col - 1).border = thin_border
        subtotal_cell = ws.cell(subtotal_row, total_col, round(service1_total, 2))
        subtotal_cell.font = header_font
        subtotal_cell.alignment = right
        subtotal_cell.number_format = '#,##0.00'
        subtotal_cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='double')
        )

        # VAT row (20%)
        vat_row = subtotal_row + 1
        ws.cell(vat_row, total_col - 1, 'ПДВ 20%').font = header_font
        ws.cell(vat_row, total_col - 1).alignment = right
        ws.cell(vat_row, total_col - 1).border = thin_border
        vat_amount = round(service1_total * 0.2, 2)
        vat_cell = ws.cell(vat_row, total_col, vat_amount)
        vat_cell.font = header_font
        vat_cell.alignment = right
        vat_cell.number_format = '#,##0.00'
        vat_cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='double')
        )

        # Total with VAT row
        total_vat_row = vat_row + 1
        ws.cell(total_vat_row, total_col - 1, 'Разом з ПДВ').font = header_font
        ws.cell(total_vat_row, total_col - 1).alignment = right
        ws.cell(total_vat_row, total_col - 1).border = thin_border
        total_vat_amount = round(service1_total + vat_amount, 2)
        total_vat_cell = ws.cell(total_vat_row, total_col, total_vat_amount)
        total_vat_cell.font = header_font
        total_vat_cell.alignment = right
        total_vat_cell.number_format = '#,##0.00'
        total_vat_cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='double')
        )

        # ============================================================
        # MHL SECTION (Service 2 - Availability) at the bottom
        # ============================================================
        mhl_start_row = total_vat_row + 2
    else:
        # No Service 1 data - MHL section starts after headers
        mhl_start_row = start_data_row + 1

    # MHL section title
    ws.merge_cells(start_row=mhl_start_row, start_column=1, end_row=mhl_start_row, end_column=end_col)
    ws.cell(mhl_start_row, 1, '2. Маркетингові послуги з забезпечення наявності у аптечному закладі Аптечної мережі Товарів із Рекомендованого переліку Товарів Замовника відповідно до Договору.')
    ws.cell(mhl_start_row, 1).font = header_font
    ws.cell(mhl_start_row, 1).alignment = left
    ws.row_dimensions[mhl_start_row].height = 40

    # MHL headers
    mhl_header_row = mhl_start_row + 1
    # Hide "Бренд" column (column 3)
    mhl_headers = ['№ п/п', 'Назва товару',
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

        is_lic_category = template_brands == LIK_BRANDS_TEMPLATE
        allowed_brands = get_mhl_lik_brands() if is_lic_category else get_mhl_dd_brands()
        brand_col = 'BRAND' if 'BRAND' in work_df.columns else 'brand'

        # Case-insensitive matching with whitespace handling
        # Normalize both MHL brands and allowed brands to lowercase for comparison
        brand_values = work_df[brand_col].fillna('').astype(str).str.strip()
        allowed_brands_lower = {b.lower() for b in allowed_brands}
        work_df = work_df[brand_values.str.lower().isin(allowed_brands_lower)].copy()

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

        # Skip "Бренд" column (column 3 - hidden)

        # PENETRATION as integer percentage with % sign (now column 3)
        penetration_pct = int(round(mhl_row['PENETRATION'] * 100))
        ws.cell(idx, 3, f'{penetration_pct}%').font = normal_font
        ws.cell(idx, 3).alignment = center
        ws.cell(idx, 3).border = thin_border

        ws.cell(idx, 4, round(mhl_row['PRICE_CP'], 2)).font = normal_font
        ws.cell(idx, 4).alignment = right
        ws.cell(idx, 4).border = thin_border
        ws.cell(idx, 4).number_format = '#,##0.00'

        ws.cell(idx, 5, int(mhl_row['SUPPLY'])).font = normal_font
        ws.cell(idx, 5).alignment = center
        ws.cell(idx, 5).border = thin_border

        ws.cell(idx, 6, round(mhl_row['Сума закупленого'], 2)).font = normal_font
        ws.cell(idx, 6).alignment = right
        ws.cell(idx, 6).border = thin_border
        ws.cell(idx, 6).number_format = '#,##0.00'

    mhl_last_data_row = max(mhl_first_data_row, mhl_first_data_row + len(mhl_data_rows) - 1)

    # MHL Total row (Total purchased)
    mhl_total_row = mhl_last_data_row + 1
    ws.cell(mhl_total_row, 1, '').font = header_font
    ws.cell(mhl_total_row, 1).border = thin_border
    ws.cell(mhl_total_row, 2, 'ВСЬОГО закуплено товару, грн.').font = header_font
    ws.cell(mhl_total_row, 2).alignment = right
    ws.cell(mhl_total_row, 2).border = thin_border

    for col in range(3, 6):
        ws.cell(mhl_total_row, col, '').font = header_font
        ws.cell(mhl_total_row, col).border = thin_border

    # "Всього без ПДВ" row - use raw total purchased goods
    total_cell = ws.cell(mhl_total_row, 6, round(service2_raw, 2))
    total_cell.font = header_font
    total_cell.alignment = right
    total_cell.number_format = '#,##0.00'
    total_cell.border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )

    # Reward row
    mhl_reward_row = mhl_total_row + 1
    # Hardcoded date as requested
    reward_text = (
        f'Виногорода за забезпечення наявності у аптечному закладі '
        f'Товарів із Рекомендованого переліку Товару Замовника, '
        f'згідно пункту 2.2.2 Договору про маркетингові послуги №{contract_no} '
        f'від 27 лютого 2026 року'
    )
    ws.merge_cells(start_row=mhl_reward_row, start_column=1, end_row=mhl_reward_row, end_column=5)
    cell = ws.cell(mhl_reward_row, 1, reward_text)
    cell.font = header_font
    cell.alignment = left
    cell.border = thin_border
    # Apply border to all cells in merged range (only first cell gets value)
    for col in range(2, 6):
        ws.cell(mhl_reward_row, col).border = thin_border

    # Reward amount - use pre-calculated reward
    reward_amount = round(service2_reward, 2)
    reward_cell = ws.cell(mhl_reward_row, 6, reward_amount)
    reward_cell.font = header_font
    reward_cell.alignment = right
    reward_cell.number_format = '#,##0.00'
    reward_cell.border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )

    # "Всього без ПДВ" row - same value as reward
    mhl_subtotal_row = mhl_reward_row + 1
    ws.cell(mhl_subtotal_row, 5, 'Всього без ПДВ').font = header_font
    ws.cell(mhl_subtotal_row, 5).alignment = right
    ws.cell(mhl_subtotal_row, 5).border = thin_border
    subtotal_cell = ws.cell(mhl_subtotal_row, 6, reward_amount)
    subtotal_cell.font = header_font
    subtotal_cell.alignment = right
    subtotal_cell.number_format = '#,##0.00'
    subtotal_cell.border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )

    # VAT row (20% of reward) - label in column 5 (right-aligned before amount)
    mhl_vat_row = mhl_subtotal_row + 1
    ws.cell(mhl_vat_row, 5, 'ПДВ 20%').font = header_font
    ws.cell(mhl_vat_row, 5).alignment = right
    ws.cell(mhl_vat_row, 5).border = thin_border
    vat_amount = round(reward_amount * 0.2, 2)
    vat_cell = ws.cell(mhl_vat_row, 6, vat_amount)
    vat_cell.font = header_font
    vat_cell.alignment = right
    vat_cell.number_format = '#,##0.00'
    vat_cell.border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )

    # Total with VAT row (reward + VAT) - label in column 5 (right-aligned before amount)
    mhl_total_vat_row = mhl_vat_row + 1
    ws.cell(mhl_total_vat_row, 5, 'Всього з ПДВ').font = header_font
    ws.cell(mhl_total_vat_row, 5).alignment = right
    ws.cell(mhl_total_vat_row, 5).border = thin_border
    total_vat_amount = round(reward_amount + vat_amount, 2)
    total_vat_cell = ws.cell(mhl_total_vat_row, 6, total_vat_amount)
    total_vat_cell.font = header_font
    total_vat_cell.alignment = right
    total_vat_cell.number_format = '#,##0.00'
    total_vat_cell.border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )

    # ============================================================
    # REQUISITES / SIGNATURE SECTION below MHL
    # ============================================================
    sig_start_row = mhl_total_vat_row + 2

    # Extract director from requisites for selected pharmacy
    director_name = ''
    if df_req is not None and not df_req.empty and selected_pharmacy:
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
                director_name = str(matched_row.get('Директор', matched_row.get('директор', ''))).strip()

    act_type = "ЛІК" if is_lic else "ДД"

    # Line 1: "Цей Звіт є невід'ємною частиною Акта{ДД | ЛІК} №__________ про надання послуг від {date}"
    ws.merge_cells(start_row=sig_start_row, start_column=1, end_row=sig_start_row, end_column=end_col)
    line1_text = f'Цей Звіт є невід\'ємною частиною Акта {act_type} №__________ про надання послуг від {period_end_ukrainian}'
    ws.cell(sig_start_row, 1, line1_text).font = normal_font
    ws.cell(sig_start_row, 1).alignment = left

    # Line 2: "інформація надана в повному обсязі"
    ws.merge_cells(start_row=sig_start_row + 1, start_column=1, end_row=sig_start_row + 1, end_column=end_col)
    ws.cell(sig_start_row + 1, 1, 'інформація надана в повному обсязі').font = normal_font
    ws.cell(sig_start_row + 1, 1).alignment = left

    # Line 3: Director signature (right-aligned)
    ws.merge_cells(start_row=sig_start_row + 2, start_column=1, end_row=sig_start_row + 2, end_column=end_col)
    director_font = Font(bold=True, size=12)
    director_text = f'Директор: {director_name}' if director_name else 'Директор: ____________________'
    ws.cell(sig_start_row + 2, 1, director_text).font = director_font
    ws.cell(sig_start_row + 2, 1).alignment = right

    # Line 4: "Від Замовника" (bold)
    ws.merge_cells(start_row=sig_start_row + 3, start_column=1, end_row=sig_start_row + 3, end_column=end_col)
    ws.cell(sig_start_row + 3, 1, 'Від Замовника').font = header_font
    ws.cell(sig_start_row + 3, 1).alignment = left

    # Line 5: "Перевірено" (bold)
    ws.merge_cells(start_row=sig_start_row + 4, start_column=1, end_row=sig_start_row + 4, end_column=end_col)
    ws.cell(sig_start_row + 4, 1, 'Перевірено').font = header_font
    ws.cell(sig_start_row + 4, 1).alignment = left

    # Line 6: "0"
    ws.merge_cells(start_row=sig_start_row + 5, start_column=1, end_row=sig_start_row + 5, end_column=end_col)
    ws.cell(sig_start_row + 5, 1, '0').font = normal_font
    ws.cell(sig_start_row + 5, 1).alignment = left

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
# TMA SHEETS (під час акції) — NEW IMPLEMENTATION
# ============================================================

import json
from pathlib import Path

TMA_AUTH_FILE = Path(__file__).resolve().parent / 'tma_auth.json'

_tma_auth_cache = None


def load_tma_auth():
    """Load TMA authorization amounts from JSON config."""
    global _tma_auth_cache
    if _tma_auth_cache is not None:
        return _tma_auth_cache
    if TMA_AUTH_FILE.exists():
        try:
            with open(TMA_AUTH_FILE, 'r', encoding='utf-8') as f:
                _tma_auth_cache = json.load(f)
        except Exception:
            _tma_auth_cache = {}
    else:
        _tma_auth_cache = {}
    return _tma_auth_cache


def save_tma_auth(data):
    """Save TMA authorization amounts to JSON config."""
    global _tma_auth_cache
    try:
        with open(TMA_AUTH_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        _tma_auth_cache = data
        return True
    except Exception as e:
        print(f"Error saving TMA auth: {e}")
        return False


def get_tma_auth_amount(chain_name, is_lic=True):
    """Get authorization amount for a chain and category (ЛІК/ДД)."""
    auth_data = load_tma_auth()
    cat = 'ЛІК' if is_lic else 'ДД'
    return auth_data.get(chain_name, {}).get(cat, 0.0)


def set_tma_auth_amount(chain_name, is_lic, amount):
    """Set authorization amount for a chain and category."""
    auth_data = load_tma_auth()
    cat = 'ЛІК' if is_lic else 'ДД'
    if chain_name not in auth_data:
        auth_data[chain_name] = {}
    auth_data[chain_name][cat] = float(amount)
    return save_tma_auth(auth_data)


def get_all_sku_for_category(template_brands, mhl_df=None):
    """Get all SKUs for a category (ЛІК or ДД) from MHL data with their brands and prices.

    If mhl_df is provided, extracts actual SKUs from MHL data for the category's brands.
    Otherwise, returns brand-level info (for backward compatibility with non-TMA uses).
    """
    # Determine which category based on template brands
    is_lic = template_brands == LIK_BRANDS_TEMPLATE

    if is_lic:
        db_brands = get_mhl_lik_brands()
    else:
        db_brands = get_mhl_dd_brands()

    # If MHL data is provided, extract actual SKUs from it
    if mhl_df is not None and not mhl_df.empty:
        sku_col = 'SKU_UA' if 'SKU_UA' in mhl_df.columns else 'Назва товару'
        brand_col = 'BRAND' if 'BRAND' in mhl_df.columns else 'brand'

        # Filter MHL data by category brands (case-insensitive)
        work_df = mhl_df.copy()
        brand_values = work_df[brand_col].fillna('').astype(str).str.strip()
        db_brands_lower = {b.lower() for b in db_brands}
        work_df = work_df[brand_values.str.lower().isin(db_brands_lower)].copy()

        if not work_df.empty:
            # Get unique SKUs with their brands
            result = []
            seen = set()
            for _, row in work_df.iterrows():
                sku_name = str(row.get(sku_col, '')).strip()
                db_brand = str(row.get(brand_col, '')).strip()
                if sku_name and sku_name not in seen:
                    seen.add(sku_name)
                    template_brand = DB_BRAND_TO_TEMPLATE.get(db_brand, db_brand)
                    price = get_la_price(template_brand)
                    result.append({
                        'sku': sku_name,
                        'brand': db_brand,
                        'template_brand': template_brand,
                        'price': price
                    })
            return result

    # Fallback: return brand-level info (for cases without MHL data)
    template_to_db = {}
    for db_brand, template_brand in DB_BRAND_TO_TEMPLATE.items():
        if template_brand not in template_to_db:
            template_to_db[template_brand] = []
        template_to_db[template_brand].append(db_brand)

    result = []
    for db_brand in db_brands:
        template_brand = DB_BRAND_TO_TEMPLATE.get(db_brand, db_brand)
        price = get_la_price(template_brand)
        result.append({
            'sku': '',  # No specific SKU without MHL data
            'brand': db_brand,
            'template_brand': template_brand,
            'price': price
        })
    return result


def fill_tma_worksheet(ws, template_brands, period_start, period_end, executor_name,
                       contract_no, contract_date, selected_pharmacy, mhl_df, full_mhl_df=None, is_lic=True, df_req=None, offtake_df=None):
    """Fill 'Під час акції' sheet with full TMA report structure."""

    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right = Alignment(horizontal='right', vertical='center', wrap_text=True)
    header_font = Font(bold=True, size=12)
    normal_font = Font(size=12)
    bold_font = Font(bold=True, size=12)

    # Category label
    act_type = "ЛІК" if is_lic else "ДД"

    # Period formatting (same as in zvit)
    month_names = [
        'січня', 'лютого', 'березня', 'квітня', 'травня', 'червня',
        'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня'
    ]

    start_dt = datetime.strptime(period_start, '%Y-%m-%d')
    end_dt = datetime.strptime(period_end, '%Y-%m-%d')
    if end_dt.month == 12:
        next_month = datetime(end_dt.year + 1, 1, 1)
    else:
        next_month = datetime(end_dt.year, end_dt.month + 1, 1)
    end_last_day = next_month - timedelta(days=1)

    period_str = f'з «{start_dt.day:02d}» {month_names[start_dt.month - 1]} по «{end_last_day.day:02d}» {month_names[end_last_day.month - 1]} {end_last_day.year} року'
    period_end_ukrainian = f'{end_last_day.day:02d} {month_names[end_last_day.month - 1]} {end_last_day.year} року'
    period_short = f'{end_last_day.month:02d}.{end_last_day.year}'  # For "Період проведення акції" column

    # Row counter
    row = 1

    # --- Row 1: Додаток №1 до Акту ---
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row, 1, f'Додаток №1 до Акту №______ про надання послуги від {period_end_ukrainian}').font = header_font
    ws.cell(row, 1).alignment = center
    row += 1

    # --- Row 2: Звіт щодо проведення ТМА ---
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row, 1, f'Звіт щодо проведення ТМА ТОВ "БАЙЄР" за період {period_str}').font = header_font
    ws.cell(row, 1).alignment = center
    row += 1

    # --- Row 3: До Акту ТМА (ЛІК/ДД) ---
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row, 1, f'До Акту ТМА ({act_type}) №_____ про надання послуги від {period_end_ukrainian}').font = header_font
    ws.cell(row, 1).alignment = center
    row += 1

    # --- Row 4: м. Київ + дата ---
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row, 1, f'м. Київ                                                                                                                                                                                                                               {period_end_ukrainian}').font = header_font
    ws.cell(row, 1).alignment = left
    row += 1

    # --- Row 5: Виконавець з реквізитів ---
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row, 1, executor_name).font = header_font
    ws.cell(row, 1).alignment = left
    row += 1

    # --- Row 6: Table headers ---
    headers = [
        '№ п/п',
        'Назва Товару',
        'Розрахункова ціна за одиницю Товару, грн',
        'Розмір наданої знижки',
        'Кількість упаковок Товару, проданого Виконавцем у період',
        'Період проведення акції (місяць і рік)',
        'Вартість послуг, грн'
    ]

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row, col_idx, header)
        cell.font = header_font
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[row].height = 40
    header_row = row
    row += 1

    # --- Get all SKUs for this category ---
    # Use full_mhl_df for complete SKU list across all pharmacies, fallback to mhl_df
    mapping_df = full_mhl_df if full_mhl_df is not None else mhl_df
    sku_data = get_all_sku_for_category(template_brands, mapping_df)

    # --- Try to get PRICE_CP from MHL data for each SKU ---
    # Build a map of SKU -> PRICE_CP from MHL data (filtered by pharmacy)
    mhl_price_map = {}
    mhl_supply_map = {}  # placeholder for future "sold by executor"
    # Build DRUG_ID -> SKU_UA mapping from FULL MHL data for offtake lookup
    drug_id_to_sku = {}

    # Use full_mhl_df for DRUG_ID -> SKU mapping (to include all products across all pharmacies)
    # Fall back to mhl_df if full_mhl_df not provided
    mapping_df = full_mhl_df if full_mhl_df is not None else mhl_df

    if mapping_df is not None and not mapping_df.empty:
        sku_col = 'SKU_UA' if 'SKU_UA' in mapping_df.columns else 'Назва товару'
        work_df = mapping_df.copy()
        # Filter to relevant brands (case-insensitive)
        allowed_brands = get_mhl_lik_brands() if is_lic else get_mhl_dd_brands()
        brand_col = 'BRAND' if 'BRAND' in work_df.columns else 'brand'
        brand_values = work_df[brand_col].fillna('').astype(str).str.strip()
        allowed_brands_lower = {b.lower() for b in allowed_brands}
        work_df = work_df[brand_values.str.lower().isin(allowed_brands_lower)].copy()

        if not work_df.empty:
            for column in ['PRICE_CP', 'SUPPLY', 'DRUG_ID']:
                if column in work_df.columns:
                    work_df[column] = pd.to_numeric(work_df[column], errors='coerce')

            # Build DRUG_ID -> SKU_UA mapping (from full data)
            for _, mhl_row in work_df.iterrows():
                sku_name = str(mhl_row.get(sku_col, '')).strip()
                drug_id = mhl_row.get('DRUG_ID')
                if sku_name and pd.notna(drug_id):
                    drug_id_to_sku[int(drug_id)] = sku_name.casefold()

    # Use mhl_df (filtered by pharmacy) for PRICE_CP and SUPPLY
    if mhl_df is not None and not mhl_df.empty:
        sku_col = 'SKU_UA' if 'SKU_UA' in mhl_df.columns else 'Назва товару'
        work_df = mhl_df.copy()
        allowed_brands = get_mhl_lik_brands() if is_lic else get_mhl_dd_brands()
        brand_col = 'BRAND' if 'BRAND' in work_df.columns else 'brand'
        brand_values = work_df[brand_col].fillna('').astype(str).str.strip()
        allowed_brands_lower = {b.lower() for b in allowed_brands}
        work_df = work_df[brand_values.str.lower().isin(allowed_brands_lower)].copy()

        if not work_df.empty:
            for column in ['PRICE_CP', 'SUPPLY']:
                if column in work_df.columns:
                    work_df[column] = pd.to_numeric(work_df[column], errors='coerce')

            # Get latest PRICE_CP per SKU from pharmacy-specific data
            for _, mhl_row in work_df.iterrows():
                sku_name = str(mhl_row.get(sku_col, '')).strip()
                if sku_name:
                    if pd.notna(mhl_row.get('PRICE_CP')):
                        mhl_price_map[sku_name.casefold()] = mhl_row['PRICE_CP']
                    if pd.notna(mhl_row.get('SUPPLY')):
                        mhl_supply_map[sku_name.casefold()] = mhl_row['SUPPLY']

    # --- Build offtake map: SKU_UA (normalized) -> total OFFTAKE ---
    offtake_map = {}
    unmapped_drug_ids = set()
    if offtake_df is not None and not offtake_df.empty:
        # Offtake data has columns: DATE, CHAIN_ID, CHAIN_NAME, DRUG_ID, OFFTAKE
        # We need to map DRUG_ID to SKU_UA using drug_id_to_sku
        for _, row in offtake_df.iterrows():
            drug_id = row.get('DRUG_ID')
            offtake = row.get('OFFTAKE')
            if pd.notna(drug_id) and pd.notna(offtake):
                drug_id_int = int(drug_id)
                if drug_id_int in drug_id_to_sku:
                    sku_normalized = drug_id_to_sku[drug_id_int]
                    # Sum offtake across all dates for this SKU
                    offtake_map[sku_normalized] = offtake_map.get(sku_normalized, 0) + float(offtake)
                else:
                    unmapped_drug_ids.add(drug_id_int)

    # Log warning for unmapped DRUG_IDs (offtake will be 0 for these)
    if unmapped_drug_ids:
        import logging
        logger = logging.getLogger(__name__)
        logger.warning(
            f"TMA Offtake: {len(unmapped_drug_ids)} DRUG_ID(s) not found in MHL data "
            f"(offtake will be 0): {sorted(unmapped_drug_ids)[:20]}"
            + ("..." if len(unmapped_drug_ids) > 20 else "")
        )

    # --- Data rows ---
    data_start_row = row
    service_values = []  # Store calculated service values for total

    for idx, sku_info in enumerate(sku_data, 1):
        sku_name = sku_info['sku']
        brand = sku_info['brand']
        base_price = sku_info['price']  # From SERVICE_PRICES (fixed price per brand)

        # Get PRICE_CP from MHL data (actual purchase price), fallback to base_price
        price_cp = mhl_price_map.get(sku_name.casefold(), base_price)

        # Discount - use action percentage from TMA action percent file (TMA.xlsx)
        discount_pct = get_tma_action_percent(sku_name, period_start, period_end, selected_pharmacy)

        # Quantity sold by executor - use offtake data if available
        sku_normalized = sku_name.casefold()
        qty_sold = offtake_map.get(sku_normalized, 0.0)
        qty_source = 'offtake'
        # Fallback to MHL supply if no offtake data (placeholder)
        if qty_sold == 0.0:
            qty_sold = mhl_supply_map.get(sku_normalized, 0.0)
            qty_source = 'mhl_supply'

        # Warn if SKU has zero quantity (no offtake, no MHL supply)
        if qty_sold == 0.0:
            import logging
            logger = logging.getLogger(__name__)
            logger.warning(f"TMA: SKU '{sku_name}' has 0 qty (no offtake, no MHL supply)")

        # Period for column 6
        period_display = f'{end_last_day.month:02d}.{end_last_day.year}'

        # Calculate service cost: price_cp * qty_sold * discount_pct / 100
        service_cost = round(price_cp * qty_sold * discount_pct / 100, 2) if discount_pct > 0 and qty_sold > 0 else 0.0
        service_values.append(service_cost)

        # Write row
        ws.cell(row, 1, idx).font = normal_font
        ws.cell(row, 1).alignment = center
        ws.cell(row, 1).border = thin_border

        ws.cell(row, 2, sku_name).font = normal_font
        ws.cell(row, 2).alignment = left
        ws.cell(row, 2).border = thin_border

        ws.cell(row, 3, round(price_cp, 2)).font = normal_font
        ws.cell(row, 3).alignment = right
        ws.cell(row, 3).border = thin_border
        ws.cell(row, 3).number_format = '#,##0.00'

        ws.cell(row, 4, f'{discount_pct:.2f}%').font = normal_font
        ws.cell(row, 4).alignment = center
        ws.cell(row, 4).border = thin_border

        ws.cell(row, 5, round(qty_sold, 2)).font = normal_font
        ws.cell(row, 5).alignment = center
        ws.cell(row, 5).border = thin_border
        ws.cell(row, 5).number_format = '#,##0.00'

        ws.cell(row, 6, period_display).font = normal_font
        ws.cell(row, 6).alignment = center
        ws.cell(row, 6).border = thin_border

        ws.cell(row, 7, service_cost).font = normal_font
        ws.cell(row, 7).alignment = right
        ws.cell(row, 7).border = thin_border
        ws.cell(row, 7).number_format = '#,##0.00'

        row += 1

    data_end_row = row - 1

    # --- Summary rows ---

    # 1. Всього вартість послуг, грн., без ПДВ
    total_services = round(sum(service_values), 2)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    ws.cell(row, 1, 'Всього вартість послуг, грн., без ПДВ').font = bold_font
    ws.cell(row, 1).alignment = left
    ws.cell(row, 1).border = thin_border
    for c in range(2, 7):
        ws.cell(row, c).border = thin_border
    ws.cell(row, 7, total_services).font = bold_font
    ws.cell(row, 7).alignment = right
    ws.cell(row, 7).border = thin_border
    ws.cell(row, 7).number_format = '#,##0.00'
    total_services_row = row
    row += 1

    # 2. Вартість послуг із авторизації, грн., без ПДВ
    auth_amount = get_tma_auth_amount(selected_pharmacy, is_lic)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    ws.cell(row, 1, 'Вартість послуг із авторизації, грн., без ПДВ').font = bold_font
    ws.cell(row, 1).alignment = left
    ws.cell(row, 1).border = thin_border
    for c in range(2, 7):
        ws.cell(row, c).border = thin_border
    ws.cell(row, 7, auth_amount).font = bold_font
    ws.cell(row, 7).alignment = right
    ws.cell(row, 7).border = thin_border
    ws.cell(row, 7).number_format = '#,##0.00'
    auth_row = row
    row += 1

    # 3. Всього без ПДВ (text under "Період" column, value under "Вартість послуг")
    subtotal_no_vat = round(total_services + auth_amount, 2)
    # Text in column 6 (Період)
    ws.cell(row, 6, 'Всього без ПДВ').font = bold_font
    ws.cell(row, 6).alignment = right
    ws.cell(row, 6).border = thin_border
    # Value in column 7
    ws.cell(row, 7, subtotal_no_vat).font = bold_font
    ws.cell(row, 7).alignment = right
    ws.cell(row, 7).border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )
    ws.cell(row, 7).number_format = '#,##0.00'
    subtotal_row = row
    row += 1

    # 4. ПДВ 20%
    vat_amount = round(subtotal_no_vat * 0.2, 2)
    ws.cell(row, 6, 'ПДВ 20%').font = bold_font
    ws.cell(row, 6).alignment = right
    ws.cell(row, 6).border = thin_border
    ws.cell(row, 7, vat_amount).font = bold_font
    ws.cell(row, 7).alignment = right
    ws.cell(row, 7).border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )
    ws.cell(row, 7).number_format = '#,##0.00'
    row += 1

    # 5. Разом з ПДВ
    total_with_vat = round(subtotal_no_vat + vat_amount, 2)
    ws.cell(row, 6, 'Разом з ПДВ').font = bold_font
    ws.cell(row, 6).alignment = right
    ws.cell(row, 6).border = thin_border
    ws.cell(row, 7, total_with_vat).font = bold_font
    ws.cell(row, 7).alignment = right
    ws.cell(row, 7).border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='double')
    )
    ws.cell(row, 7).number_format = '#,##0.00'

    # ============================================================
    # PHOTO SECTION (Фотопідтвердження)
    # ============================================================
    photo_start_row = row + 2

    # Header for photo section
    ws.merge_cells(start_row=photo_start_row, start_column=1, end_row=photo_start_row, end_column=7)
    ws.cell(photo_start_row, 1, 'Фотопідтвердження наданих послуг за звітний період.').font = header_font
    ws.cell(photo_start_row, 1).alignment = center
    ws.row_dimensions[photo_start_row].height = 30

    # 6 photo placeholders in 3x2 grid
    # Each photo spans ~2-3 columns, 2 rows of photos
    photo_grid_start = photo_start_row + 1
    photo_height = 60  # Height for each photo row

    # Column layout: 7 columns total
    # Photo 1: A-B, Photo 2: C-D, Photo 3: E-F (G for spacing/border)
    # Photo 4: A-B, Photo 5: C-D, Photo 6: E-F
    photo_cols = [
        (1, 2),   # Photo 1: cols A-B
        (3, 4),   # Photo 2: cols C-D
        (5, 6),   # Photo 3: cols E-F
        (1, 2),   # Photo 4: cols A-B
        (3, 4),   # Photo 5: cols C-D
        (5, 6),   # Photo 6: cols E-F
    ]

    for i, (start_col, end_col) in enumerate(photo_cols):
        row_idx = photo_grid_start + (i // 3) * 3  # Each photo takes 2 rows + 1 gap
        if i >= 3:
            row_idx = photo_grid_start + 3  # Second row of photos starts 3 rows below first

        # Merge cells for photo placeholder
        ws.merge_cells(start_row=row_idx, start_column=start_col, end_row=row_idx + 1, end_column=end_col)

        photo_cell = ws.cell(row_idx, start_col)
        photo_num = i + 1
        photo_cell.value = f'[Фото {photo_num}]'
        photo_cell.font = Font(size=11, color='808080', italic=True)
        photo_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Add border around photo placeholder
        for r in range(row_idx, row_idx + 2):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(r, c)
                cell.border = thin_border
                if r == row_idx and c == start_col:
                    cell.value = f'[Фото {photo_num}]'
                    cell.font = Font(size=11, color='808080', italic=True)
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        ws.row_dimensions[row_idx].height = photo_height
        ws.row_dimensions[row_idx + 1].height = photo_height

    # ============================================================
    # SIGNATURE / REQUISITES SECTION (під час акції)
    # ============================================================
    sig_start_row = photo_grid_start + 6  # After 2 rows of photos (3 rows each)

    # Extract director and BAM from requisites for selected pharmacy
    director_name = ''
    bam_name = ''
    if df_req is not None and not df_req.empty and selected_pharmacy:
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
                director_name = str(matched_row.get('Директор', matched_row.get('директор', ''))).strip()
                # Try to get BAM (Business Area Manager) from requisites
                bam_name = str(matched_row.get('BAM', matched_row.get('БАМ', matched_row.get('bam', '')))).strip()

    act_type = "ЛІК" if is_lic else "ДД"

    # Line 1: "Цей Звіт є невід'ємною частиною Акта{ЛІК|ДД} №__________ про надання послуг від {date}"
    ws.merge_cells(start_row=sig_start_row, start_column=1, end_row=sig_start_row, end_column=7)
    line1_text = f'Цей Звіт є невід\'ємною частиною Акта {act_type} №__________ про надання послуг від {period_end_ukrainian}'
    ws.cell(sig_start_row, 1, line1_text).font = normal_font
    ws.cell(sig_start_row, 1).alignment = left

    # Line 2: Director signature (right-aligned) - only name, no "Директор:" prefix
    ws.merge_cells(start_row=sig_start_row + 1, start_column=1, end_row=sig_start_row + 1, end_column=7)
    director_font = Font(bold=True, size=12)
    director_text = director_name if director_name else '____________________'
    ws.cell(sig_start_row + 1, 1, director_text).font = director_font
    ws.cell(sig_start_row + 1, 1).alignment = right

    # Line 3: "Від Замовника" (left, bold)
    ws.merge_cells(start_row=sig_start_row + 2, start_column=1, end_row=sig_start_row + 2, end_column=7)
    ws.cell(sig_start_row + 2, 1, 'Від Замовника').font = header_font
    ws.cell(sig_start_row + 2, 1).alignment = left

    # Line 4: "Перевірено" (left, bold)
    ws.merge_cells(start_row=sig_start_row + 3, start_column=1, end_row=sig_start_row + 3, end_column=7)
    ws.cell(sig_start_row + 3, 1, 'Перевірено').font = header_font
    ws.cell(sig_start_row + 3, 1).alignment = left

    # Line 5: "Менеджер з питань регіонального розвитку бізнесу {БАМ з реквізитів}" (left)
    ws.merge_cells(start_row=sig_start_row + 4, start_column=1, end_row=sig_start_row + 4, end_column=7)
    bam_text = f'Менеджер з питань регіонального розвитку бізнесу {bam_name}' if bam_name else 'Менеджер з питань регіонального розвитку бізнесу ____________________'
    ws.cell(sig_start_row + 4, 1, bam_text).font = normal_font
    ws.cell(sig_start_row + 4, 1).alignment = left

    # --- Column widths ---
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 35
    ws.column_dimensions['F'].width = 22
    ws.column_dimensions['G'].width = 20

    ws.title = f'під час акції ({act_type})'

    # Return calculated amounts for ACT_TMA sheet
    return {
        'total_services': total_services,
        'auth_amount': auth_amount,
        'subtotal_no_vat': subtotal_no_vat,
        'vat_amount': vat_amount,
        'total_with_vat': total_with_vat
    }


# Legacy empty sheet function (kept for compatibility if needed)
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
        df_req_lik,
        df_req_dd,
        output_filename='marketing_report.xlsx',
        period_start=None,
        period_end=None,
        selected_pharmacy='Усі аптеки',
        mhl_df=None,
        full_mhl_df=None,
        offtake_df=None,
):
    """Generate Excel report with 8 sheets matching template exactly:
    1. акт_ЛІК
    2. акт_ДД
    3. звіт викладка+MHL(ЛІК)
    4. звіт викладка+MHL(ДД)
    5. під час акції (ЛІК)
    6. під час акції (ДД)
    7. АКТ_ТМА_ЛІК
    8. АКТ_ТМА_ДД
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
    # This is the total purchased goods where penetration >= target
    service2_lik_raw = get_mhl_availability_sum(mhl_df, LIK_BRANDS_TEMPLATE)
    service2_dd_raw = get_mhl_availability_sum(mhl_df, DD_BRANDS_TEMPLATE)

    # Calculate reward amounts (what goes into ACT sheets)
    # Reward = service2_total * MHL_REWARD_MULTIPLIER (e.g., 10%)
    service2_lik = round(service2_lik_raw * MHL_REWARD_MULTIPLIER, 2)
    service2_dd = round(service2_dd_raw * MHL_REWARD_MULTIPLIER, 2)

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

    # Helper to extract contract details from requisites
    def _extract_contract_details(df_req, pharmacy_name):
        """Extract executor_name, contract_no, contract_date from requisites for a pharmacy."""
        if df_req is None or df_req.empty:
            return DEFAULT_EXECUTOR_NAME, DEFAULT_CONTRACT_NO, DEFAULT_CONTRACT_DATE

        clean_selected = str(pharmacy_name).strip()
        target_col = None
        for col_candidate in ['chain_name', 'Клиент из базы данных', 'Мережа', 'Назва']:
            if col_candidate in df_req.columns:
                target_col = col_candidate
                break

        matched_row = None
        if clean_selected and clean_selected != 'Усі аптеки' and target_col:
            filtered_req = df_req[
                df_req[target_col].astype(str).str.strip().str.lower()
                == clean_selected.lower()
            ]
            if not filtered_req.empty:
                matched_row = filtered_req.iloc[0]

        if matched_row is None:
            matched_row = df_req.iloc[0]

        executor = str(
            matched_row.get('Виконавець', matched_row.get('Юридическое лицо', DEFAULT_EXECUTOR_NAME))
        )
        contract_no = str(matched_row.get('№ договора', DEFAULT_CONTRACT_NO))
        contract_date = str(matched_row.get('Дата договора', DEFAULT_CONTRACT_DATE))
        return executor, contract_no, contract_date

    # Extract contract details separately for ЛИК and ДД
    executor_lik, contract_no_lik, contract_date_lik = _extract_contract_details(df_req_lik, selected_pharmacy)
    executor_dd, contract_no_dd, contract_date_dd = _extract_contract_details(df_req_dd, selected_pharmacy)

    import openpyxl
    wb = openpyxl.Workbook()
    default_sheet = wb.active

    # ============================================================
    # SHEET 1: акт_ЛІК (uses act_generator for correct act structure)
    # ============================================================
    ws_akt_lik = wb.create_sheet(title='акт_ЛІК')
    generate_act_sheet(
        ws=ws_akt_lik,
        df_req=df_req_lik,
        service1_sum=service1_lik,
        service2_sum=service2_lik,
        period_start=period_start,
        period_end=period_end,
        selected_pharmacy=selected_pharmacy,
        act_number='1',
        contract_no=contract_no_lik,
        contract_date=contract_date_lik,
        executor_name=executor_lik,
        is_lic=True
    )

    # ============================================================
    # SHEET 2: акт_ДД
    # ============================================================
    ws_akt_dd = wb.create_sheet(title='акт_ДД')
    generate_act_sheet(
        ws=ws_akt_dd,
        df_req=df_req_dd,
        service1_sum=service1_dd,
        service2_sum=service2_dd,
        period_start=period_start,
        period_end=period_end,
        selected_pharmacy=selected_pharmacy,
        act_number='2',
        contract_no=contract_no_dd,
        contract_date=contract_date_dd,
        executor_name=executor_dd,
        is_lic=False
    )

    # ============================================================
    # SHEET 3: звіт викладка+MHL(ЛІК)
    # ============================================================
    ws_zvit_lik = wb.create_sheet(title='звіт викладка+MHL(ЛІК)')
    fill_zvit_worksheet(ws_zvit_lik, pivot_lik, LIK_BRANDS_TEMPLATE, period_str,
                        period_end_ukrainian,
                        executor_lik, contract_no_lik, contract_date_lik,
                        service1_lik, service2_lik_raw, service2_lik, is_lic=True, mhl_df=mhl_df,
                        df_req=df_req_lik, selected_pharmacy=selected_pharmacy)

    # ============================================================
    # SHEET 4: звіт викладка+MHL(ДД)
    # ============================================================
    ws_zvit_dd = wb.create_sheet(title='звіт викладка+MHL(ДД)')
    fill_zvit_worksheet(ws_zvit_dd, pivot_dd, DD_BRANDS_TEMPLATE, period_str,
                        period_end_ukrainian,
                        executor_dd, contract_no_dd, contract_date_dd,
                        service1_dd, service2_dd_raw, service2_dd, is_lic=False, mhl_df=mhl_df,
                        df_req=df_req_dd, selected_pharmacy=selected_pharmacy)

    # ============================================================
    # SHEET 5: під час акції (ЛІК) — full TMA report
    # ============================================================
    ws_tma_lik = wb.create_sheet(title='під час акції (ЛІК)')
    tma_lik_amounts = fill_tma_worksheet(
        ws=ws_tma_lik,
        template_brands=LIK_BRANDS_TEMPLATE,
        period_start=period_start,
        period_end=period_end,
        executor_name=executor_lik,
        contract_no=contract_no_lik,
        contract_date=contract_date_lik,
        selected_pharmacy=selected_pharmacy,
        mhl_df=mhl_df,
        full_mhl_df=full_mhl_df,
        is_lic=True,
        df_req=df_req_lik,
        offtake_df=offtake_df
    )

    # ============================================================
    # SHEET 6: під час акції (ДД) — full TMA report
    # ============================================================
    ws_tma_dd = wb.create_sheet(title='під час акції (ДД)')
    tma_dd_amounts = fill_tma_worksheet(
        ws=ws_tma_dd,
        template_brands=DD_BRANDS_TEMPLATE,
        period_start=period_start,
        period_end=period_end,
        executor_name=executor_dd,
        contract_no=contract_no_dd,
        contract_date=contract_date_dd,
        selected_pharmacy=selected_pharmacy,
        mhl_df=mhl_df,
        full_mhl_df=full_mhl_df,
        is_lic=False,
        df_req=df_req_dd,
        offtake_df=offtake_df
    )

    # ============================================================
    # SHEET 7: АКТ_ТМА_ЛІК
    # ============================================================
    ws_akt_tma_lik = wb.create_sheet(title='АКТ_ТМА_ЛІК')
    generate_tma_act_sheet(
        ws=ws_akt_tma_lik,
        df_req=df_req_lik,
        tma_subtotal_no_vat=tma_lik_amounts['subtotal_no_vat'] if tma_lik_amounts else 0,
        period_start=period_start,
        period_end=period_end,
        selected_pharmacy=selected_pharmacy,
        act_number='3',
        contract_no=contract_no_lik,
        contract_date=contract_date_lik,
        executor_name=executor_lik,
        is_lic=True
    )

    # ============================================================
    # SHEET 8: АКТ_ТМА_ДД
    # ============================================================
    ws_akt_tma_dd = wb.create_sheet(title='АКТ_ТМА_ДД')
    generate_tma_act_sheet(
        ws=ws_akt_tma_dd,
        df_req=df_req_dd,
        tma_subtotal_no_vat=tma_dd_amounts['subtotal_no_vat'] if tma_dd_amounts else 0,
        period_start=period_start,
        period_end=period_end,
        selected_pharmacy=selected_pharmacy,
        act_number='4',
        contract_no=contract_no_dd,
        contract_date=contract_date_dd,
        executor_name=executor_dd,
        is_lic=False
    )

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
                        price = get_la_price(brand)
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

    # Use brand-based filtering (case-insensitive)
    brand_col = 'BRAND' if 'BRAND' in df.columns else 'brand'
    if brand_col not in df.columns:
        return 0, 0, 0

    brand_names = df[brand_col].fillna('').astype(str).str.strip().str.lower()

    lik_brands = {b.lower() for b in get_mhl_lik_brands()}
    dd_brands = {b.lower() for b in get_mhl_dd_brands()}

    lik_match = brand_names.isin(lik_brands)
    dd_match = brand_names.isin(dd_brands)
    return int(lik_match.sum()), int(dd_match.sum()), int(
        (~(lik_match | dd_match)).sum()
    )


def mhl_unknown_samples(df, limit=8):
    """Return real database values that are not classified by the brand lists."""
    if df is None or df.empty:
        return []

    brand_col = 'BRAND' if 'BRAND' in df.columns else 'brand'
    if brand_col not in df.columns:
        return []

    brand_names = df[brand_col].fillna('').astype(str).str.strip().str.lower()

    lik_brands = {b.lower() for b in get_mhl_lik_brands()}
    dd_brands = {b.lower() for b in get_mhl_dd_brands()}
    known_brands = lik_brands | dd_brands

    unknown = df.loc[
        ~brand_names.isin(known_brands),
        [brand_col, 'SKU_UA'],
    ].drop_duplicates()
    return [
        f'{brand} / {sku}'
        for brand, sku in unknown.head(limit).itertuples(index=False, name=None)
    ]