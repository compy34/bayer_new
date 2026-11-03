import pandas as pd
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from act_generator import generate_act_sheet

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

# MHL SKU lists for filtering MHL data
MHL_LIK_SKU = [
    'Еріус сироп 0.5 мг/мл 60мл',
    'Еріус таблетки 5 мг № 10',
    'Антифлу Кідс, порошок пакетики №5',
    'Аспірин С, таб. № 10',
    'Кальцемін Адванс №60',
    'Кальцемін Сільвер №60',
    'Бепантен, крем 5%, туба 30г №1',
    'Бепантен, мазь 5%, туба 30г №1',
    'Бепантен, мазь 5%, туба 100г №1',
    'Бепантен Плюс, крем, туба 30г №1',
    'Елевіт Пронаталь, т.в/о №100',
    'Реліф, супозиторїї ректальні №12',
    'Реліф Адванс, супозиторїї ректальні №12',
    'Реліф Про, супозиторїї ректальні №10',
    'Реліф Ультра, супозиторїї ректальні №12',
    'Ренні, табл. жув., мент. смак. №24',
    "Ренні б/ц м'ятні,табл. жув.№24",
    'Ренні з апельсиновим смаком, табл. жув. №24',
    'Тералів, табл. 220мг №12',
    'Терафлекс, капсули №60',
    'Терафлекс Адванс, капсули №60',
    'Терафлекс, капсули №120',
]

MHL_DD_SKU = [
    'Супрадин Актив, табл. №30',
    'Супрадин Актив, табл. шип., №15',
    'Супрадин Ведмежуйки №30',
    'Супрадин Ведмежуйки №60',
    'Супрадин Бустер, табл. шип., №15',
    'Бепантен Сенсідерм',
    'Елевіт-Комплекс 2, капс. № 30',
    'Елевіт-Комплекс 3, капс. № 30',
]


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

    # Filter MHL data for relevant SKUs
    is_lic = template_brands == LIK_BRANDS_TEMPLATE
    allowed_values = {str(v).strip().casefold() for v in (MHL_LIK_SKU if is_lic else MHL_DD_SKU)}
    sku_values = work_df[sku_col].fillna('').astype(str).str.strip().str.casefold()
    work_df = work_df[sku_values.isin(allowed_values)].copy()

    if work_df.empty:
        return 0.0

    for column in ['PENETRATION', 'PRICE_CP', 'SUPPLY', 'PENETRATION_TARGET']:
        work_df[column] = pd.to_numeric(work_df[column], errors='coerce').fillna(0)

    # Sum of (SUPPLY * PRICE_CP) where PENETRATION >= PENETRATION_TARGET
    work_df['Сума закупленого'] = work_df['SUPPLY'] * work_df['PRICE_CP']
    work_df['eligible'] = work_df['PENETRATION'] >= work_df['PENETRATION_TARGET']

    total = work_df.loc[work_df['eligible'], 'Сума закупленого'].sum()
    return round(total, 2)


# ============================================================
# ZVIT SHEET GENERATOR (звіт викладка+MHL) - matches template exactly
# ============================================================

def fill_zvit_worksheet(ws, pivot_df, template_brands, period_str, executor_name, contract_no, contract_date,
                         service1_total, service2_total, is_lic=True):
    """Fill звіт викладка+MHL sheet matching template structure exactly."""
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
    ws.cell(6, 1, f'м. Київ                                                                                                                                                                                                                               {period_str}')
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

    # --- Total row ---
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

    # --- Column widths ---
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 30
    for i in range(num_brands):
        ws.column_dimensions[get_column_letter(brand_start_col + i)].width = 12
    ws.column_dimensions[get_column_letter(total_col)].width = 18


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
    period_str = f'з «{period_start[:2]}» {period_start[3:]} по «{period_end[:2]}» {period_end[3:]}'

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
                        executor_name, contract_no, contract_date,
                        service1_lik, service2_lik, is_lic=True)

    # ============================================================
    # SHEET 4: звіт викладка+MHL(ДД)
    # ============================================================
    ws_zvit_dd = wb.create_sheet(title='звіт викладка+MHL(ДД)')
    fill_zvit_worksheet(ws_zvit_dd, pivot_dd, DD_BRANDS_TEMPLATE, period_str,
                        executor_name, contract_no, contract_date,
                        service1_dd, service2_dd, is_lic=False)

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
    if df is None or df.empty or 'BRAND' not in df.columns:
        return 0, 0, 0
    sku_names = df['SKU_UA'].fillna('').astype(str).str.strip().str.casefold()
    lik = {str(value).strip().casefold() for value in MHL_LIK_SKU}
    dd = {str(value).strip().casefold() for value in MHL_DD_SKU}
    lik_match = sku_names.isin(lik)
    dd_match = sku_names.isin(dd)
    return int(lik_match.sum()), int(dd_match.sum()), int(
        (~(lik_match | dd_match)).sum()
    )


def mhl_unknown_samples(df, limit=8):
    """Return real database values that are not classified by the lists."""
    if df is None or df.empty:
        return []
    sku_names = df['SKU_UA'].fillna('').astype(str).str.strip()
    known = {
        str(value).strip().casefold()
        for value in MHL_LIK_SKU + MHL_DD_SKU
    }
    unknown = df.loc[
        ~sku_names.str.casefold().isin(known),
        ['BRAND', 'SKU_UA'],
    ].drop_duplicates()
    return [
        f'{brand} / {sku}'
        for brand, sku in unknown.head(limit).itertuples(index=False, name=None)
    ]