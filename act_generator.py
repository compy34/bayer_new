from datetime import datetime, timedelta
import pandas as pd
from num2words import num2words
from openpyxl.styles import Alignment, Border, Font, Side
from config import (
    DEFAULT_CONTRACT_NO,
    DEFAULT_CONTRACT_DATE,
    DEFAULT_TMA_CONTRACT_DATE,
    DEFAULT_EXECUTOR_NAME,
)

# Column name variants for client (user may have different column names in requisites file)
CLIENT_COL_VARIANTS = [
    'Клиент из базы данных', 'Клиент с базы данных', 'Клиент из базы', 'Клиент с базы',
    'Замовник', 'Клиент', 'Customer', 'Покупець', 'Замовник з бази', 'Client'
]


def generate_act_sheet(
        ws,
        df_req,
        service1_sum,
        service2_sum,
        period_start,
        period_end,
        selected_pharmacy='Усі аптеки',
        act_number='1',
        contract_no=DEFAULT_CONTRACT_NO,
        contract_date=DEFAULT_CONTRACT_DATE,
        executor_name=DEFAULT_EXECUTOR_NAME,
        is_lic=True
):
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000'),
    )
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right_alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)

    # Дефолтні реквізити на випадок, якщо файл не завантажено або колонка порожня
    full_executor_requisites = [
        executor_name,
        'Юридична адреса:',
        '69050, м. Запоріжжя',
        'вул. Складська, буд. 6',
        'Код ЄДРПОУ 01977334',
        'ІПН 019773308278, св. № 200049431',
        'Банківські реквізити:',
        'IBAN UA633204780000026007924921260',
        'в АБ "УКРГАЗБАНК"',
        'МФО 320478',
        'Ел. пошта: sales1@apteki.zp.ua',
        'Голова правління',
        '_____________________ / ________ /',
    ]

    customer_name = 'ТОВ «БАЙЄР»'
    legal_entity = executor_name  # default for row 3 - use executor name if no requisites

    if df_req is not None and not df_req.empty:
        matched_row = None
        clean_selected = str(selected_pharmacy).strip()
        target_col = None
        for col_candidate in ['chain_name'] + CLIENT_COL_VARIANTS + ['Мережа', 'Назва']:
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

        # Helper to validate values aren't client identifiers
        def _is_valid_entity(val):
            if val is None:
                return False
            if isinstance(val, float) and pd.isna(val):
                return False
            s = str(val).strip()
            if not s:
                return False
            s_lower = s.lower()
            client_keywords = ['клиент', 'база', 'customer', 'client', 'покупець', 'замовник']
            if any(kw in s_lower for kw in client_keywords):
                return False
            return True

        # Виконавець — використовуємо вибрану аптеку/мережу, якщо вибрано конкретну
        if clean_selected and clean_selected != 'Усі аптеки':
            executor_name = clean_selected
        else:
            # Якщо "Усі аптеки" — беремо з реквізитів (Виконавець або Юридическое лицо)
            # But reject client-like values
            executor_val = matched_row.get('Виконавець')
            if not _is_valid_entity(executor_val):
                executor_val = matched_row.get('Юридическое лицо')
            if not _is_valid_entity(executor_val):
                executor_val = executor_name  # DEFAULT_EXECUTOR_NAME
            executor_name = str(executor_val)

        # Юридическое лицо — для використання в рядку "між ТОВ "БАЙЄР" та ..."
        # ONLY use 'Юридическое лицо' column for legal entity
        # If not available or invalid, fall back to executor_name
        # Final safety: if legal_entity still looks like client identifier, use safe default
        legal_entity_val = matched_row.get('Юридическое лицо')
        if not _is_valid_entity(legal_entity_val):
            legal_entity_val = executor_name
        # Final safety check
        if not _is_valid_entity(legal_entity_val):
            legal_entity_val = clean_selected if clean_selected and clean_selected != 'Усі аптеки' else DEFAULT_EXECUTOR_NAME
        legal_entity = str(legal_entity_val)
        # Замовник — пробуємо різні варіанти назв колонки
        customer_name = customer_name  # default
        for col in CLIENT_COL_VARIANTS:
            if col in matched_row.index and pd.notna(matched_row[col]) and str(matched_row[col]).strip():
                customer_name = str(matched_row[col]).strip()
                break
        contract_no = str(matched_row.get('№ договора', contract_no))
        contract_date = str(matched_row.get('Дата договора', contract_date))

        # БЕРЕМО ГОТОВІ РЕКВІЗИТИ З КОЛОНКИ ФАЙЛУ (для Виконавця)
        raw_req_text = str(matched_row.get('Реквизити', ''))
        if raw_req_text and raw_req_text.lower() != 'nan':
            full_executor_requisites = [
                line.strip() for line in raw_req_text.split('\n') if line.strip()
            ]

    # Act type for header
    act_type = "ЛІК" if is_lic else "ДД"

    # Row 1: "АКТ ДД|ЛІК №_____"
    ws.merge_cells('A1:E1')
    ws['A1'] = f'АКТ {act_type} № _____'
    ws['A1'].font = Font(bold=True, size=11)
    ws['A1'].alignment = center_alignment

    # Row 2: "про надання послуги, відповідно до Договору про маркетингові послуги {номер} від {дата}"
    ws.merge_cells('A2:E2')
    ws['A2'] = (
        f'про надання послуги, відповідно до Договору про маркетингові послуги '
        f'№ {contract_no} від {contract_date}'
    )
    ws['A2'].font = Font(bold=True, size=10)
    ws['A2'].alignment = center_alignment
    ws.row_dimensions[2].height = 30

    # Row 3: "між ТОВ "БАЙЄР" та {Юридическое лицо}"
    ws.merge_cells('A3:E3')
    ws['A3'] = f'між ТОВ «БАЙЄР» та {legal_entity}'
    ws['A3'].font = Font(bold=True, size=10)
    ws['A3'].alignment = center_alignment
    ws.row_dimensions[3].height = 30

    ws['A5'] = 'м. Київ'
    ws['A5'].font = Font(size=10)
    ws['A5'].alignment = left_alignment

    ws.merge_cells('D5:E5')
    # Format period_end properly (user enters in format YYYY-MM-DD, convert to DD.MM.YYYY)
    try:
        from datetime import datetime
        end_dt = datetime.strptime(period_end, '%Y-%m-%d')
        month_names = [
            'січня', 'лютого', 'березня', 'квітня', 'травня', 'червня',
            'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня'
        ]
        period_end_formatted = f'"{end_dt.day:02d}" {month_names[end_dt.month - 1]} {end_dt.year} року'
    except Exception:
        period_end_formatted = period_end
    ws['D5'] = period_end_formatted
    ws['D5'].font = Font(size=10)
    ws['D5'].alignment = right_alignment

    ws.merge_cells('A7:E7')
    ws['A7'] = (
        f'{executor_name} виконав, а {customer_name} прийняв наступні послуги '
        f'згідно з договором № {contract_no} від {contract_date}.'
    )
    ws['A7'].font = Font(size=10)
    ws['A7'].alignment = left_alignment
    ws.row_dimensions[7].height = 35

    ws['A9'] = (
        f'1. Період надання послуг: з "{period_start}" по "{period_end}".'
    )
    ws['A9'].font = Font(size=10)
    ws['A10'] = '2. Якість послуг відповідає умовам договору.'
    ws['A10'].font = Font(size=10)
    ws['A12'] = '3. Перелік наданих послуг:'
    ws['A12'].font = Font(bold=True, size=10)

    headers = [
        'Найменування послуг',
        'Од.',
        'Сума без ПДВ, грн.',
        'ПДВ 20%',
        'Загальна сума з ПДВ, грн.',
    ]
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=14, column=col_idx, value=h_text)
        cell.font = Font(bold=True, size=9)
        cell.alignment = center_alignment
        cell.border = thin_border
    ws.row_dimensions[14].height = 25

    # Послуга 1: Викладка
    # Input amounts are WITHOUT VAT (same as in report/zvit)
    sum1_no_vat = float(service1_sum)
    vat1 = round(sum1_no_vat * 0.2, 2)
    total1 = round(sum1_no_vat + vat1, 2)

    # Послуга 2: Наявність (MHL)
    # service2_sum = reward_amount (WITHOUT VAT, matches MHL section "Всього без ПДВ")
    sum2_no_vat = float(service2_sum)
    vat2 = round(sum2_no_vat * 0.2, 2)
    total2 = round(sum2_no_vat + vat2, 2)

    services_data = [
        (
            'Маркетингові послуги з розміщення Товару Замовника на полицях аптечних закладів аптечної мережі відповідно до встановленого Замовником плану викладки згідно з Планограмою та адресною програмою, затверджених Додатком.',
            'послуга',
            sum1_no_vat,
            vat1,
            total1,
        ),
        (
            'Маркетингові послуги з забезпечення наявності у аптечному закладі аптечної мережі Товарів із Рекомендованого переліку Товарів Замовника відповідно до Договору.',
            'послуга',
            sum2_no_vat,
            vat2,
            total2,
        ),
    ]

    row_idx = 15
    for s_name, s_unit, s_sum, s_vat, s_tot in services_data:
        c1 = ws.cell(row=row_idx, column=1, value=s_name)
        c1.alignment = left_alignment
        c1.border = thin_border
        c2 = ws.cell(row=row_idx, column=2, value=s_unit)
        c2.alignment = center_alignment
        c2.border = thin_border
        c3 = ws.cell(row=row_idx, column=3, value=s_sum)
        c3.alignment = right_alignment
        c3.border = thin_border
        c3.number_format = '#,##0.00'
        c4 = ws.cell(row=row_idx, column=4, value=s_vat)
        c4.alignment = right_alignment
        c4.border = thin_border
        c4.number_format = '#,##0.00'
        c5 = ws.cell(row=row_idx, column=5, value=s_tot)
        c5.alignment = right_alignment
        c5.border = thin_border
        c5.number_format = '#,##0.00'
        ws.row_dimensions[row_idx].height = 45  # Increased height for long text
        row_idx += 1

    last_service_row = row_idx - 1

    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=2)
    total_label_cell = ws.cell(row=row_idx, column=1, value='Всього:')
    total_label_cell.font = Font(bold=True, size=9)
    total_label_cell.alignment = right_alignment
    total_label_cell.border = thin_border
    ws.cell(row=row_idx, column=2).border = thin_border

    c_sum = ws.cell(row=row_idx, column=3, value=f'=SUM(C14:C{last_service_row})')
    c_sum.font = Font(bold=True, size=9)
    c_sum.alignment = right_alignment
    c_sum.border = thin_border
    c_sum.number_format = '#,##0.00'

    c_vat = ws.cell(row=row_idx, column=4, value=f'=SUM(D14:D{last_service_row})')
    c_vat.font = Font(bold=True, size=9)
    c_vat.alignment = right_alignment
    c_vat.border = thin_border
    c_vat.number_format = '#,##0.00'

    c_tot = ws.cell(row=row_idx, column=5, value=f'=SUM(E14:E{last_service_row})')
    c_tot.font = Font(bold=True, size=9)
    c_tot.alignment = right_alignment
    c_tot.border = thin_border
    c_tot.number_format = '#,##0.00'

    row_idx += 2

    total_sum_val = sum1_no_vat + sum2_no_vat
    total_vat_val = total_sum_val * 0.2
    grand_total_val = total_sum_val + total_vat_val

    def format_in_words(amount):
        uah_int = int(amount)
        kop = int(round((amount - uah_int) * 100))
        uah_text = num2words(uah_int, lang='uk')
        return f'{uah_text} грн. {kop:02d} коп.'

    text_lines = [
        ('Всього без ПДВ:', total_sum_val, format_in_words(total_sum_val)),
        ('ПДВ 20%:', total_vat_val, format_in_words(total_vat_val)),
        ('Всього з ПДВ:', grand_total_val, format_in_words(grand_total_val)),
    ]

    for label, num_val, words_str in text_lines:
        ws.cell(row=row_idx, column=1, value=label).font = Font(bold=True, size=9)
        ws.cell(row=row_idx, column=1).alignment = left_alignment
        num_cell = ws.cell(row=row_idx, column=3, value=num_val)
        num_cell.font = Font(bold=True, size=9)
        num_cell.alignment = right_alignment
        num_cell.number_format = '#,##0.00'
        ws.merge_cells(start_row=row_idx, start_column=4, end_row=row_idx, end_column=5)
        words_cell = ws.cell(row=row_idx, column=4, value=words_str)
        words_cell.font = Font(size=9)
        words_cell.alignment = left_alignment
        row_idx += 1

    row_idx += 2
    clauses = [
        ('4. Послуги надані вчасно, в повному обсязі та належної якості.'),
        ('5. Сторони претензій одна до одної не мають.'),
        ('6. Цей акт складено в двох примірниках, по одному для кожної із Сторін.'),
    ]
    for clause in clauses:
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=5)
        ws.cell(row=row_idx, column=1, value=clause).font = Font(size=9)
        ws.cell(row=row_idx, column=1).alignment = left_alignment
        row_idx += 1

    row_idx += 2
    ws.cell(row=row_idx, column=1, value='ВИКОНАВЕЦЬ:').font = Font(bold=True, size=10)
    ws.cell(row=row_idx, column=4, value='ЗАМОВНИК:').font = Font(bold=True, size=10)
    row_idx += 1

    customer_details = [
        'ТОВ «БАЙЕР»',
        'Юридична адреса:',
        'вул. Б. Хмельницького, буд. 20',
        'Код ЄДРПОУ 22911794',
        'IBAN UA903006140000026002500345476',
        'в АТ "ОЩАДБАНК"',
    ]

    max_lines = max(len(full_executor_requisites), len(customer_details))
    for i in range(max_lines):
        curr_row = row_idx + i
        if i < len(full_executor_requisites):
            ws.cell(row=curr_row, column=1, value=full_executor_requisites[i]).font = Font(size=9)
        if i < len(customer_details):
            ws.cell(row=curr_row, column=4, value=customer_details[i]).font = Font(size=9)

    ws.column_dimensions['A'].width = 55  # Increased for long service names
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 25


# ============================================================
# TMA ACT SHEET (АКТ_ТМА)
# ============================================================

def generate_tma_act_sheet(
        ws,
        df_req,
        tma_subtotal_no_vat,  # "Всього без ПДВ" from TMA sheet (services + auth)
        period_start,
        period_end,
        selected_pharmacy='Усі аптеки',
        act_number='1',
        contract_no=DEFAULT_CONTRACT_NO,
        contract_date=DEFAULT_TMA_CONTRACT_DATE,
        executor_name=DEFAULT_EXECUTOR_NAME,
        is_lic=True
):
    """Generate АКТ_ТМА sheet for ЛИК or ДД category."""
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000'),
    )
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right_alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    bold_font = Font(bold=True, size=11)
    normal_font = Font(size=11)
    bold_normal_font = Font(bold=True, size=11)

    # Category label
    act_type = "ЛІК" if is_lic else "ДД"

    # Default requisites
    full_executor_requisites = [
        executor_name,
        'Юридична адреса:',
        '69050, м. Запоріжжя',
        'вул. Складська, буд. 6',
        'Код ЄДРПОУ 01977334',
        'ІПН 019773308278, св. № 200049431',
        'Банківські реквізити:',
        'IBAN UA633204780000026007924921260',
        'в АБ "УКРГАЗБАНК"',
        'МФО 320478',
        'Ел. пошта: sales1@apteki.zp.ua',
        'Голова правління',
        '_____________________ / ________ /',
    ]

    customer_name = 'ТОВ «БАЙЄР»'

    # Extract data from requisites
    if df_req is not None and not df_req.empty:
        matched_row = None
        clean_selected = str(selected_pharmacy).strip()
        target_col = None
        for col_candidate in ['chain_name'] + CLIENT_COL_VARIANTS + ['Мережа', 'Назва']:
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

        # Helper to validate values aren't client identifiers
        def _is_valid_entity(val):
            if val is None:
                return False
            if isinstance(val, float) and pd.isna(val):
                return False
            s = str(val).strip()
            if not s:
                return False
            s_lower = s.lower()
            client_keywords = ['клиент', 'база', 'customer', 'client', 'покупець', 'замовник']
            if any(kw in s_lower for kw in client_keywords):
                return False
            return True

        # Виконавець — використовуємо вибрану аптеку/мережу, якщо вибрано конкретну
        if clean_selected and clean_selected != 'Усі аптеки':
            executor_name = clean_selected
        else:
            # Якщо "Усі аптеки" — беремо з реквізитів (Виконавець або Юридическое лицо)
            # But reject client-like values
            executor_val = matched_row.get('Виконавець')
            if not _is_valid_entity(executor_val):
                executor_val = matched_row.get('Юридическое лицо')
            if not _is_valid_entity(executor_val):
                executor_val = executor_name  # DEFAULT_EXECUTOR_NAME
            executor_name = str(executor_val)
        # Замовник — пробуємо різні варіанти назв колонки
        for col in CLIENT_COL_VARIANTS:
            if col in matched_row.index and pd.notna(matched_row[col]) and str(matched_row[col]).strip():
                customer_name = str(matched_row[col]).strip()
                break
        contract_no = str(matched_row.get('№ договора', contract_no))
        contract_date = str(matched_row.get('Дата договора', contract_date))

        # Реквізити виконавця
        raw_req_text = str(matched_row.get('Реквизити', ''))
        if raw_req_text and raw_req_text.lower() != 'nan':
            full_executor_requisites = [
                line.strip() for line in raw_req_text.split('\n') if line.strip()
            ]

    # Legal entity for "між ТОВ «БАЙЄР» та {юридическое лицо}" - use 'Юридическое лицо' from requisites
    legal_entity_val = matched_row.get('Юридическое лицо') if df_req is not None and not df_req.empty else None
    if not _is_valid_entity(legal_entity_val):
        legal_entity_val = executor_name
    if not _is_valid_entity(legal_entity_val):
        legal_entity_val = clean_selected if clean_selected and clean_selected != 'Усі аптеки' else DEFAULT_EXECUTOR_NAME
    network_name = str(legal_entity_val)

    # Period formatting (same as other sheets)
    month_names = [
        'січня', 'лютого', 'березня', 'квітня', 'травня', 'червня',
        'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня'
    ]
    try:
        end_dt = datetime.strptime(period_end, '%Y-%m-%d')
        if end_dt.month == 12:
            next_month = datetime(end_dt.year + 1, 1, 1)
        else:
            next_month = datetime(end_dt.year, end_dt.month + 1, 1)
        end_last_day = next_month - timedelta(days=1)
        period_end_ukrainian = f'{end_last_day.day:02d} {month_names[end_last_day.month - 1]} {end_last_day.year} року'
        period_display = f'з «{period_start[8:10]}» {month_names[int(period_start[5:7]) - 1]} по «{end_last_day.day:02d}» {month_names[end_last_day.month - 1]} {end_last_day.year} року'
    except Exception:
        period_end_ukrainian = period_end
        period_display = period_start

    # ===== Row 1: "АКТ ТМА {ЛИК/ДД} № _______" =====
    ws.merge_cells('A1:E1')
    ws['A1'] = f'АКТ ТМА {act_type} № _______'
    ws['A1'].font = bold_font
    ws['A1'].alignment = center_alignment

    # ===== Row 2: "про надані послуги відповідно до Договору про маркетингові послуги {№ договора} від {дата}" =====
    ws.merge_cells('A2:E2')
    ws['A2'] = f'про надані послуги відповідно до Договору про маркетингові послуги № {contract_no} від {contract_date}'
    ws['A2'].font = normal_font
    ws['A2'].alignment = center_alignment
    ws.row_dimensions[2].height = 30

    # ===== Row 3: "між ТОВ «БАЙЄР» та {мережа}" =====
    ws.merge_cells('A3:E3')
    ws['A3'] = f'між ТОВ «БАЙЄР» та {network_name}'
    ws['A3'].font = normal_font
    ws['A3'].alignment = center_alignment

    # ===== Row 4: "м. Київ" + date =====
    ws['A4'] = 'м. Київ'
    ws['A4'].font = normal_font
    ws['A4'].alignment = left_alignment
    ws.merge_cells('D4:E4')
    ws['D4'] = f'"{period_end_ukrainian[:2]}" {period_end_ukrainian[3:]}'
    ws['D4'].font = normal_font
    ws['D4'].alignment = right_alignment

    # ===== Row 6: Виконавець =====
    ws.merge_cells('A6:E6')
    ws['A6'] = executor_name
    ws['A6'].font = bold_normal_font
    ws['A6'].alignment = left_alignment
    ws.row_dimensions[6].height = 25

    # ===== Row 7: Hardcoded intro text =====
    ws.merge_cells('A7:E7')
    intro_text = (
        'Виконавець надав ТОВ «БАЙЄР» (Замовнику) маркетингові послуги '
        'шляхом проведення трейд-маркетингових заходів (активностей), а саме:'
    )
    ws['A7'] = intro_text
    ws['A7'].font = normal_font
    ws['A7'].alignment = left_alignment
    ws.row_dimensions[7].height = 40

    # ===== Row 8: Services description (1.1, 1.2, medicine note) - multi-line in one cell =====
    ws.merge_cells('A8:E8')
    medicine_text = 'який є лікарським засобом' if is_lic else 'який не є лікарським засобом'
    services_text = (
        f'1.1. проведення акції зі зниження вартості щодо Товару Замовника для кінцевих споживачів '
        f'у аптечних закладах Аптечної мережі, що спрямовані на стимулювання збуту та просування Товарів Замовника;\n'
        f'1.2. послуги із авторизації, направлені на підготовку Товару до проведення Акції;\n\n'
        f'Послуги надані стосовно Товару Замовника, {medicine_text}.'
    )
    ws['A8'] = services_text
    ws['A8'].font = normal_font
    ws['A8'].alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
    ws.row_dimensions[8].height = 80

    # ===== Row 9: Empty spacing row =====
    ws.row_dimensions[9].height = 15

    # ===== Row 10: "2. Маркетингові послуги були надані у період {період}" =====
    ws.merge_cells('A10:E10')
    ws['A10'] = f'2. Маркетингові послуги були надані у період {period_display}'
    ws['A10'].font = normal_font
    ws['A10'].alignment = left_alignment

    # ===== Row 11: Empty spacing row =====
    ws.row_dimensions[11].height = 15

    # ===== Rows 12-14: Amounts in words =====
    from num2words import num2words

    def format_in_words(amount):
        uah_int = int(amount)
        kop = int(round((amount - uah_int) * 100))
        uah_text = num2words(uah_int, lang='uk')
        return f'{uah_text} грн. {kop:02d} коп.'

    # Calculate amounts (tma_subtotal_no_vat already includes services + auth)
    total_no_vat = float(tma_subtotal_no_vat)
    vat_amount = total_no_vat * 0.2
    total_with_vat = total_no_vat + vat_amount

    # Row 12: "3. Вартість наданих послуг без ПДВ складає" + amount in words
    ws.merge_cells('A12:E12')
    ws['A12'] = f'3. Вартість наданих послуг без ПДВ складає — {format_in_words(total_no_vat)}'
    ws['A12'].font = bold_normal_font
    ws['A12'].alignment = left_alignment

    # Row 13: "Вартість наданих послуг з ПДВ складає" + amount in words
    ws.merge_cells('A13:E13')
    ws['A13'] = f'Вартість наданих послуг з ПДВ складає — {format_in_words(total_with_vat)}'
    ws['A13'].font = bold_normal_font
    ws['A13'].alignment = left_alignment

    # Row 14: "ПДВ 20%" + amount in words
    ws.merge_cells('A14:E14')
    ws['A14'] = f'ПДВ 20% — {format_in_words(vat_amount)}'
    ws['A14'].font = bold_normal_font
    ws['A14'].alignment = left_alignment

    # ===== Row 15: Empty spacing row =====
    ws.row_dimensions[15].height = 15

    # ===== Row 16: Clause 4 =====
    ws.merge_cells('A16:E16')
    ws['A16'] = '4. Усі терміни вживаються в цьому Акті в тому ж значенні, у якому вони вживаються в Договорі.'
    ws['A16'].font = normal_font
    ws['A16'].alignment = left_alignment

    # ===== Row 17: Clause 5 =====
    ws.merge_cells('A17:E17')
    ws['A17'] = '5. Цей Акт складений у двох ідентичних примірниках, по одному для кожної Сторони.'
    ws['A17'].font = normal_font
    ws['A17'].alignment = left_alignment

    # ===== Row 18: Clause 6 =====
    ws.merge_cells('A18:E18')
    ws['A18'] = (
        '6. Замовник підтверджує, що послуги, зазначені у цьому Акті, були надані Виконавцем в повному обсязі, '
        'зауваження щодо якості та своєчасності надання послуг на момент підписання Акту відсутні.'
    )
    ws['A18'].font = normal_font
    ws['A18'].alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
    ws.row_dimensions[18].height = 40

    # ===== Requisites section (same as other acts) =====
    row_idx = 20
    ws.cell(row=row_idx, column=1, value='ВИКОНАВЕЦЬ:').font = Font(bold=True, size=11)
    ws.cell(row=row_idx, column=4, value='ЗАМОВНИК:').font = Font(bold=True, size=11)
    row_idx += 1

    customer_details = [
        'ТОВ «БАЙЄР»',
        'Юридична адреса:',
        'вул. Б. Хмельницького, буд. 20',
        'Код ЄДРПОУ 22911794',
        'IBAN UA903006140000026002500345476',
        'в АТ "ОЩАДБАНК"',
    ]

    max_lines = max(len(full_executor_requisites), len(customer_details))
    for i in range(max_lines):
        curr_row = row_idx + i
        if i < len(full_executor_requisites):
            ws.cell(row=curr_row, column=1, value=full_executor_requisites[i]).font = Font(size=10)
        if i < len(customer_details):
            ws.cell(row=curr_row, column=4, value=customer_details[i]).font = Font(size=10)

    # Column widths
    ws.column_dimensions['A'].width = 60
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 25


# Need to import datetime and timedelta at the top
# Added at module level