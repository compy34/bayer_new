from num2words import num2words
from openpyxl.styles import Alignment, Border, Font, Side


def generate_act_sheet(
        ws,
        df_req,
        total_sum_without_vat,
        period_start,
        period_end,
        selected_pharmacy='Усі аптеки',
        act_number='1',
        contract_no='260326',
        contract_date='01 березня 2026 року',
        executor_name='ПрАТ «Аптеки Запоріжжя»'
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

        # БЕРЕМО ГОТОВІ РЕКВІЗИТИ З КОЛОНКИ ФАЙЛУ
        raw_req_text = str(matched_row.get('Реквизити', ''))
        if raw_req_text and raw_req_text.lower() != 'nan':
            full_executor_requisites = [
                line.strip() for line in raw_req_text.split('\n') if line.strip()
            ]

    customer_name = 'ТОВ «БАЙЕР»'

    ws.merge_cells('A1:E1')
    ws['A1'] = f'АКТ № {act_number}'
    ws['A1'].font = Font(bold=True, size=11)
    ws['A1'].alignment = center_alignment

    ws.merge_cells('A2:E2')
    ws['A2'] = (
        f'надання послуг (виконання робіт)\n'
        f'до договору № {contract_no} від {contract_date}\n'
        f'Замовник: {customer_name} | Виконавець: {executor_name}'
    )
    ws['A2'].font = Font(bold=True, size=10)
    ws['A2'].alignment = center_alignment
    ws.row_dimensions[2].height = 40

    ws['A4'] = 'м. Київ'
    ws['A4'].font = Font(size=10)
    ws['A4'].alignment = left_alignment

    ws.merge_cells('D4:E4')
    ws['D4'] = f'"{period_end[:2]}" {period_end[3:]}'
    ws['D4'].font = Font(size=10)
    ws['D4'].alignment = right_alignment

    ws.merge_cells('A6:E6')
    ws['A6'] = (
        f'{executor_name} виконав, а {customer_name} прийняв наступні послуги '
        f'згідно з договором № {contract_no} від {contract_date}.'
    )
    ws['A6'].font = Font(size=10)
    ws['A6'].alignment = left_alignment
    ws.row_dimensions[6].height = 35

    ws['A8'] = (
        f'1. Період надання послуг: з "{period_start}" по "{period_end}".'
    )
    ws['A8'].font = Font(size=10)
    ws['A9'] = '2. Якість послуг відповідає умовам договору.'
    ws['A9'].font = Font(size=10)
    ws['A11'] = '3. Перелік наданих послуг:'
    ws['A11'].font = Font(bold=True, size=10)

    headers = [
        'Найменування послуг',
        'Од.',
        'Сума без ПДВ, грн.',
        'ПДВ 20%',
        'Загальна сума з ПДВ, грн.',
    ]
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=13, column=col_idx, value=h_text)
        cell.font = Font(bold=True, size=9)
        cell.alignment = center_alignment
        cell.border = thin_border
    ws.row_dimensions[13].height = 25

    sum1_no_vat = float(total_sum_without_vat)
    vat1 = sum1_no_vat * 0.2
    total1 = sum1_no_vat + vat1

    sum2_no_vat = 0.00
    vat2 = sum2_no_vat * 0.2
    total2 = sum2_no_vat + vat2

    services_data = [
        (
            'Маркетингові та рекламні послуги (мерчандайзинг)',
            'послуга',
            sum1_no_vat,
            vat1,
            total1,
        ),
    ]

    row_idx = 14
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
        ws.row_dimensions[row_idx].height = 30
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

    ws.column_dimensions['A'].width = 42
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 25