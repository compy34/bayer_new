import openpyxl

wb = openpyxl.load_workbook('/home/user/bayer_new/test_output_report.xlsx')

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    if 'під час акції' in sheet_name:
        print(f"\n{'='*80}")
        print(f"Sheet: {sheet_name} ({ws.max_row} rows x {ws.max_column} cols)")
        print(f"{'='*80}")
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=False):
            vals = [(cell.value, cell.coordinate) for cell in row if cell.value is not None]
            if vals:
                print(f"  {vals}")
