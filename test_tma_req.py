import pandas as pd
from convertor import generate_excel_report

# Load test data
df_req = pd.read_excel('/home/user/bayer_new/тест_ТМА_РЕК_1.xlsx')
df_pharm = pd.read_excel('/home/user/bayer_new/test_data_output/test_pharmacy_data.xlsx')
df_mhl = pd.read_excel('/home/user/bayer_new/test_data_output/test_mhl_data.xlsx')

# Filter for first chain
chain = 'Тестова Мережа'
df_pharm_filtered = df_pharm[df_pharm['CHAIN_NAME'] == chain]
df_mhl_filtered = df_mhl[df_mhl['CHAIN_NAME'] == chain]

print(f"Pharmacy data: {len(df_pharm_filtered)} rows")
print(f"MHL data: {len(df_mhl_filtered)} rows")
print(f"Requisites columns: {list(df_req.columns)}")

# Generate report
generate_excel_report(
    df_pharm_filtered,
    df_req,
    '/home/user/bayer_new/test_output_tma_req.xlsx',
    period_start='2026-08-01',
    period_end='2026-08-31',
    selected_pharmacy=chain,
    mhl_df=df_mhl_filtered
)

print("Report generated successfully!")

# Verify TMA sheets signature section
import openpyxl
wb = openpyxl.load_workbook('/home/user/bayer_new/test_output_tma_req.xlsx')
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    if 'під час акції' in sheet_name:
        print(f"\n{'='*60}")
        print(f"Sheet: {sheet_name}")
        print(f"{'='*60}")
        # Print last 8 rows (signature section)
        for row in ws.iter_rows(min_row=max(1, ws.max_row-7), max_row=ws.max_row, values_only=False):
            vals = [(cell.value, cell.coordinate) for cell in row if cell.value is not None]
            if vals:
                print(f"  {vals}")
