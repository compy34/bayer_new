import pandas as pd
from convertor import generate_excel_report

# Load test data
df_req = pd.read_excel('/home/user/bayer_new/test_data_output/test_requisites.xlsx')
df_pharm = pd.read_excel('/home/user/bayer_new/test_data_output/test_pharmacy_data.xlsx')
df_mhl = pd.read_excel('/home/user/bayer_new/test_data_output/test_mhl_data.xlsx')

# Filter for first chain
chain = 'Тестова Мережа'
df_pharm_filtered = df_pharm[df_pharm['CHAIN_NAME'] == chain]
df_mhl_filtered = df_mhl[df_mhl['CHAIN_NAME'] == chain]

print(f"Pharmacy data: {len(df_pharm_filtered)} rows")
print(f"MHL data: {len(df_mhl_filtered)} rows")

# Generate report
generate_excel_report(
    df_pharm_filtered,
    df_req,
    '/home/user/bayer_new/test_output_report.xlsx',
    period_start='2026-08-01',
    period_end='2026-08-31',
    selected_pharmacy=chain,
    mhl_df=df_mhl_filtered
)

print("Report generated successfully!")

# Verify sheets
import openpyxl
wb = openpyxl.load_workbook('/home/user/bayer_new/test_output_report.xlsx')
print(f"Sheets: {wb.sheetnames}")
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    print(f"  {sheet_name}: {ws.max_row} rows x {ws.max_column} cols")
    # Print first few rows of TMA sheets
    if 'під час акції' in sheet_name:
        print(f"    First 5 rows:")
        for row in ws.iter_rows(min_row=1, max_row=min(5, ws.max_row), values_only=False):
            vals = [(cell.value, cell.coordinate) for cell in row if cell.value is not None]
            if vals:
                print(f"      {vals}")
