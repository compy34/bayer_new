import pyodbc
import pandas as pd
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def get_pharmacy_data(period_start, period_end):
    try:
        # Налаштування підключення до MS SQL Server через Windows Auth та Native Client 11.0
        conn_str = (
            r'DRIVER={SQL Server Native Client 11.0};'
            r'SERVER=BYDE276YZ8D.bayer.cnb\I0940,58321;'
            r'DATABASE=BAYER_CH_UA;'
            r'Trusted_Connection=yes;'
        )
        
        connection = pyodbc.connect(conn_str)
        
        # У pyodbc для параметрів використовується '?' замість '%s'
        query = """
            SELECT 
                t.CHAIN_NAME,
                t.PHARMACY_NAME, 
                t.BRAND, 
                t.MERCH,
                COALESCE(p.Price, 0) AS Price
            FROM calc_Merch_KA t
            LEFT JOIN LA_Price p ON t.BRAND = p.Brand
            WHERE t.PERIOD >= ? AND t.PERIOD <= ?
        """
        # Передаємо параметри у вигляді кортежу
        df = pd.read_sql(query, connection, params=(period_start, period_end))
        return df
        
    except pyodbc.Error as e:
        print(f'Помилка MS SQL Server: {e}')
        return None
        
    finally:
        if 'connection' in locals():
            connection.close()


def get_mhl_data(period_start, period_end, chain_name=None):
    """Load MHL metrics and the latest penetration target for each product.

    Args:
        period_start: Start date in YYYY-MM-DD format
        period_end: End date in YYYY-MM-DD format
        chain_name: Optional pharmacy chain/network name to filter by (e.g., from UI selection)
    """
    connection = None
    try:
        conn_str = (
            r'DRIVER={SQL Server Native Client 11.0};'
            r'SERVER=BYDE276YZ8D.bayer.cnb\I0940,58321;'
            r'DATABASE=BAYER_CH_UA;'
            r'Trusted_Connection=yes;'
        )
        connection = pyodbc.connect(conn_str)

        # Base query
        query = """
            SELECT
                main.[DATE],
                main.[CHAIN_ID],
                main.[CHAIN_NAME],
                main.[DRUG_ID],
                sku.[SKU_UA],
                main.[BRAND],
                main.[PENETRATION],
                main.[PRICE_CP],
                main.[SUPPLY],
                target.[PENETRATION_TARGET]
            FROM [dbo].[calc_KA] AS main
            LEFT JOIN [dbo].[key_SKU] AS sku
                ON sku.[DRUG_ID] = main.[DRUG_ID]
            OUTER APPLY (
                SELECT TOP 1
                    [plan].[PENETRATION_TARGET]
                FROM [dbo].[KA_Penetration_plan] AS [plan]
                WHERE [plan].[DRUG_ID] = main.[DRUG_ID]
                  AND [plan].[PERIOD] <= main.[DATE]
                ORDER BY [plan].[PERIOD] DESC
            ) AS target
            WHERE main.[DATE] >= ? AND main.[DATE] <= ?
        """

        params = [period_start, period_end]

        # Add chain filter if provided
        if chain_name and chain_name.strip() and chain_name.strip() != 'Усі аптеки':
            query += " AND main.[CHAIN_NAME] = ?"
            params.append(chain_name.strip())

        return pd.read_sql(query, connection, params=params)
    except pyodbc.Error as error:
        raise OSError(f'Помилка завантаження MHL з MS SQL Server: {error}') from error
    finally:
        if connection is not None:
            connection.close()


def get_requisites_data():
    return load_requisites_file(PROJECT_ROOT / 'rekvizity.xlsx')


def load_requisites_file(filename):
    """Load requisites from a user-selected CSV or Excel file."""
    path = Path(filename)
    if path.suffix.lower() == '.csv':
        df_req = pd.read_csv(path, encoding='utf-8', dtype=str)
    elif path.suffix.lower() in {'.xlsx', '.xls'}:
        df_req = pd.read_excel(path, dtype=str)
    else:
        raise ValueError('Підтримуються тільки файли CSV, XLSX або XLS.')

    # Очищуємо заголовки колонок від зайвих пробілів
    df_req.columns = df_req.columns.astype(str).str.strip()
    return df_req.fillna('')