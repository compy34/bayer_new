import pyodbc
import pandas as pd
from pathlib import Path

from config import DB_CONNECTION_STRING

PROJECT_ROOT = Path(__file__).resolve().parent.parent


class DataLoadError(Exception):
    """Exception raised when data loading fails."""
    pass


def _get_connection():
    """Create and return a database connection."""
    return pyodbc.connect(DB_CONNECTION_STRING)


def create_contract_skus_table():
    """Create table for 'для договорів' SKU data if not exists.

    Returns True on success, False on failure (e.g., DB not available).
    """
    try:
        conn = _get_connection()
    except Exception:
        return False
    try:
        cursor = conn.cursor()
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='contract_skus' AND xtype='U')
            CREATE TABLE contract_skus (
                id INT IDENTITY(1,1) PRIMARY KEY,
                sku_ukr NVARCHAR(255) NOT NULL,
                category NVARCHAR(10) NOT NULL CHECK (category IN ('ЛІК', 'ДД')),
                brand NVARCHAR(255),
                created_at DATETIME DEFAULT GETDATE(),
                updated_at DATETIME DEFAULT GETDATE()
            )
        """)
        conn.commit()
        return True
    except Exception:
        return False
    finally:
        conn.close()


# Table creation is lazy - call create_contract_skus_table() when needed
# create_contract_skus_table()  # Uncomment to create on import


def _apply_chain_filter(query: str, params: list, chain_name: str, column: str = 't.CHAIN_NAME') -> tuple:
    """Apply chain filter to query if chain_name is provided."""
    if chain_name and chain_name.strip() and chain_name.strip() != 'Усі аптеки':
        clean_chain = chain_name.strip()
        query += f" AND LTRIM(RTRIM(LOWER({column}))) = LOWER(?)"
        params.append(clean_chain)
    return query, params


def get_pharmacy_data(period_start, period_end, chain_name=None):
    """Load pharmacy merchandising data from calc_Merch_KA.

    Args:
        period_start: Start date in YYYY-MM-DD format
        period_end: End date in YYYY-MM-DD format
        chain_name: Optional pharmacy chain/network name to filter by

    Returns:
        DataFrame with columns: CHAIN_NAME, PHARMACY_NAME, BRAND, MERCH, Price

    Raises:
        DataLoadError: If database query fails
    """
    connection = None
    try:
        connection = _get_connection()

        query = """
            SELECT
                t.CHAIN_NAME,
                t.PHARMACY_NAME,
                t.BRAND,
                t.MERCH,
                COALESCE(p.Price, 0) AS Price
            FROM calc_Merch_KA t
            LEFT JOIN LA_Price p ON t.BRAND = p.Brand
            WHERE YEAR(t.PERIOD) * 100 + MONTH(t.PERIOD) BETWEEN YEAR(?) * 100 + MONTH(?) AND YEAR(?) * 100 + MONTH(?)
        """

        params = [period_start, period_start, period_end, period_end]
        query, params = _apply_chain_filter(query, params, chain_name)

        df = pd.read_sql(query, connection, params=params)
        return df

    except pyodbc.Error as e:
        raise DataLoadError(f'Помилка завантаження даних аптек (Merch): {e}') from e

    finally:
        if connection is not None:
            connection.close()


def get_mhl_data(period_start, period_end, chain_name=None):
    """Load MHL metrics and the latest penetration target for each product.

    Args:
        period_start: Start date in YYYY-MM-DD format
        period_end: End date in YYYY-MM-DD format
        chain_name: Optional pharmacy chain/network name to filter by

    Returns:
        DataFrame with columns: DATE, CHAIN_ID, CHAIN_NAME, DRUG_ID, SKU_UA, BRAND,
        PENETRATION, PRICE_CP, SUPPLY, PENETRATION_TARGET

    Raises:
        DataLoadError: If database query fails
    """
    connection = None
    try:
        connection = _get_connection()

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
        query, params = _apply_chain_filter(query, params, chain_name, 'main.[CHAIN_NAME]')

        df = pd.read_sql(query, connection, params=params)
        return df

    except pyodbc.Error as e:
        raise DataLoadError(f'Помилка завантаження MHL даних: {e}') from e

    finally:
        if connection is not None:
            connection.close()


def get_tma_offtake_data(period_start, period_end, chain_name=None):
    """Load TMA offtake (sold quantity) data from dbo.supply_KA_TMA.

    Args:
        period_start: Start date in YYYY-MM-DD format
        period_end: End date in YYYY-MM-DD format
        chain_name: Optional pharmacy chain/network name to filter by

    Returns:
        DataFrame with columns: DATE, CHAIN_ID, CHAIN_NAME, DRUG_ID, OFFTAKE

    Raises:
        DataLoadError: If database query fails
    """
    connection = None
    try:
        connection = _get_connection()

        query = """
            SELECT
                t.[DATE],
                t.[CHAIN_ID],
                t.[CHAIN_NAME],
                t.[DRUG_ID],
                t.[OFFTAKE]
            FROM [dbo].[supply_KA_TMA] AS t
            WHERE t.[DATE] >= ? AND t.[DATE] <= ?
        """

        params = [period_start, period_end]
        query, params = _apply_chain_filter(query, params, chain_name, 't.[CHAIN_NAME]')

        df = pd.read_sql(query, connection, params=params)
        return df

    except pyodbc.Error as e:
        raise DataLoadError(f'Помилка завантаження TMA offtake даних: {e}') from e

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


def get_all_active_skus(period_start=None, period_end=None):
    """Get all unique SKUs with their brands from the database that have recent activity.

    Args:
        period_start: Optional start date (YYYY-MM-DD). If not provided, uses 12 months before max date in calc_KA.
        period_end: Optional end date (YYYY-MM-DD). If not provided, uses max date in calc_KA.

    Returns:
        DataFrame with columns: SKU_UA, BRAND
        Contains unique product names (SKU_UA) from key_SKU table
        with their corresponding brand from calc_KA, filtered to only SKUs
        that have data in calc_KA within the specified period (active products).
    """
    connection = None
    try:
        connection = _get_connection()

        # Build period filter subquery
        period_filter = ""
        params = []
        if period_start and period_end:
            period_filter = """
                AND EXISTS (
                    SELECT 1 FROM [dbo].[calc_KA] ka2
                    WHERE ka2.[DRUG_ID] = sku.[DRUG_ID]
                      AND ka2.[DATE] >= ? AND ka2.[DATE] <= ?
                )
            """
            params.extend([period_start, period_end])
        else:
            # Default: last 12 months from max date in calc_KA
            period_filter = """
                AND EXISTS (
                    SELECT 1 FROM [dbo].[calc_KA] ka2
                    WHERE ka2.[DRUG_ID] = sku.[DRUG_ID]
                      AND ka2.[DATE] >= DATEADD(MONTH, -12, (SELECT MAX([DATE]) FROM [dbo].[calc_KA]))
                )
            """

        # Get all unique SKU-Brand pairs from the database
        # key_SKU has SKU_UA and DRUG_ID, calc_KA has DRUG_ID and BRAND
        query = f"""
            SELECT DISTINCT
                sku.[SKU_UA],
                ka.[BRAND]
            FROM [dbo].[key_SKU] AS sku
            INNER JOIN [dbo].[calc_KA] AS ka
                ON ka.[DRUG_ID] = sku.[DRUG_ID]
            WHERE sku.[SKU_UA] IS NOT NULL
              AND sku.[SKU_UA] <> ''
              AND ka.[BRAND] IS NOT NULL
              AND ka.[BRAND] <> ''
              {period_filter}
        """

        df = pd.read_sql(query, connection, params=params)
        # Clean up
        df['SKU_UA'] = df['SKU_UA'].astype(str).str.strip()
        df['BRAND'] = df['BRAND'].astype(str).str.strip()
        df = df[(df['SKU_UA'] != '') & (df['BRAND'] != '')]
        return df.drop_duplicates().sort_values(['BRAND', 'SKU_UA']).reset_index(drop=True)

    except pyodbc.Error as e:
        raise DataLoadError(f'Помилка завантаження списку продуктів: {e}') from e

    finally:
        if connection is not None:
            connection.close()


# ============================================================
# CONTRACT SKUS TABLE (for "для договорів" data)
# ============================================================

def save_contract_skus(skus_data):
    """Save SKUs from contract file to database.

    Args:
        skus_data: List of dicts with keys: sku_ukr, category, brand

    Returns:
        (success_count, error_message) tuple
    """
    if not skus_data:
        return 0, "No data to save"

    conn = None
    try:
        conn = _get_connection()
        cursor = conn.cursor()

        saved = 0
        for item in skus_data:
            sku_ukr = str(item.get('sku_ukr', '')).strip()
            category = str(item.get('category', '')).strip().upper()
            brand = str(item.get('brand', '')).strip()

            if not sku_ukr or category not in ('ЛІК', 'ДД'):
                continue

            # Upsert: insert or update
            cursor.execute("""
                MERGE contract_skus AS target
                USING (SELECT ? AS sku_ukr) AS source
                ON target.sku_ukr = source.sku_ukr
                WHEN MATCHED THEN
                    UPDATE SET category = ?, brand = ?, updated_at = GETDATE()
                WHEN NOT MATCHED THEN
                    INSERT (sku_ukr, category, brand) VALUES (?, ?, ?);
            """, sku_ukr, category, brand, sku_ukr, category, brand)
            saved += 1

        conn.commit()
        return saved, None

    except pyodbc.Error as e:
        error_msg = f'Помилка збереження в БД: {e}'
        print(error_msg)
        return 0, error_msg
    finally:
        if conn is not None:
            conn.close()


def load_contract_skus():
    """Load all SKUs from contract_skus table.

    Returns:
        DataFrame with columns: sku_ukr, category, brand
    """
    conn = None
    try:
        conn = _get_connection()
        query = """
            SELECT sku_ukr, category, brand
            FROM contract_skus
            ORDER BY category, sku_ukr
        """
        df = pd.read_sql(query, conn)
        return df
    except pyodbc.Error as e:
        print(f'Помилка завантаження contract_skus: {e}')
        return pd.DataFrame(columns=['sku_ukr', 'category', 'brand'])
    finally:
        if conn is not None:
            conn.close()


def delete_contract_sku(sku_ukr):
    """Delete a SKU from contract_skus table."""
    conn = None
    try:
        conn = _get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM contract_skus WHERE sku_ukr = ?", sku_ukr)
        conn.commit()
        return cursor.rowcount > 0
    except pyodbc.Error as e:
        print(f'Помилка видалення: {e}')
        return False
    finally:
        if conn is not None:
            conn.close()