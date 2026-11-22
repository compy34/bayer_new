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


def get_all_active_skus():
    """Get all unique SKUs with their brands from the database.

    Returns:
        DataFrame with columns: SKU_UA, BRAND
        Contains all unique product names (SKU_UA) from key_SKU table
        with their corresponding brand from calc_KA.
    """
    connection = None
    try:
        connection = _get_connection()

        # Get all unique SKU-Brand pairs from the database
        # key_SKU has SKU_UA and DRUG_ID, calc_KA has DRUG_ID and BRAND
        query = """
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
        """

        df = pd.read_sql(query, connection)
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