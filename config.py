"""Configuration constants for Bayer Marketing Report Generator."""

# Contract defaults (can be overridden by requisites file)
DEFAULT_CONTRACT_NO = '260326'
DEFAULT_CONTRACT_DATE = '01 березня 2026 року'
DEFAULT_TMA_CONTRACT_DATE = '27 лютого 2026 року'
DEFAULT_EXECUTOR_NAME = 'ПрАТ «Аптеки Запоріжжя»'

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

# MHL reward multiplier (percentage of service2_total)
# 0.1 = 10%
MHL_REWARD_MULTIPLIER = 0.1

# Database connection
DB_CONNECTION_STRING = (
    r'DRIVER={SQL Server Native Client 11.0};'
    r'SERVER=BYDE276YZ8D.bayer.cnb\I0940,58321;'
    r'DATABASE=BAYER_CH_UA;'
    r'Trusted_Connection=yes;'
)