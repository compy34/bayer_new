#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test data generator for Bayer marketing reports.
Creates mock pharmacy data, MHL data, and requisites for testing.
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta


# ============================================================
# TEST REQUISITES DATA
# ============================================================

def create_test_requisites():
    """Create test requisites CSV file."""
    data = {
        'chain_name': ['Тестова Мережа', 'Інша Мережа', 'Третя Мережа'],
        'Клиент из базы данных': ['Тестова Мережа', 'Інша Мережа', 'Третя Мережа'],
        'Мережа': ['Тестова Мережа', 'Інша Мережа', 'Третя Мережа'],
        'Назва': ['Тестова Мережа', 'Інша Мережа', 'Третя Мережа'],
        'Юридическое лицо': [
            'ТОВ "Тестова Компанія"',
            'ТОВ "Інша Компанія"',
            'ТОВ "Третя Компанія"'
        ],
        'Виконавець': [
            'ТОВ "Тестова Компанія"',
            'ТОВ "Інша Компанія"',
            'ТОВ "Третя Компанія"'
        ],
        '№ договора': ['260326', '260327', '260328'],
        'Дата договора': ['01 березня 2026 року', '15 березня 2026 року', '01 квітня 2026 року'],
        'Реквизити': [
            'ТОВ "Тестова Компанія"\nЮридична адреса:\n01001, м. Київ\nвул. Тестова, 1\nКод ЄДРПОУ 12345678\nІПН 123456789012\nБанківські реквізити:\nIBAN UA123456789012345678901234567\nв АТ "ТестБанк"\nМФО 123456',
            'ТОВ "Інша Компанія"\nЮридична адреса:\n02002, м. Львів\nвул. Інша, 2\nКод ЄДРПОУ 87654321\nІПН 210987654321\nБанківські реквізити:\nIBAN UA987654321098765432109876543\nв АТ "ІншийБанк"\nМФО 654321',
            'ТОВ "Третя Компанія"\nЮридична адреса:\n03003, м. Одеса\nвул. Третя, 3\nКод ЄДРПОУ 11223344\nІПН 443322110099\nБанківські реквізити:\nIBAN UA112233445566778899001122334\nв АТ "ТретійБанк"\nМФО 112233'
        ],
        'Директор': ['Іванов І.І.', 'Петров П.П.', 'Сидоров С.С.'],
    }
    df = pd.DataFrame(data)
    return df


# ============================================================
# TEST PHARMACY DATA (calc_Merch_KA)
# ============================================================

def create_test_pharmacy_data(period_start='2026-08-01', period_end='2026-08-31'):
    """Create test merchandising/planogram data."""
    # All template brands for ЛИК and ДД
    lik_brands = ['Антифлу', 'Бепантен', 'Бепантен Плюс', 'Елевіт-3', 'Еріус', 'Реліф', 'Супрадин-10', 'Ренні', 'Терафлекс + Тералів']
    dd_brands = ['Елевіт', 'Супрадин-13', 'Супрадин-10', 'Бепантен Сенсідерм-4', 'Бепантен Сенсідерм-2']

    chains = ['Тестова Мережа', 'Інша Мережа', 'Третя Мережа']

    # Pharmacy names per chain
    pharmacies = {
        'Тестова Мережа': ['Аптека №1 (Київ, вул. Хрещатик 1)', 'Аптека №2 (Київ, вул./Base 2)', 'Аптека №3 (Київ, вул. Тестова 3)'],
        'Інша Мережа': ['Аптека \"Здоров\'я\" (Львів, вул. Свободи 10)', 'Аптека \"Аптека\" (Львів, вул. Франка 5)'],
        'Третя Мережа': ['Аптека №100 (Одеса, вул. Дерибасівська 1)', 'Аптека №101 (Одеса, вул. Пушкінська 2)', 'Аптека №102 (Одеса, вул. Рішельєвська 3)'],
    }

    rows = []
    for chain in chains:
        for pharmacy in pharmacies[chain]:
            # ЛИК brands - random merch counts
            for brand in lik_brands:
                merch = 0 if brand in ['Антифлу', 'Супрадин-10'] else np.random.randint(1, 15)
                rows.append({
                    'CHAIN_NAME': chain,
                    'PHARMACY_NAME': pharmacy,
                    'BRAND': brand,
                    'MERCH': merch,
                    'Price': 800 if brand in ['Бепантен', 'Реліф', 'Терафлекс + Тералів', 'Супрадин-10'] else
                             600 if brand in ['Еріус', 'Ренні', 'Елевіт-3'] else
                             300 if brand in ['Антифлу', 'Бепантен Плюс'] else 0
                })
            # ДД brands
            for brand in dd_brands:
                merch = 0 if brand in ['Бепантен Сенсідерм-2'] else np.random.randint(1, 10)
                rows.append({
                    'CHAIN_NAME': chain,
                    'PHARMACY_NAME': pharmacy,
                    'BRAND': brand,
                    'MERCH': merch,
                    'Price': 800
                })

    df = pd.DataFrame(rows)
    return df


# ============================================================
# TEST MHL DATA (calc_KA + key_SKU + KA_Penetration_plan)
# ============================================================

def create_test_mhl_data(period_start='2026-08-01', period_end='2026-08-31'):
    """Create test MHL (availability) data."""
    # Sample SKUs from sku_classification.csv
    skus_data = [
        # ЛИК SKUs
        {'sku': 'Еріус сироп 0.5 мг/мл 60мл', 'brand': 'Еріус', 'drug_id': 100, 'category': 'ЛІК'},
        {'sku': 'Еріус таблетки 5 мг № 10', 'brand': 'Еріус', 'drug_id': 101, 'category': 'ЛІК'},
        {'sku': 'Антифлу Кідс, порошок пакетики №5', 'brand': 'Антифлу', 'drug_id': 102, 'category': 'ЛІК'},
        {'sku': 'Бепантен, крем 5%, туба 30г №1', 'brand': 'Бепантен', 'drug_id': 103, 'category': 'ЛІК'},
        {'sku': 'Бепантен, мазь 5%, туба 30г №1', 'brand': 'Бепантен', 'drug_id': 104, 'category': 'ЛІК'},
        {'sku': 'Бепантен Плюс, крем, туба 30г №1', 'brand': 'Бепантен Плюс', 'drug_id': 105, 'category': 'ЛІК'},
        {'sku': 'Елевіт Пронаталь, т.в/о №100', 'brand': 'Елевіт-3', 'drug_id': 106, 'category': 'ЛІК'},
        {'sku': 'Реліф, супозиторїї ректальні №12', 'brand': 'Реліф', 'drug_id': 107, 'category': 'ЛІК'},
        {'sku': 'Ренні, табл. жув., мент. смак. №24', 'brand': 'Ренні', 'drug_id': 108, 'category': 'ЛІК'},
        {'sku': 'Терафлекс, капсули №60', 'brand': 'Терафлекс + Тералів', 'drug_id': 109, 'category': 'ЛІК'},
        # ДД SKUs
        {'sku': 'Супрадин Актив, табл. №30', 'brand': 'Супрадин-13', 'drug_id': 200, 'category': 'ДД'},
        {'sku': 'Супрадин Ведмежуйки №30', 'brand': 'Супрадин-10', 'drug_id': 201, 'category': 'ДД'},
        {'sku': 'Бепантен Сенсідерм', 'brand': 'Бепантен Сенсідерм-4', 'drug_id': 202, 'category': 'ДД'},
        {'sku': 'Елевіт-Комплекс 2, капс. № 30', 'brand': 'Елевіт', 'drug_id': 203, 'category': 'ДД'},
    ]

    chains = ['Тестова Мережа', 'Інша Мережа', 'Третя Мережа']

    rows = []
    for chain in chains:
        chain_id = chains.index(chain) + 1
        for sku_info in skus_data:
            # Generate random but realistic data
            penetration = round(np.random.uniform(0.5, 1.0), 2)
            target = round(np.random.uniform(0.4, 0.7), 2)
            price_cp = round(np.random.uniform(80, 200), 2)
            supply = np.random.randint(50, 300)

            # Date in the middle of period
            date = '2026-08-15'

            rows.append({
                'DATE': date,
                'CHAIN_ID': chain_id,
                'CHAIN_NAME': chain,
                'DRUG_ID': sku_info['drug_id'],
                'SKU_UA': sku_info['sku'],
                'BRAND': sku_info['brand'],
                'PENETRATION': penetration,
                'PRICE_CP': price_cp,
                'SUPPLY': supply,
                'PENETRATION_TARGET': target,
            })

    df = pd.DataFrame(rows)
    return df


# ============================================================
# MAIN - Generate all test files
# ============================================================

def main():
    output_dir = Path(__file__).resolve().parent / 'test_data_output'
    output_dir.mkdir(exist_ok=True)

    print(f"Generating test data in: {output_dir}")

    # 1. Requisites
    print("\n1. Creating requisites...")
    df_req = create_test_requisites()
    req_file = output_dir / 'test_requisites.csv'
    df_req.to_csv(req_file, index=False, encoding='utf-8')
    print(f"   Saved: {req_file} ({len(df_req)} rows)")

    # 2. Pharmacy data
    print("\n2. Creating pharmacy (merchandising) data...")
    df_pharm = create_test_pharmacy_data()
    pharm_file = output_dir / 'test_pharmacy_data.csv'
    df_pharm.to_csv(pharm_file, index=False, encoding='utf-8')
    print(f"   Saved: {pharm_file} ({len(df_pharm)} rows)")
    print(f"   Chains: {df_pharm['CHAIN_NAME'].unique()}")
    print(f"   Brands: {df_pharm['BRAND'].nunique()}")

    # 3. MHL data
    print("\n3. Creating MHL (availability) data...")
    df_mhl = create_test_mhl_data()
    mhl_file = output_dir / 'test_mhl_data.csv'
    df_mhl.to_csv(mhl_file, index=False, encoding='utf-8')
    print(f"   Saved: {mhl_file} ({len(df_mhl)} rows)")
    print(f"   Chains: {df_mhl['CHAIN_NAME'].unique()}")
    print(f"   SKUs: {df_mhl['SKU_UA'].nunique()}")

    # 4. Also create Excel versions
    print("\n4. Creating Excel versions...")
    df_req.to_excel(output_dir / 'test_requisites.xlsx', index=False)
    df_pharm.to_excel(output_dir / 'test_pharmacy_data.xlsx', index=False)
    df_mhl.to_excel(output_dir / 'test_mhl_data.xlsx', index=False)
    print("   Excel files created")

    print("\n✅ Test data generation complete!")
    print(f"\nFiles created in {output_dir}:")
    for f in sorted(output_dir.glob('*')):
        print(f"  - {f.name}")


if __name__ == '__main__':
    main()