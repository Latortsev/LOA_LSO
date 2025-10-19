import os

import pandas as pd
import requests
WEBHOOK_URL = "https://labkabinet.bitrix24.ru/rest/6808/9wti8nc7t0j9r2c7/"
DEAL_ID = 25034
LOCAL_APP_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR= os.path.join(LOCAL_APP_DIR, f"Шаблоны")
OUTPUT_DIR = os.path.join(LOCAL_APP_DIR, f"Расчеты")

TEMPLATE_FILE = os.path.join(INPUT_DIR, f"Расчет_шаблон_V1.xlsx")

def get_deal(deal_id):
    """Получает общие данные из сделки"""
    url = f"{WEBHOOK_URL}crm.deal.get"
    response = requests.post(url, json={"id": DEAL_ID})
    response.raise_for_status()
    return response.json().get("result", {})


def get_deal_products(deal_id):
    """Получает товары из сделки"""
    url = f"{WEBHOOK_URL}crm.deal.productrows.get"
    response = requests.post(url, json={"id": deal_id})
    response.raise_for_status()
    return response.json().get("result", [])


def get_catalog_element(element_id):
    """Получает детальные данные товара из каталога"""
    url = f"{WEBHOOK_URL}crm.product.get"
    try:
        response = requests.post(url, json={"id": element_id})
        response.raise_for_status()
        return response.json().get("result", {})
    except requests.exceptions.HTTPError as e:
        if response.status_code == 400:
            print(f"Товар с ID {element_id} не найден в каталоге (ручная позиция)")
            return None  # Возвращаем None для ручных позиций
        else:
            raise e


def load_bitrix_data_to_dataframe(deal_id):
    """
    Загружает данные из Битрикса в pandas DataFrame
    Включает: данные сделки, товары в сделке, детальные данные товаров из каталога
    """
    print(f"Загружаем данные для сделки {deal_id}...")

    # Получаем данные сделки
    print("1. Получаем данные сделки...")
    deal_data = get_deal(deal_id)

    # Получаем товары в сделке
    print("2. Получаем товары в сделке...")
    deal_products = get_deal_products(deal_id)
    print (deal_products)

    # Список для хранения всех данных
    all_data = []

    print("3. Получаем детальные данные товаров из каталога...")
    for i, product_row in enumerate(deal_products):
        print(f"Обрабатываем товар {i + 1}/{len(deal_products)}: {product_row.get('PRODUCT_NAME', 'Unknown')}")

        # Получаем детальные данные товара из каталога
        product_details = get_catalog_element(product_row['PRODUCT_ID'])
        print (product_details)

        # Объединяем данные: сделка + товар в сделке + детали товара
        combined_data = {
            # Данные из сделки
            'DEAL_ID': deal_data.get('ID'),
            'DEAL_TITLE': deal_data.get('TITLE'),
            'DEAL_STAGE': deal_data.get('STAGE_ID'),
            'DEAL_OPPORTUNITY': deal_data.get('OPPORTUNITY'),
            'DEAL_CURRENCY': deal_data.get('CURRENCY_ID'),
            'DEAL_BEGINDATE': deal_data.get('BEGINDATE'),
            'DEAL_CLOSEDATE': deal_data.get('CLOSEDATE'),
            'DEAL_ASSIGNED_BY_ID': deal_data.get('ASSIGNED_BY_ID'),
            'DEAL_CONTACT_ID': deal_data.get('CONTACT_ID'),
            'DEAL_COMPANY_ID': deal_data.get('COMPANY_ID'),

            # Данные из товара в сделке
            'PRODUCT_ROW_ID': product_row.get('ID'),
            'PRODUCT_NAME_IN_DEAL': product_row.get('PRODUCT_NAME'),
            'PRODUCT_PRICE_IN_DEAL': product_row.get('PRICE'),
            'PRODUCT_QUANTITY_IN_DEAL': product_row.get('QUANTITY'),
            'PRODUCT_SUM_IN_DEAL': product_row.get('SUM'),
            'PRODUCT_CURRENCY_IN_DEAL': product_row.get('CURRENCY_ID'),
            'PRODUCT_ID_IN_DEAL': product_row.get('PRODUCT_ID'),

            # Данные из каталога (если товар есть в каталоге)
        }

        # Добавляем данные из каталога или заполняем нулями/пустыми значениями
        if product_details is not None:
            # Товар есть в каталоге
            combined_data.update({
                'CATALOG_PRODUCT_ID': product_details.get('ID'),
                'CATALOG_PRODUCT_NAME': product_details.get('NAME'),
                'CATALOG_PRODUCT_PRICE': product_details.get('PRICE'),
                'CATALOG_PRODUCT_CURRENCY': product_details.get('CURRENCY'),
                'CATALOG_PRODUCT_WEIGHT': product_details.get('WEIGHT'),
                'CATALOG_PRODUCT_WIDTH': product_details.get('WIDTH'),
                'CATALOG_PRODUCT_HEIGHT': product_details.get('HEIGHT'),
                'CATALOG_PRODUCT_LENGTH': product_details.get('LENGTH'),
                'CATALOG_PRODUCT_MEASURE': product_details.get('MEASURE'),
                'CATALOG_PRODUCT_DESCRIPTION': product_details.get('DESCRIPTION'),
                'CATALOG_PRODUCT_SECTION_ID': product_details.get('SECTION_ID'),

                # Пользовательские поля из каталога
                'CATALOG_UF_ARTICLE': product_details.get('UF_ARTICLE'),
                'CATALOG_UF_MANUFACTURER': product_details.get('UF_MANUFACTURER'),
                'CATALOG_UF_VENDOR_CODE': product_details.get('UF_VENDOR_CODE'),
            })

            # Добавляем другие пользовательские поля из каталога
            for key, value in product_details.items():
                if key.startswith('UF_') and not key.startswith('UF_CRM_'):
                    combined_data[f'CATALOG_{key}'] = value
        else:
            # Ручная позиция - заполняем нулями/пустыми значениями
            print(f"  -> Ручная позиция: заполняем поля нулями")
            combined_data.update({
                'CATALOG_PRODUCT_ID': product_row.get('PRODUCT_ID'),  # ID из строки сделки
                'CATALOG_PRODUCT_NAME': product_row.get('PRODUCT_NAME'),  # Название из строки сделки
                'CATALOG_PRODUCT_PRICE': 0,
                'CATALOG_PRODUCT_CURRENCY': product_row.get('CURRENCY_ID'),
                'CATALOG_PRODUCT_WEIGHT': 0,
                'CATALOG_PRODUCT_WIDTH': 0,
                'CATALOG_PRODUCT_HEIGHT': 0,
                'CATALOG_PRODUCT_LENGTH': 0,
                'CATALOG_PRODUCT_MEASURE': 0,
                'CATALOG_PRODUCT_DESCRIPTION': '',
                'CATALOG_PRODUCT_SECTION_ID': 0,

                # Пользовательские поля - заполняем None
                'CATALOG_UF_ARTICLE': None,
                'CATALOG_UF_MANUFACTURER': None,
                'CATALOG_UF_VENDOR_CODE': None,
            })

        # Добавляем все пользовательские поля из сделки
        for key, value in deal_data.items():
            if key.startswith('UF_CRM_'):
                combined_data[f'DEAL_{key}'] = value

        all_data.append(combined_data)

    # Создаем DataFrame
    df = pd.DataFrame(all_data)

    print(f"Данные успешно загружены. DataFrame содержит {len(df)} строк и {len(df.columns)} столбцов")

    return df, deal_data


# Пример использования:
def deal_to_dataframe(deal_id):
    df, deal_info = load_bitrix_data_to_dataframe(deal_id)

    # Сохраняем в Excel если нужно
    output_file = os.path.join(OUTPUT_DIR, str(deal_id), f"данные_сделки_{deal_id}.xlsx")
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Основные данные
        df.to_excel(writer, sheet_name='Товары_сделки', index=False)

        # Данные сделки отдельно
        deal_df = pd.DataFrame([deal_info])
        deal_df.to_excel(writer, sheet_name='Инфо_сделки', index=False)

    print(f"Данные сохранены в {output_file}")
    return df


# Если нужно только загрузить данные без сохранения в Excel:
def load_deal_data(deal_id):
    df, deal_info = load_bitrix_data_to_dataframe(deal_id)
    return df, deal_info


def save_deal_data_ordered(deal_id, df, deal_info):
    """Сохраняет данные в Excel с заданным порядком колонок и ссылкой на сделку"""

    # Генерируем ссылку на сделку
    bitrix_url = WEBHOOK_URL.split('/rest')[0]  # Извлекаем базовый URL из webhook
    deal_link = generate_deal_link(deal_id, bitrix_url)

    # Определяем желаемый порядок колонок
    desired_columns = [
        'DEAL_ID',
        'DEAL_TITLE',
        'DEAL_LINK',  # Ссылка на сделку (добавляем вручную)
        'DEAL_STAGE',
        'DEAL_OPPORTUNITY',
        'DEAL_CURRENCY',
        'DEAL_BEGINDATE',
        'DEAL_CLOSEDATE',
        'DEAL_ASSIGNED_BY_ID',
        'DEAL_CONTACT_ID',
        'DEAL_COMPANY_ID',
        'PRODUCT_ROW_ID',
        'PRODUCT_NAME_IN_DEAL',
        'PRODUCT_ID_IN_DEAL',
        'PRODUCT_PRICE_IN_DEAL',
        'PRODUCT_QUANTITY_IN_DEAL',
        'PRODUCT_SUM_IN_DEAL',
        'PRODUCT_CURRENCY_IN_DEAL',
        'CATALOG_PRODUCT_ID',
        'CATALOG_PRODUCT_NAME',
        'CATALOG_PRODUCT_PRICE',
        'CATALOG_PRODUCT_CURRENCY',
        'CATALOG_PRODUCT_WEIGHT',
        'CATALOG_PRODUCT_WIDTH',
        'CATALOG_PRODUCT_HEIGHT',
        'CATALOG_PRODUCT_LENGTH',
        'CATALOG_PRODUCT_MEASURE',
        'CATALOG_PRODUCT_DESCRIPTION',
        'CATALOG_PRODUCT_SECTION_ID',
        'CATALOG_UF_ARTICLE',
        'CATALOG_UF_MANUFACTURER',
        'CATALOG_UF_VENDOR_CODE'
    ]

    # Добавляем ссылку на сделку в DataFrame
    df['DEAL_LINK'] = deal_link

    # Переставляем колонки в нужном порядке
    # Сначала выбираем только те колонки, которые есть в DataFrame
    available_columns = [col for col in desired_columns if col in df.columns]

    # Добавляем остальные колонки, которые не были в desired_columns
    remaining_columns = [col for col in df.columns if col not in desired_columns]
    final_columns = available_columns + remaining_columns

    df_ordered = df[final_columns]

    # Сохраняем в Excel
    output_file = os.path.join(OUTPUT_DIR, str(deal_id), f"данные_сделки_{deal_id}.xlsx")
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Основные данные с нужным порядком
        df_ordered.to_excel(writer, sheet_name='Товары_сделки', index=False)

        # Данные сделки отдельно
        deal_df = pd.DataFrame([deal_info])
        deal_df.to_excel(writer, sheet_name='Инфо_сделки', index=False)

    print(f"Данные сохранены в {output_file}")
    print(f"Колонки в нужном порядке: {available_columns}")

    return df_ordered

def load_bitrix_products_to_excel(deal_id, output_file=None, catalog_id=None):
    """
    Выгружает товары из сделки в Excel с русскими названиями колонок.
    Для товаров из каталога также добавляются данные из строки сделки.
    Пользовательские свойства PROPERTY_XXX раскрываются до значений (value).
    """
    print(f"Загружаем товарные позиции для сделки {deal_id}...")

    deal_products = get_deal_products(deal_id)
    if not deal_products:
        print("В сделке нет товаров.")
        df = pd.DataFrame()
        output_file = output_file or f"deal_{deal_id}_products.xlsx"
        df.to_excel(output_file, sheet_name='Товары', index=False)
        return df

    all_rows = []

    for i, product_row in enumerate(deal_products):
        product_id = product_row.get('PRODUCT_ID')
        print(f"Обрабатываем позицию {i + 1}/{len(deal_products)} (ID: {product_id})")

        catalog_data = get_catalog_element(product_id) if product_id else None

        # Начинаем с данных из строки сделки
        row_data = {f"DEAL_ROW_{k}": v for k, v in product_row.items()}

        if catalog_data is not None:
            # Добавляем данные из каталога, но НЕ перезаписываем уже существующие поля из сделки
            for k, v in catalog_data.items():
                if k not in row_data:  # Не перезаписываем, если поле уже есть (например, PRICE из сделки важнее)
                    row_data[k] = v

            # Раскрываем пользовательские свойства PROPERTY_XXX (если это словарь)
            for key in list(row_data.keys()):
                if key.startswith('PROPERTY_') and isinstance(row_data[key], dict):
                    # Извлекаем значение 'value' из словаря
                    row_data[key] = row_data[key].get('value', '')

        row_data['SOURCE_ROW_ID'] = product_row.get('ID')
        row_data['IS_CATALOG_ITEM'] = catalog_data is not None

        all_rows.append(row_data)

    df = pd.DataFrame(all_rows)

    # === Заменяем названия колонок на русские labels ===
    column_labels = {}
    for col in df.columns:
        if col.startswith('DEAL_ROW_'):
            base_name = col.replace('DEAL_ROW_', '')
            column_labels[col] = f"Строка сделки: {base_name}"
        elif col == 'SOURCE_ROW_ID':
            column_labels[col] = 'ID строки в сделке'
        elif col == 'IS_CATALOG_ITEM':
            column_labels[col] = 'Из каталога?'
        else:
            # Это поле из каталога — можно оставить как есть или добавить префикс "Каталог: "
            column_labels[col] = f"Каталог: {col}"

    df.rename(columns=column_labels, inplace=True)

    # Сохраняем
    if output_file is None:
        output_file = f"deal_{deal_id}_products.xlsx"

    df.to_excel(output_file, sheet_name='Товары', index=False)

    print(f"✅ Сохранено {len(df)} позиций в файл: {output_file}")
    print(f"📊 Использовано {len(df.columns)} полей с русскими названиями.")

    return df

# Пример использования:
deal_id = 25034
#deal_to_dataframe(deal_id)
load_bitrix_products_to_excel(deal_id)