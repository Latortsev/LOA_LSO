import os
import shutil
import requests
from openpyxl import load_workbook

# === Настройки ===
WEBHOOK_URL = "https://labkabinet.bitrix24.ru/rest/6808/9wti8nc7t0j9r2c7/"
DEAL_ID = 25034
TEMPLATE_FILE = "Расчет_шаблон_V1.xlsx"
OUTPUT_DIR = "Расчеты"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"расчет_{DEAL_ID}.xlsx")

# === 1. Получение строк товаров из сделки ===
def get_deal_product_rows(deal_id):
    url = f"{WEBHOOK_URL}crm.deal.productrows.get"
    response = requests.post(url, json={"id": deal_id})
    response.raise_for_status()
    data = response.json()
    return data.get("result", [])

# === 2. Получение данных товара из каталога магазина ===
def get_catalog_product(product_id):
    url = f"{WEBHOOK_URL}catalog.product.get"
    response = requests.post(url, json={"id": product_id})
    if response.status_code == 200:
        return response.json().get("result", {})
    else:
        print(f"⚠️ Не удалось загрузить товар ID={product_id}: {response.text}")
        return None

# === 3. Заполнение Excel-файла ===
def fill_excel_template(products, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    shutil.copy(TEMPLATE_FILE, output_path)

    wb = load_workbook(output_path)
    ws_calc = wb["Калькулятор"]
    ws_ship = wb["Доставка"]

    # Очистка диапазонов
    for row in range(3, 23):
        for col in "BCDEFG":
            ws_calc[f"{col}{row}"].value = None
    for row in range(2, 22):
        for col in "ABCD":
            ws_ship[f"{col}{row}"].value = None
        for col in "EFGH":
            ws_ship[f"{col}{row}"].value = None

    for i, item in enumerate(products):
        row_calc = i + 3
        row_ship = i + 2

        name = item["NAME"]
        qty = item["QUANTITY"]
        price_with_vat = item["PRICE"]

        # === Определяем поставщика и НДС по property196 ===
        supplier_enum = ""
        if "product" in item and "property196" in item["product"]:
            supplier_enum = item["product"]["property196"].get("valueEnum", "")

        if supplier_enum == "ООО":
            supplier_name = 'ООО "Научные развлечения"'
            vat_in = "НДС 20%"
            shipping_city = "Москва"
        else:
            # fallback — можно расширить под другие поставщики
            supplier_name = supplier_enum if supplier_enum else "Не указан"
            vat_in = "НДС 20%"
            shipping_city = "Москва"

        # === Габариты и вес из product ===
        product_info = item.get("product", {})
        weight = product_info.get("weight", 0)
        length = product_info.get("length", 0)
        width = product_info.get("width", 0)
        height = product_info.get("height", 0)

        # === Калькулятор ===
        ws_calc[f"B{row_calc}"] = name
        ws_calc[f"C{row_calc}"] = qty
        ws_calc[f"D{row_calc}"] = price_with_vat
        ws_calc[f"E{row_calc}"] = supplier_name
        ws_calc[f"F{row_calc}"] = vat_in
        ws_calc[f"G{row_calc}"] = ""  # ссылка не приходит из Bitrix24

        # === Доставка ===
        ws_ship[f"A{row_ship}"] = name
        ws_ship[f"B{row_ship}"] = qty
        ws_ship[f"C{row_ship}"] = supplier_name
        ws_ship[f"D{row_ship}"] = shipping_city
        ws_ship[f"E{row_ship}"] = weight
        ws_ship[f"F{row_ship}"] = length
        ws_ship[f"G{row_ship}"] = width
        ws_ship[f"H{row_ship}"] = height

    wb.save(output_path)
    return output_path

# === Основная логика ===
def main():
    print("1️⃣ Получаем строки товаров из сделки...")
    rows = get_deal_product_rows(DEAL_ID)
    if not rows:
        print("❌ В сделке нет товаров.")
        return

    print("2️⃣ Загружаем данные из каталога...")
    enriched_products = []
    for row in rows:
        product_id = row.get("PRODUCT_ID")
        name = row.get("PRODUCT_NAME", "").strip() or f"Товар ID={product_id}"
        qty = row.get("QUANTITY", 1)
        price = row.get("PRICE", 0)

        catalog_data = None
        if product_id:
            catalog_data = get_catalog_product(product_id)

        item = {
            "NAME": name,
            "QUANTITY": qty,
            "PRICE": price,
        }
        if catalog_data:
            item["product"] = catalog_data  # сохраняем как вложенный объект

        enriched_products.append(item)

    print("3️⃣ Заполняем Excel...")
    output_file = fill_excel_template(enriched_products, OUTPUT_FILE)
    print(f"✅ Готово! Файл сохранён: {output_file}")

# === Запуск ===
if __name__ == "__main__":
    main()