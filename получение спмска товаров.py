from fast_bitrix24 import Bitrix
from openpyxl import Workbook
import json

# Твой webhook
WEBHOOK = "https://labkabinet.bitrix24.ru/rest/6808/9wti8nc7t0j9r2c7/"

# Инициализация клиента
bx = Bitrix(WEBHOOK)

def get_all_products():
    # fast-bitrix сам обрабатывает постраничность
    return bx.get_all('crm.product.list')

def get_product_details(product_id):
    return bx.call('crm.product.get', {'id': product_id})

def get_catalog_product(product_id):
    try:
        return bx.call('catalog.product.get', {'id': product_id})
    except:
        return {}

def get_product_prices(product_id):
    try:
        return bx.get_all('catalog.price.list', {'filter': {'PRODUCT_ID': product_id}})
    except:
        return []

def export_to_excel(products):
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    headers = set()
    rows = []

    for product in products:
        pid = product["ID"]
        details = get_product_details(pid)
        catalog_info = get_catalog_product(pid)
        prices = get_product_prices(pid)

        row_data = {}
        row_data.update(product)
        row_data.update(details or {})
        row_data.update(catalog_info or {})
        row_data["PRICES"] = prices

        headers.update(row_data.keys())
        rows.append(row_data)

    headers = list(headers)
    ws.append(headers)

    def safe_value(value):
        if isinstance(value, (dict, list)):
            return json.dumps(value, ensure_ascii=False)
        return value

    for row in rows:
        ws.append([safe_value(row.get(h, "")) for h in headers])

    wb.save("products_full.xlsx")
    print("✅ Файл products_full.xlsx успешно создан.")

if __name__ == "__main__":
    print("📦 Получаем товары из Bitrix24 через fast-bitrix24...")
    products = get_all_products()
    print(f"🔗 Найдено товаров: {len(products)}")
    export_to_excel(products)


