import requests
import json

webhook_url = "https://labkabinet.bitrix24.ru/rest/6808/9wti8nc7t0j9r2c7/"
DEAL_ID = 25034

# === 1. Получаем строки товаров сделки ===
deal_rows_url = f"{webhook_url}crm.deal.productrows.get"
resp = requests.post(deal_rows_url, json={"id": DEAL_ID})
resp.raise_for_status()
data = resp.json()

if not data.get('result'):
    print("❌ В сделке нет строк товаров.")
    exit()

product_rows = data['result']
product_ids = [str(row['PRODUCT_ID']) for row in product_rows if row.get('PRODUCT_ID')]

print(f"📦 Найдено {len(product_ids)} товаров: {product_ids}")

# === 2. Получаем детали через catalog.product.get ===
catalog_products = {}

for pid in product_ids:
    try:
        url = f"{webhook_url}catalog.product.get"
        response = requests.post(url, json={"id": pid})
        if response.status_code == 200:
            prod = response.json().get('result', {})
            # Сохраняем количество и цену из строки сделки
            for row in product_rows:
                print(row)
                if str(row.get('PRODUCT_ID')) == pid:
                    prod['QUANTITY'] = row.get('QUANTITY', 1)
                    prod['PRICE'] = row.get('PRICE', prod.get('PRICE', 0))
                    break
            catalog_products[pid] = prod
            print(f"✅ Загружен товар ID={pid}: {prod.get('NAME')}")
        else:
            print(f"⚠️ catalog.product.get вернул ошибку для ID={pid}: {response.status_code} — {response.text}")
    except Exception as e:
        print(f"❌ Ошибка при загрузке товара {pid}: {e}")

# === 3. Выводим ключевые поля ===
print("\n=== Данные товаров из каталога ===")
for pid, p in catalog_products.items():
    print({
        "ID": pid,
        "Название": p.get("NAME"),
        "Кол-во": p.get("QUANTITY"),
        "Цена закуп с НДС": p.get("PRICE"),
        "Поставщик": p.get("PROPERTY_123"),  # ← замените 123 на реальный ID свойства
        "НДС/входящий": p.get("PROPERTY_124"),
        "Вес, г": p.get("PROPERTY_125"),
        "Длина, мм": p.get("PROPERTY_126"),
        "Ширина, мм": p.get("PROPERTY_127"),
        "Высота, мм": p.get("PROPERTY_128"),
    })
