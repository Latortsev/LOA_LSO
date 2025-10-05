import requests

WEBHOOK = "https://labkabinet.bitrix24.ru/rest/6808/9wti8nc7t0j9r2c7/"



def update_deal_prices(deal_id: int, shipper: str):
    rows_resp = requests.post(WEBHOOK + "crm.deal.productrows.get", json={"id": deal_id}).json()
    rows = rows_resp.get("result", [])

    new_rows = []
    for row in rows:
        product_id = row["PRODUCT_ID"]

        # Получаем карточку товара
        product_resp = requests.post(WEBHOOK + "catalog.product.get", json={"id": product_id}).json()
        product = product_resp.get("result", {}).get("product", {})

        # Берём цену из пользовательского свойства
        if shipper == "ООО":
            raw_price = product.get("property222", {}).get("value")
        else:
            raw_price = product.get("property224", {}).get("value")

        print(f"Цена товара {product_id}: {raw_price}")

        # Убираем валюту, если приходит "120|RUB"
        price = None
        if raw_price:
            price = raw_price.split("|")[0]

        if price:
            row["PRICE"] = float(price)
            row["PRICE_EXCLUSIVE"] = float(price)
            row["PRICE_NETTO"] = float(price)
            row["PRICE_BRUTTO"] = float(price)

        # Обязательно передаём все ключевые поля
        new_rows.append({
            "PRODUCT_ID": row["PRODUCT_ID"],
            "PRODUCT_NAME": row["PRODUCT_NAME"],
            "PRICE": row["PRICE"],
            "QUANTITY": row["QUANTITY"],
            "DISCOUNT_TYPE_ID": row.get("DISCOUNT_TYPE_ID", 1),
            "DISCOUNT_RATE": row.get("DISCOUNT_RATE", 0),
            "DISCOUNT_SUM": row.get("DISCOUNT_SUM", 0),
            "TAX_RATE": row.get("TAX_RATE", 0),
            "TAX_INCLUDED": row.get("TAX_INCLUDED", "Y"),
            "MEASURE_CODE": row.get("MEASURE_CODE", 796),
            "MEASURE_NAME": row.get("MEASURE_NAME", "шт"),
            "SORT": row.get("SORT", 100)
        })

    update_resp = requests.post(WEBHOOK + "crm.deal.productrows.set", json={"id": deal_id, "rows": new_rows}).json()
    return update_resp




import json, requests

resp = requests.post(WEBHOOK + "catalog.product.get", json={"id": 1036}).json()
print(json.dumps(resp, indent=2, ensure_ascii=False))

# Для сделки 25034 и компании отгрузки ООО
print(update_deal_prices(25034, "ООО"))

# Для сделки 25034 и компании отгрузки ИП
# print(update_deal_prices(25034, "ИП"))