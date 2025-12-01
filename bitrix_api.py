import requests

# === Настройки ===
WEBHOOK_URL = "https://labkabinet.bitrix24.ru/rest/6808/9wti8nc7t0j9r2c7/"

def get_deal(deal_id):
    """Получает общие данные из сделки"""
    url = f"{WEBHOOK_URL}crm.deal.get"
    response = requests.post(url, json={"id": deal_id})
    response.raise_for_status()
    return response.json().get("result", {})

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

def get_deal_products(deal_id):
    """Получает товары из сделки"""
    url = f"{WEBHOOK_URL}crm.deal.productrows.get"
    response = requests.post(url, json={"id": deal_id})
    response.raise_for_status()
    return response.json().get("result", [])

def get_catalog_product(product_id):
    url = f"{WEBHOOK_URL}catalog.product.get"
    response = requests.post(url, json={"id": product_id})
    if response.status_code == 200:
        return response.json().get("result", {})
    else:
        print(f"⚠️ Не удалось загрузить товар ID={product_id}")
        return None