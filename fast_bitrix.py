import json
from fast_bitrix24 import Bitrix

webhook = "https://labkabinet.bitrix24.ru/rest/6808/9wti8nc7t0j9r2c7/  "
bx = Bitrix(webhook)

def print_deal_fields(deal_id: int):
    try:
        deal = get_deal(deal_id)

        if not deal:
            print(f"Сделка с ID {deal_id} не найдена.")
            return

        print(f"\n{'=' * 70}")
        print(f"СДЕЛКА ID: {deal_id} (только коды)")
        print(f"{'=' * 70}\n")

        for key in sorted(deal.keys()):
            value = deal[key]
            if isinstance(value, (dict, list)):
                value_str = json.dumps(value, ensure_ascii=False, indent=2)
            else:
                value_str = str(value) if value is not None else "—"

            print(f"{key:<35} : {value_str}")
            print("-" * 70)

    except Exception as e:
        print(f"❌ Ошибка в print_deal_fields: {e}")

def get_deal(deal_id: int):
    # Получаем сделку и возвращаем словарь с её полями
    result=bx.call('crm.deal.get', {'id': deal_id})
    return result['order0000000000']

if __name__ == "__main__":
# --- Пример использования ---
    deal = get_deal(25034)
    address = deal["UF_CRM_1757930626746"]
    print(address)
    #print_deal_fields(25034)
    products = deal['PRODUCT_ROWS'] or deal['ORDER_PRODUCT_ID']
    print  (products)


