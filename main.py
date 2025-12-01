import os
from bitrix_api import *
from excel_handler import *
from updater import *

# === Настройки ===
DEAL_ID = 25034

# === Основная логика ===
def main():
    print("=== ЗАПУСК ПРИЛОЖЕНИЯ ===")
    
    # Проверка обновлений
    updater = Updater(
        "https://bitrix24public.com/labkabinet.bitrix24.ru/docs/pub/a74e057419b211005403b334135e4de9/default/?&",
        [
            "main.py",
            "gui.pyw",
            "install.bat",
            "Шаблоны/Расчет_шаблон_V1.xlsx",
            "README.docx",
            "requirements.txt"
        ],
        os.path.dirname(os.path.abspath(__file__))
    )
    updater.auto_update_check()
    
    # Получение товаров из сделки
    print("1️⃣ Получаем товары из сделки...")
    deal_products = get_deal_products(DEAL_ID)
    print(f"Получено {len(deal_products)} товаров из сделки {DEAL_ID}")
    
    # Экспорт товаров в базу данных (Excel)
    print("2️⃣ Экспортируем товары в Excel...")
    output_dir = os.path.join(OUTPUT_DIR, str(DEAL_ID))  # используем путь из excel_handler
    db_output_path = os.path.join(output_dir, f"база_данных_{DEAL_ID}.xlsx")
    os.makedirs(os.path.dirname(db_output_path), exist_ok=True)
    export_products_to_db(deal_products, db_output_path, get_catalog_product)
    
    # Заполнение Excel шаблона
    print("4️⃣ Заполняем Excel шаблон...")
    fill_excel(deal_products, DEAL_ID, get_deal)
    
    print("=== ЗАВЕРШЕНИЕ РАБОТЫ ===")

if __name__ == "__main__":
    main()