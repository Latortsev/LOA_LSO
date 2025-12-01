# === Настройки обновления ===
import os
UPDATE_BASE_URL = "https://raw.githubusercontent.com/Latortsev/LOA_LSO/main/"

FILES_TO_UPDATE = [
    "main.py",
    "gui.pyw",
    "install.bat",
    "Шаблоны/Расчет_шаблон_V1.xlsx",
    "README.docx",
    "requirements.txt",
    "updater.py",
    "config.py",
]
LOCAL_APP_DIR = os.path.dirname(os.path.abspath(__file__))


# === Настройки ===

DEAL_ID = 25034
INPUT_DIR = os.path.join(LOCAL_APP_DIR, f"Шаблоны")
OUTPUT_DIR = os.path.join(LOCAL_APP_DIR, f"Расчеты")

TEMPLATE_FILE = os.path.join(INPUT_DIR, f"Расчет_шаблон_V1.xlsx")

COLUMN_LABELS = {
    # === Поля строки сделки (DEAL_ROW_*) ===
    "PRODUCT_ID": "ID товара в каталоге",
    "PRODUCT_NAME": "Название товара (в сделке)",
    "PROPERTY_216": "Цена закупа",
    "QUANTITY": "Количество",
    "PROPERTY_228": "Поставщик",
    "PROPERTY_236": "Ставка НДС/входящий",
    "PROPERTY_234": "Артикул поставщика",
    "PROPERTY_206": "Ссылка на товар",
    "PROPERTY_200": "Наценка",
    "PRICE": "Цена в сделке",
    "TAX_RATE": "Ставка налога (%)",
    "TAX_INCLUDED": "Налог включён в цену",
    "MEASURE_NAME": "Единица измерения",
    "PROPERTY_214": "Объём, м³",
    "PROPERTY_232": "Вес,г",
    "PROPERTY_270": "Высота, мм",
    "PROPERTY_268": "Ширина, мм",
    "PROPERTY_266": "Длина, мм",
    "PROPERTY_238": "Срок отгрузки",
    "PROPERTY_242": "Бронирование от производителя",
    "PROPERTY_244": "Реестр Минпрома (Да/Нет)",
    "PROPERTY_204": "Страна производства",
    "PROPERTY_212": "Реестровая запись в Минпроме",
    "PROPERTY_262": "Наличие ограничений по поставке",
    "PROPERTY_260": "Классификация по ЕАЭС (ТН ВЭД ЕАЭС)",
    "PROPERTY_272": "РРЦ",
    "PROPERTY_258": "Производитель",
    "PROPERTY_256": "Техническое задание (ТЗ) с защитой позиции",
    "PROPERTY_254": "Приказ 838",
    "PROPERTY_252": "Номер ГОСТ / ТУ / СТО",
    "PROPERTY_250": "Сертификаты / Декларации соответствия",
    "PROPERTY_248": "Код КТРУ",
    "PROPERTY_246": "Код ОКПД2",




    "PRODUCT_DESCRIPTION": "Описание товара",
    "PROPERTY_194": "Техническое описание",
    "SORT": "Сортировка",
    "XML_ID": "Внешний ID (XML_ID)",
    "TYPE": "Тип строки",
    "STORE_ID": "ID склада",
    "RESERVE_ID": "ID резерва",
    "DATE_RESERVE_END": "Дата окончания резерва",
    "RESERVE_QUANTITY": "Зарезервированное количество",
    "ID": "ID строки товара",
    "OWNER_ID": "ID сделки",
    "OWNER_TYPE": "Тип владельца",
    "ORIGINAL_PRODUCT_NAME": "Оригинальное название",
    "PRICE_EXCLUSIVE": "Цена без скидок",
    "PRICE_NETTO": "Цена нетто",
    "PRICE_BRUTTO": "Цена брутто",
    "PRICE_ACCOUNT": "Бухгалтерская цена",
    "DISCOUNT_TYPE_ID": "Тип скидки",
    "DISCOUNT_RATE": "Размер скидки (%)",
    "DISCOUNT_SUM": "Сумма скидки",
    "CUSTOMIZED": "Изменено вручную",
    "MEASURE_CODE": "Код единицы измерения",

    # === Поля из каталога (товара) ===
    "NAME": "Название товара (каталог)",
    "CODE": "Символьный код",
    "ACTIVE": "Активен",
    "CATALOG_ID": "ID каталога",
    "SECTION_ID": "ID раздела",
    "DESCRIPTION": "Описание (каталог)",
    "VAT_ID": "Ставка НДС",
    "VAT_INCLUDED": "НДС включён",
    "DESCRIPTION_TYPE": "Тип описания",
    "CURRENCY_ID": "Валюта",
    "MEASURE": "Единица измерения (каталог)",
    "PREVIEW_PICTURE": "Превью изображение",
    "DETAIL_PICTURE": "Детальное изображение",
    "TIMESTAMP_X": "Дата изменения",
    "DATE_CREATE": "Дата создания",
    "MODIFIED_BY": "Изменил",
    "CREATED_BY": "Создал",

    # === Пользовательские свойства (PROPERTY_XXX) ===
    # ⚠️ Эти названия — предположения! Замени на реальные, если знаешь точные.
    "PROPERTY_108": "Картинка товара",
    "PROPERTY_218": "ООО с НДС",
    "PROPERTY_220": "ИП без НДС",
    "PROPERTY_240": "Актуальная цена",
    "PROPERTY_202": "Дата расчета",


    # === Служебные поля из твоего кода ===
    #"ID строки в сделке": "ID строки в сделке",
    #"Из каталога?": "Из каталога?",
}