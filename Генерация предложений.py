import requests
import json

def generate_document_with_bx24_method(webhook_url, template_id, entity_type_id, entity_id):
    """
    Вызывает метод crm.documentgenerator.document.add через вебхук, 
    имитируя BX24.callMethod.

    Args:
        webhook_url (str): Базовый URL вебхука (например, https://your_domain.bitrix24.ru/rest/USER_ID/AUTH_CODE/).
        template_id (str): ID шаблона документа.
        entity_type_id (str): ID типа сущности (например, '2' для Lead, '3' для Deal).
        entity_id (str): ID конкретной сущности (например, ID сделки).

    Returns:
        dict: Результат вызова API или информация об ошибке.
    """
    # Построение URL для метода API
    # Вебхук обычно имеет структуру: https://domain.bitrix24.ru/rest/USER_ID/AUTH_CODE/
    # Метод добавляется как: /crm.documentgenerator.document.add.json
    api_method = 'crm.documentgenerator.document.add'
    url = f"{webhook_url.rstrip('/')}/{api_method}.json"

    # Подготовка данных запроса (аналогично параметрам BX24.callMethod)
    # Обычно в вебхуках данные отправляются как JSON в теле POST-запроса
    payload = {
        'templateId': template_id,
        'entityTypeId': entity_type_id,
        'entityId': entity_id
    }

    print(f"Попытка вызвать метод {api_method} с параметрами: {payload}")
    print(f"URL запроса: {url}")

    try:
        # Отправка POST-запроса с JSON-данными
        # Вебхуки Bitrix24 обычно не требуют дополнительных заголовков, таких как Authorization,
        # так как аутентификация встроена в URL вебхука.
        # Однако Content-Type может потребоваться указать, особенно если сервер строго проверяет.
        # headers = {'Content-Type': 'application/json'} # Опционально для вебхуков
        response = requests.post(url, json=payload) # Используем json=payload для отправки JSON
        response.raise_for_status()  # Проверка статуса HTTP (2xx)

        result = response.json()
        print("--- Результат вызова API ---")
        print(json.dumps(result, indent=2, ensure_ascii=False))

        # Проверка на наличие ошибки в теле ответа API (BX24 стиль)
        # Некоторые методы возвращают структурированный ответ
        # Однако стандартный REST API обычно выбрасывает исключение на 4xx/5xx
        # и возвращает данные в 'result' при успехе.
        # BX24.callMethod внутри JS обрабатывает это по-своему.
        # Для стандартного REST API проверяем 'result'.
        if 'result' in result:
            print("Документ успешно создан.")
            return result
        else:
            # Если 'result' нет, но статус 200 OK, возможно, формат ответа другой
            # или произошла ошибка, не обработанная raise_for_status
            print(f"Ответ API не содержит ключа 'result'. Тело: {result}")
            return result

    except requests.exceptions.HTTPError as e:
        # Обработка ошибок HTTP (например, 400, 404, 500)
        print(f"Ошибка HTTP при вызове {api_method}: {e}")
        print(f"Статус код: {response.status_code}")
        print(f"Текст ответа: {response.text}")
        # Попробуем распарсить текст ответа, может содержать информацию об ошибке в стиле API
        try:
            error_details = response.json()
            print(f"Детали ошибки из API: {error_details}")
            return {"error": f"HTTP {response.status_code}", "details": error_details}
        except json.JSONDecodeError:
            print("Не удалось распарсить детали ошибки как JSON.")
            return {"error": f"HTTP {response.status_code}", "details": response.text}
    except requests.exceptions.RequestException as e:
        # Обработка других ошибок запроса (например, проблемы сети)
        print(f"Ошибка запроса: {e}")
        print(f"Текст ответа: {response.text if response else 'N/A'}")
        return {"error": str(e)}
    except json.JSONDecodeError:
        # Обработка ошибки парсинга JSON ответа
        print(f"Ошибка: Не удалось распарсить JSON ответа от {api_method}. Статус: {response.status_code if response else 'N/A'}, Текст: {response.text if response else 'N/A'}")
        return {"error": "Invalid JSON response", "details": response.text if response else 'N/A'}


# --- Пример использования ---
# Замените значения на реальные
webhook_url = "https://labkabinet.bitrix24.ru/rest/6808/9wti8nc7t0j9r2c7/" # Базовый URL вебхука
template_id_to_use = "46" # Замените на реальный ID шаблона КП ЛШО
entity_type_id_to_use = "2" # '2' обычно соответствует CRM_DEAL
entity_id_to_use = "25034" # ID сделки

result = generate_document_with_bx24_method(
     webhook_url,
     template_id_to_use,
     entity_type_id_to_use,
     entity_id_to_use
 )
print("Финальный результат:", result)
