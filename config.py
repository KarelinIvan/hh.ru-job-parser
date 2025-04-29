import requests

def get_hh_vacancies():
    """ Функция для получения данных с hh.ru по API """
    base_url = 'https://api.hh.ru/vacancies'
    headers = {'User-Agent': 'hh.ru-job-parser/1.0 (ivan.karelin.1993@mail.ru)'}

    try:
        # Запрашиваем данные с API hh.ru
        response = requests.get(base_url, headers=headers)
        # Проверка статуса запроса
        response.raise_for_status()
        # Преобразуем ответ в формат JSON
        return response.json().get('items', [])
    except requests.exceptions.RequestException as e:
        print(f'Ошибка при запросе к API: {e}')
        return []