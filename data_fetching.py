import json
import requests
import pandas as pd

def fetch_data_from_stream_or_file(stream_url):
    """
    Получает данные из потока или локального файла, если поток недоступен.

    :param stream_url: URL потока данных
    :param local_file_path: Путь к локальному CSV файлу
    :return: DataFrame с данными
    """


    all_movies = []

    try:
        # Отправляем запрос
        response = requests.get(stream_url, stream=True, timeout=10)
        # Проверяем статус ответа
        if response.status_code == 200:
            print("Данные успешно получены из потока")
            # Обрабатываем поток построчно
            for line in response.iter_lines():
                if line:
                    all_movies.append(json.loads(line.decode('utf-8')))
            # Преобразуем данные из потока в DataFrame
            df = pd.json_normalize(all_movies)
        else:
            print(f"Ошибка при запросе данных: статус"
                  f" {response.status_code}. Использую локальный файл.")

            file_path = 'stream-data'
            # Читаем файл построчно
            with open(file_path, 'r', encoding='utf-8') as file:
                for line in file:
                    if line.strip():  # Проверка на пустые строки
                        all_movies.append(json.loads(line.strip()))

            # Преобразуем данные из файла в DataFrame
            df = pd.json_normalize(all_movies)

            # Сообщение о завершении обработки
            print("Данные успешно считаны из файла и преобразованы в DataFrame")
    except Exception:
        file_path = 'stream-data'
        # Читаем файл построчно
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                if line.strip():  # Проверка на пустые строки
                    all_movies.append(json.loads(line.strip()))

        # Преобразуем данные из файла в DataFrame
        df = pd.json_normalize(all_movies)

        # Сообщение о завершении обработки
        print("Данные успешно считаны из файла и преобразованы в DataFrame")

    return df
