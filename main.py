from data_fetching import fetch_data_from_stream_or_file
from data_preparation import prepare_data
from data_analysis import analyze_all
import os

# URL для потока данных из переменной окружения
STREAM_URL = os.getenv("STREAM_URL", "http://5.181.20.204:8080/api/v1/stream-data")

# Получение данных
df = fetch_data_from_stream_or_file(STREAM_URL)

# Подготовка данных
df = prepare_data(df)

# Выполнение анализа и визуализации

analyze_all(df)
