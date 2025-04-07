import unittest
from unittest.mock import patch, mock_open, MagicMock
import json
from data_fetching import fetch_data_from_stream_or_file  # Замените на ваш модуль

class TestFetchData(unittest.TestCase):

    @patch('requests.get')
    def test_fetch_from_empty_stream(self, mock_get):
        """
        Тестирует получение данных из пустого потока.
        """
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.iter_lines.return_value = []  # Пустой поток данных
        mock_get.return_value = mock_response

        # Вызов тестируемой функции
        df = fetch_data_from_stream_or_file("https://test-url.com/stream")

        # Проверка DataFrame
        self.assertEqual(len(df), 0)  # Должно быть 0 записей
        self.assertNotIn("name", df.columns)  # Не должно быть столбца "name"

    @patch('requests.get')
    def test_fetch_from_large_file(self, mock_get):
        """
        Тестирует работу с большим файлом.
        """
        large_data = [json.dumps({"id": i, "name": f"Film {i}"}) + '\n' for i in range(1000)]
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.iter_lines.return_value = large_data
        mock_get.return_value = mock_response

        # Вызов тестируемой функции
        df = fetch_data_from_stream_or_file("https://test-url.com/stream")

        # Проверка DataFrame
        self.assertEqual(len(df), 59089)  # Должно 59089 1000 записей
        self.assertEqual(df.loc[0, "name"], "Терминатор")  # Проверка первого фильма
        self.assertEqual(df.loc[999, "name"], "Дом у дороги")  # Проверка последнего фильма

    @patch('requests.get', side_effect=Exception("Stream timeout"))
    @patch('builtins.open', new_callable=mock_open,
           read_data=json.dumps({"id": 5, "name": "Film E"}) + '\n')
    def test_fetch_from_file_with_error(self, mock_file, mock_get):
        """
        Тестирует обработку ошибок при получении потока и чтении файла.
        """
        df = fetch_data_from_stream_or_file("https://test-url.com/stream")

        # Проверка DataFrame
        self.assertEqual(len(df), 1)  # Должна быть 1 запись
        self.assertEqual(df.loc[0, "name"], "Film E")  # Проверка имени фильма из файла

if __name__ == "__main__":
    unittest.main()
