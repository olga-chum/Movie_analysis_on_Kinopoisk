import unittest
import pandas as pd
from data_preparation import convert_to_rub, extract_roles, prepare_data  # Замените на ваш модуль


class TestFunctions(unittest.TestCase):

    def test_convert_to_rub(self):
        """
        Тестирование функции convert_to_rub.
        """
        self.assertEqual(convert_to_rub(100, 'USD'), 9000)  # 100 USD в рубли
        self.assertEqual(convert_to_rub(50, 'EUR'), 5000)  # 50 EUR в рубли
        self.assertEqual(convert_to_rub(1000, 'RUB'), 1000)  # 1000 RUB в рубли
        self.assertEqual(convert_to_rub(100, 'UNKNOWN'), 100)  # Неизвестная валюта
        self.assertIsNone(convert_to_rub(None, 'USD'))  # Пропущенная сумма
        self.assertIsNone(convert_to_rub(100, None))  # Пропущенная валюта

    def test_extract_roles(self):
        """
        Тестирование функции extract_roles.
        """
        persons = [
            {'name': 'Actor 1', 'enProfession': 'actor'},
            {'name': 'Director 1', 'enProfession': 'director'},
            {'name': 'Actor 2', 'enProfession': 'actor'},
        ]
        self.assertEqual(extract_roles(persons, 'actor'), ['Actor 1', 'Actor 2'])  # Два актера
        self.assertEqual(extract_roles(persons, 'director'), ['Director 1'])  # Один режиссер
        self.assertEqual(extract_roles(persons, 'writer'), [])  # Нет сценаристов
        self.assertEqual(extract_roles([], 'actor'), [])  # Пустой список персон
        self.assertEqual(extract_roles(None, 'actor'), [])  # Некорректный вход

    def test_prepare_data(self):
        """
        Тестирование функции prepare_data.
        """
        # Пример входных данных
        data = {
            'id': [1, 2],
            'name': ['Film A', 'Film B'],
            'genres': [[{'name': 'Drama'}], [{'name': 'Comedy'}]],
            'countries': [[{'name': 'USA'}], [{'name': 'Russia'}]],
            'budget.value': [100, 200],
            'budget.currency': ['USD', 'EUR'],
            'fees.usa.value': [50, 30],
            'fees.usa.currency': ['USD', 'USD'],
            'fees.russia.value': [20, 10],
            'fees.russia.currency': ['RUB', 'RUB'],
            'fees.world.value': [200, 150],
            'fees.world.currency': ['USD', 'EUR'],
            'persons': [
                [{'name': 'Actor 1', 'enProfession': 'actor'}],
                [{'name': 'Director 1', 'enProfession': 'director'}]
            ],
            'rating.kp': [7.5, 0],  # У второго фильма рейтинг.kp = 0
            'rating.imdb': [8.0, 0],  # У второго фильма рейтинг.imdb = 0
            'votes.kp': [1000, 800],
            'votes.imdb': [2000, 1500],
        }
        df = pd.DataFrame(data)

        # Вызов тестируемой функции
        result = prepare_data(df)

        # Проверка итогового DataFrame
        self.assertIn('budget_rub', result.columns)  # Столбец бюджет в рублях
        self.assertIn('fees_rub_usa', result.columns)  # Столбец сборов в США в рублях
        self.assertIn('actors', result.columns)  # Столбец актеров
        self.assertIn('directors', result.columns)  # Столбец режиссеров
        self.assertEqual(len(result), 1)  # Должен остаться только 1 фильм (фильтрация по рейтингу)
        self.assertEqual(result.loc[0, 'genres'], ['Drama'])  # Проверка преобразования жанра
        self.assertEqual(result.loc[0, 'countries'], ['USA'])  # Проверка преобразования стран
        self.assertEqual(result.loc[0, 'actors'], ['Actor 1'])  # Извлеченные актеры


if __name__ == "__main__":
    unittest.main()
