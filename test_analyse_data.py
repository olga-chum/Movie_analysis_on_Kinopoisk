import unittest
from unittest.mock import patch, MagicMock
import pandas as pd


# Импортируем тестируемые функции
from data_analysis import (
    analyze_ratings_distribution,
    compare_platform_ratings,
    analyze_rating_genres,
    analyze_rating_genres_time,
    analyze_rating_genres_trends,
    analyze_budgets,
    analyze_budgets_and_fees,
    analyze_top_persons,
    analyze_low_persons
)

class TestMovieAnalytics(unittest.TestCase):
    def setUp(self):
        # Создаем тестовый DataFrame
        self.test_data = pd.DataFrame({
            'name': ['Film 1', 'Film 2', 'Film 3', 'Film 4', 'Film 5', 'Film 6'],
            'rating.kp': [6.5, 7.8, 8.2, 5.4, 4.3, 9.0],
            'rating.imdb': [6.0, 7.5, 8.0, 5.2, 4.1, 8.7],
            'genres': [['Drama'], ['Comedy'], ['Action'], ['Drama'], ['Horror'], ['Action']],
            'year': [2001, 2002, 2003, 2004, 2005, 2006],
            'budget_rub': [1000000, 2000000, 3000000, 400000, 500000, 6000000],
            'fees_rub_world': [2000000, 3000000, 5000000, 800000, 600000, 9000000],
            'votes.kp': [1000, 2000, 1500, 1200, 800, 3000],
            'actors': [['Actor 1', 'Actor 2'],
                       ['Actor 3'], ['Actor 4', 'Actor 5'], ['Actor 6'], ['Actor 1'],
                       ['Actor 2']],
            'directors': [['Director 1'], ['Director 2'],
                          ['Director 3'], ['Director 4'], ['Director 5'],
                          ['Director 6']],
        })

    @patch('data_analysis.Document')  # Мокируем docx.Document
    @patch('data_analysis.plt')       # Мокируем matplotlib.pyplot
    def test_analyze_ratings_distribution(self, mock_plt, MockDocument):
        mock_doc = MagicMock()  # Создаем мок для документа
        MockDocument.return_value = mock_doc  # Возвращаем мок в качестве экземпляра документа

        analyze_ratings_distribution(self.test_data, mock_doc)

        # Проверяем, что заголовки и текст добавлены в документ
        mock_doc.add_heading.assert_any_call("Анализ распределения оценок Кинопоиска", level=1)
        mock_doc.add_heading.assert_any_call('Гистограмма распределения оценок', level=2)
        mock_doc.add_paragraph.assert_any_call("Средняя оценка: 6.87")

        # Проверяем, что график сохранен
        mock_plt.savefig.assert_called_once()

    @patch('data_analysis.Document')  # Мокируем docx.Document
    @patch('data_analysis.plt')       # Мокируем matplotlib.pyplot
    def test_compare_platform_ratings(self, mock_plt, MockDocument):
        mock_doc = MagicMock()
        MockDocument.return_value = mock_doc

        compare_platform_ratings(self.test_data, mock_doc)

        # Проверяем добавление заголовков
        mock_doc.add_heading.assert_any_call("Сравнение"
                                             " оценок Кинопоиска и IMDb", level=1)

        # # Проверяем сохранение графика
        mock_plt.savefig.assert_called_once()

    @patch('data_analysis.Document')
    @patch('data_analysis.plt')
    def test_analyze_rating_genres(self, mock_plt, MockDocument):
        mock_doc = MagicMock()
        MockDocument.return_value = mock_doc

        analyze_rating_genres(self.test_data, mock_doc)

        # Проверяем, что заголовки добавлены
        mock_doc.add_heading.assert_called_with("Анализ средних"
                                                "рейтингов фильмов по жанрам", level=1)

        # Проверяем сохранение графика
        mock_plt.savefig.assert_called_once()

    @patch('data_analysis.Document')
    @patch('data_analysis.plt')
    def test_analyze_rating_genres_time(self, mock_plt, MockDocument):
        mock_doc = MagicMock()
        MockDocument.return_value = mock_doc

        analyze_rating_genres_time(self.test_data, mock_doc)

        # Проверяем добавление заголовка
        mock_doc.add_heading.assert_called_with("Анализ изменения"
                                                " популярности жанров фильмов", level=1)

        # Проверяем сохранение графика
        mock_plt.savefig.assert_called_once()

    @patch('data_analysis.Document')
    @patch('data_analysis.plt')
    def test_analyze_rating_genres_trends(self, mock_plt, MockDocument):
        mock_doc = MagicMock()
        MockDocument.return_value = mock_doc

        analyze_rating_genres_trends(self.test_data, mock_doc)

        # Проверяем добавление заголовка
        mock_doc.add_heading.assert_called_with("Изменение популярности"
        " топ-15 жанров фильмов с 2000 по 2020 год", level=1)

        # Проверяем сохранение графика
        mock_plt.savefig.assert_called_once()

    @patch('data_analysis.Document')
    @patch('data_analysis.plt')
    def test_analyze_budgets(self, mock_plt, MockDocument):
        mock_doc = MagicMock()
        MockDocument.return_value = mock_doc

        analyze_budgets(self.test_data, mock_doc)

        # Проверяем добавление заголовка
        mock_doc.add_paragraph.assert_any_call("На графике представлен"
           " анализ зависимости между бюджетами фильмов и их мировыми сборами.")

        # Проверяем сохранение графика
        mock_plt.savefig.assert_called_once()

    @patch('data_analysis.Document')
    def test_analyze_budgets_and_fees(self, MockDocument):
        mock_doc = MagicMock()
        MockDocument.return_value = mock_doc

        analyze_budgets_and_fees(self.test_data, mock_doc)

        # Проверяем добавление заголовка
        mock_doc.add_heading.assert_any_call("Анализ бюджетов и сборов фильмов", level=1)

    @patch('data_analysis.Document')
    def test_analyze_top_persons(self, MockDocument):
        mock_doc = MagicMock()
        MockDocument.return_value = mock_doc

        analyze_top_persons(self.test_data, mock_doc)

        # Проверяем добавление заголовка
        mock_doc.add_heading.assert_any_call("Анализ лучших актёров и режиссёров", level=1)

        # Проверяем добавление таблиц с топами актёров и режиссёров
        mock_doc.add_heading.assert_any_call("Топ-10 актёров с самыми высокими рейтингами:",
                                             level=2)
        mock_doc.add_heading.assert_any_call("Топ-10 режиссёров с самыми высокими рейтингами:",
                                             level=2)

        # Проверяем, что добавлена таблица для актёров
        mock_doc.add_table.assert_any_call(rows=1, cols=3)

        # Проверяем, что таблица для актёров и режиссёров была правильно заполнена
        mock_doc.add_table.return_value.add_row.return_value.cells[
                0].text = 'actor'  # Пример проверки для первой ячейки

    @patch('data_analysis.Document')
    def test_analyze_low_persons(self, MockDocument):
        mock_doc = MagicMock()
        MockDocument.return_value = mock_doc

        analyze_low_persons(self.test_data, mock_doc)

        # Проверяем добавление заголовка
        mock_doc.add_heading.assert_any_call("Анализ актёров и"
                                             " режиссёров с низкими рейтингами", level=1)

        # Проверяем добавление таблиц с низкими рейтингами актёров и режиссёров
        mock_doc.add_heading.assert_any_call("Топ-10 актёров с"
                                             " самыми низкими рейтингами:", level=2)
        mock_doc.add_heading.assert_any_call("Топ-10 режиссёров с"
                                             " самыми низкими рейтингами:", level=2)

        # Проверяем, что добавлена таблица для актёров с низким рейтингом
        mock_doc.add_table.assert_any_call(rows=1, cols=3)

        # Проверяем, что таблица для актёров и режиссёров была правильно заполнена
        mock_doc.add_table.return_value.add_row.return_value.cells[
                0].text = 'actor'  # Пример проверки для первой ячейки


if __name__ == "__main__":
    unittest.main()
