import pandas as pd

# Фиксированные курсы валют
EXCHANGE_RATES = {
    'USD': 90,  # 1 доллар = 90 рублей
    'EUR': 100,  # 1 евро = 100 рублей
    'RUB': 1  # 1 рубль = 1 рублю
}

def convert_to_rub(amount, currency):
    """
    Конвертирует сумму в рубли.
    """
    if pd.isna(amount) or pd.isna(currency):
        return None
    return amount * EXCHANGE_RATES.get(currency, 1)

def extract_roles(persons, role_name):
    """
    Извлекает имена людей, соответствующих заданной роли.
    """
    return [p['name'] for p in persons
            if role_name in p['enProfession']] if isinstance(persons, list) else []

def prepare_data(df):
    """
    Подготавливает данные для анализа.
    """

    # Преобразуем столбцы genres и countries
    df['genres'] = (df['genres']
                    .apply(lambda x: [genre['name'] for genre in x] if isinstance(x, list) else x))
    df['countries'] = (df['countries']
                    .apply(lambda x: [country['name'] for country in x] if isinstance(x, list) else x))

    # Конвертируем валюты
    df['budget_rub'] = df.apply(lambda row: convert_to_rub(row['budget.value'],
                                                           row['budget.currency']), axis=1)

    df['fees_rub_usa'] = df.apply(lambda row: convert_to_rub(row['fees.usa.value'],
                                                             row['fees.usa.currency']), axis=1)

    df['fees_rub_russia'] = df.apply(lambda row: convert_to_rub(row['fees.russia.value'],
                                                                row['fees.russia.currency']), axis=1)

    df['fees_rub_world'] = df.apply(lambda row: convert_to_rub(row['fees.world.value'],
                                                               row['fees.world.currency']), axis=1)

    # Приведение данных к числовому типу
    numeric_columns = ['budget_rub', 'fees_rub_usa', 'fees_rub_russia', 'fees_rub_world']
    df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')

    # Извлечение актеров и режиссеров
    df['actors'] = df['persons'].apply(lambda x: extract_roles(x, 'actor'))
    df['directors'] = df['persons'].apply(lambda x: extract_roles(x, 'director'))

    # Оставляем только указанные столбцы
    columns_to_keep = [
        'id', 'name', 'year', 'genres', 'countries', 'rating.kp', 'rating.imdb',
        'votes.kp', 'votes.imdb', 'budget_rub', 'fees_rub_usa', 'fees_rub_russia', 'fees_rub_world',
        'actors', 'directors'
    ]
    df = df[[col for col in columns_to_keep if col in df.columns]]

    # Очистка данных
    df.loc[:, 'votes.kp'] = pd.to_numeric(df['votes.kp'], errors='coerce')
    df = df.dropna(subset=['name'])
    df.loc[:, 'genres'] = df['genres'].fillna('неизвестно')
    df.loc[:, 'countries'] = df['countries'].fillna('неизвестно')
    df = df[(df['rating.kp'] > 0) & (df['rating.imdb'] > 0)]

    return df
