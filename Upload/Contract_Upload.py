import pandas as pd
from sqlalchemy import create_engine
import sqlalchemy

engine = create_engine('postgresql+psycopg2://postgres:root@localhost:5432/id_table')

file_path = r'\\Fbox\Exchange\Плановый отдел\Благов Виталий\SQL\Выгрузки\КОНТРАКТЫ\Contract.xlsx'
df = pd.read_excel(file_path)

df.columns = df.columns.str.replace(' ', '')
df.columns = df.columns.str.replace('  ', '')

# Переименование столбцов
df.rename(columns={
    'Заказчик': 'ИНН',
    '1год':'СуммаКонтракты',
    '2год':'Сумма1',
    '3год':'Сумма2',
}, inplace=True)


# Функция для преобразования строковых значений в числовой формат
def convert_to_numeric(value):
    if isinstance(value, str):
        # Удаляем пробелы и заменяем запятые на точки
        value = value.replace(' ', '').replace(',', '.')
    return pd.to_numeric(value, errors='coerce')

# Преобразование столбцов "Сумма", "Сумма1", "Сумма2" в числовой формат
df['СуммаКонтракты'] = df['СуммаКонтракты'].apply(convert_to_numeric)
df['Сумма1'] = df['Сумма1'].apply(convert_to_numeric)
df['Сумма2'] = df['Сумма2'].apply(convert_to_numeric)

# Проверка на наличие NaN после преобразования
print("Количество NaN в 'СуммаКонтракты':", df['СуммаКонтракты'].isna().sum())
print("Количество NaN в 'Сумма1':", df['Сумма1'].isna().sum())
print("Количество NaN в 'Сумма2':", df['Сумма2'].isna().sum())

df.to_sql('Contracts_TABLE', engine, schema='uploads',
          if_exists='replace', index=False,
          dtype={'Целеваястатья': sqlalchemy.types.TEXT()})  # Указываем тип TEXT для столбца

print("Загрузка Контрактов успешно завершена")