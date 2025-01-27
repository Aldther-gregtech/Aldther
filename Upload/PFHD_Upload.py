import pandas as pd
from sqlalchemy import create_engine
import re

# Cоздание подключения к базе данных
engine = create_engine('postgresql+psycopg2://postgres:root@localhost:5432/id_table')

# Путь к файлу Excel
file_path = r'\\Fbox\Exchange\Плановый отдел\Благов Виталий\SQL\Выгрузки\ПФХД\ПФХД.xlsx'

# Загрузка данных из Excel
df = pd.read_excel(file_path)

# Удаление пробелов из заголовков столбцов
df.columns = [re.sub(r'\s+', '', col) for col in df.columns]

# Проверка наличия нужного столбца и его заполнение
if 'Целеваястатья' not in df.columns:
    if 'Целевая статья' in df.columns:
        # Заменить NaN в столбце "Целевая статья" на "0000000000"
        df['Целевая статья'] = df['Целевая статья'].fillna('0000000000')
    else:
        print("Столбец 'Целевая статья' не найден в DataFrame.")
else:
    # Если все же нужно работать со 'Целеваястатья', используйте это имя.
    df['Целеваястатья'] = df['Целеваястатья'].fillna('0000000000')

# Приведение столбца "ПФХД" к строковому типу для корректного сравнения
df['ПФХД'] = df['ПФХД'].astype(str)

# Функция для определения типа обеспечения
def determine_type_obespecheniya(row):
    if isinstance(row['Признакэтапа'], str):
        if row['Признакэтапа'].strip() == 'план':
            if row['ПФХД'] == '1100.0':
                return 'СГЗ'
            elif row['ПФХД'] in ['1200.0', '1209.0']:
                return 'СИЦ'
            elif row['ПФХД'] == '1401.0':
                return 'ОМС'
            elif row['ПФХД'] == '1402.0':
                return 'ОМС'
        elif row['Признакэтапа'].strip() == 'остатки':
            if row['ПФХД'] == '1100.0':
                return 'ОСГЗ'
            elif row['ПФХД'] in ['1200.0', '1209.0']:
                return 'ОСИЦ'
            elif row['ПФХД'] == '1401.0':
                return 'ОМС'
            elif row['ПФХД'] == '1402.0':
                return 'ОМС'
    return 'ПД'  # В других случаях

# Добавление нового столбца "Типобеспечения"
df['Типобеспечения'] = df.apply(determine_type_obespecheniya, axis=1)

# Функция для преобразования строковых значений в числовой формат
def convert_to_numeric(value):
    if isinstance(value, str):
        value = value.replace(' ', '').replace(',', '.')
        if value == '':
            return pd.NA  # Возвращаем NA для пустых строк
    return pd.to_numeric(value, errors='coerce')

# Преобразование столбцов "Сумма", "Сумма1", "Сумма2" в числовой формат
df['Сумма'] = df['Сумма'].apply(convert_to_numeric)
df['Сумма1'] = df['Сумма1'].apply(convert_to_numeric)
df['Сумма2'] = df['Сумма2'].apply(convert_to_numeric)

# Попытка загрузить данные в базу данных
try:
    df.to_sql('PFHD_TABLE', engine, schema='uploads', if_exists='replace', index=False)
    print("Загрузка ПФХД успешно завершена")
except Exception as e:
    print(f"Ошибка при загрузке данных в базу: {e}")
