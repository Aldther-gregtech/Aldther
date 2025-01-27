import pyodbc
from sqlalchemy import create_engine, Column, Integer, String
from sqlalchemy.orm import declarative_base, sessionmaker
import pandas as pd

# Параметры подключения к базе данных PostgreSQL
username = 'postgres'  # Замените на ваше имя пользователя
password = 'root'      # Замените на ваш пароль
database_name = 'id_table'  # Замените на имя вашей базы данных

# Создание подключения к базе данных PostgreSQL с указанием схемы
engine = create_engine(f'postgresql+psycopg2://{username}:{password}@localhost:5432/{database_name}?options=-csearch_path=Online')

# Замените 'path_to_your_database.accdb' на путь к вашей базе данных Access
access_driver = 'Microsoft Access Driver (*.mdb, *.accdb)'
db_path = r'B:\Razmechenie\2025\Razmechenie_be.mdb'

# Устанавливаем соединение с базой данных Access
connection = pyodbc.connect(f'DRIVER={access_driver};DBQ={db_path};')

# Создаем курсор для выполнения запросов
cursor = connection.cursor()

# Выполняем объединенный запрос с использованием дополнительных скобок
query = """
SELECT RazmKosguSam.[Статья] AS статья_1, RazmKosguSam.[лицевой], ORGb.[ИНН], RazmKosguSam.[ПФХД], 
RazmKosguSam.[TO], VrSam.[Статья] AS Косгу, RazmKosguSam.[КодКЗ] AS КЗ, RazmKosguSam.[КВР], RazmKosguSam.[N] AS n,
RazmKosguSam.[00], RazmKosguSam.[83], RazmKosguSam.[90], RazmKosguSam.[в работе 00, 30,39,99] AS в_работе, RazmKosguSam.[зарегистрировано 00, 30,39,99] AS зарегистрировано
FROM (RazmKosguSam
INNER JOIN VrSam ON RazmKosguSam.[N] = VrSam.[N])
INNER JOIN ORGb ON RazmKosguSam.[лицевой] = ORGb.[Лицевой счет]
"""

try:
    cursor.execute(query)

    # Получаем все строки результата
    rows = cursor.fetchall()

    # Получаем имена столбцов
    columns = [column[0] for column in cursor.description]

    # Преобразуем данные в DataFrame
    df = pd.DataFrame.from_records(rows, columns=columns)

    # Приводим значения к русским буквам
    df = df.apply(lambda x: x.str.replace('C', 'С').replace('A', 'А').replace('O', 'О') if x.dtype == "object" else x)

except pyodbc.Error as e:
    print("Ошибка при выполнении запроса:", e)

# Закрываем курсор и соединение с Access
cursor.close()
connection.close()

# Проверяем содержимое DataFrame перед изменениями
print("Содержимое DataFrame до изменений:")
print(df.head())  # Вывод первых нескольких строк DataFrame

# Вывод уникальных значений в столбце 'Косгу'
print("Уникальные значения в столбце 'Косгу':")
print(df['Косгу'].unique())

# Удаляем пробелы для точного сравнения и приводим к нижнему регистру
df['Косгу'] = df['Косгу'].str.strip()

# Обновляем значения в столбцах "Косгу" и "КВР"
mask = df['Косгу'] == '266/111'
if mask.any():
    df.loc[mask, ['Косгу', 'КВР']] = ['266', 111]
    print("Изменения были внесены:")
else:
    print("Значений для замены не найдено.")

# Проверяем содержимое DataFrame после изменений
print("Содержимое DataFrame после изменений:")
print(df.head())  # Вывод первых нескольких строк DataFrame

# Определение модели для таблицы Sam в PostgreSQL
Base = declarative_base()

class Sam(Base):
    __tablename__ = 'Sam'
    __table_args__ = {'schema': 'Online'}  # Указываем схему

    id = Column(Integer, primary_key=True)  # Добавляем первичный ключ для удобства работы с таблицей
    статья_1 = Column(String)  # Переименуйте в соответствии с вашими данными (если нужно)
    лицевой = Column(String)
    ИНН = Column(String)
    ПФХД = Column(String)
    TO = Column(String)
    Косгу = Column(String)  # Переименуйте в соответствии с вашими данными (если нужно)
    КЗ = Column(String)
    КВР = Column(Integer)
    N = Column(Integer)
    ноль = Column(String)
    _83 = Column(String)
    _90 = Column(String)
    в_работе = Column(String)  
    зарегистрировано = Column(String)  
    ТО2 = Column(String)

# Создаем сессию для работы с базой данных
Session = sessionmaker(bind=engine)
session = Session()

# Попытка загрузить данные в базу данных PostgreSQL
try:
    df.to_sql('Sam', engine, schema='Online', if_exists='replace', index=False)  # Заменяем существующую таблицу
    print("Загрузка Онлайн Самост. успешно завершена")
except Exception as e:
    print(f"Ошибка при загрузке данных в базу: {e}")

# Закрываем сессию
session.close()
