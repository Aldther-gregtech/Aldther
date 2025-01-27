import pyodbc
from sqlalchemy import create_engine, Column, Integer, String
from sqlalchemy.orm import declarative_base, sessionmaker
import pandas as pd

# Параметры подключения к базе данных PostgreSQL
username = 'postgres'  # Ваше имя пользователя
password = 'root'      # Ваш пароль
database_name = 'id_table'  # Имя базы данных

# Создание подключения к базе данных PostgreSQL с указанием схемы
engine = create_engine(f'postgresql+psycopg2://{username}:{password}@localhost:5432/{database_name}?options=-csearch_path=Online')

# Путь к базе данных Access
access_driver = 'Microsoft Access Driver (*.mdb, *.accdb)'
db_path = r'B:\Razmechenie\2025\Razmechenie_be.mdb'

# Устанавливаем соединение с базой данных Access
connection = pyodbc.connect(f'DRIVER={access_driver};DBQ={db_path};')

# Создаем курсор для выполнения запросов
cursor = connection.cursor()

# Выполняем объединенный запрос
query = """
SELECT 
    RazmKosguGdou.[Статья] AS статья_1, 
    RazmKosguGdou.[лицевой], 
    ORGb.[ИНН], 
    RazmKosguGdou.[ПФХД],  
    RazmKosguGdou.[TO], 
    VrGDOU.[Статья] AS Косгу, 
    RazmKosguGdou.[КодКЗ] AS КЗ, 
    RazmKosguGdou.[КВР], 
    RazmKosguGdou.[N] AS n, 
    RazmKosguGdou.[00], 
    RazmKosguGdou.[83], 
    RazmKosguGdou.[90], 
    RazmKosguGdou.[в работе 00, 30,39,99] AS в_работе, 
    RazmKosguGdou.[зарегистрировано 00, 30,39,99] AS зарегистрировано
FROM 
    (RazmKosguGdou 
    INNER JOIN VrGDOU ON RazmKosguGdou.[N] = VrGDOU.[N]) 
    INNER JOIN ORGb ON RazmKosguGdou.[лицевой] = ORGb.[Лицевой счет]
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

# Проверяем содержимое DataFrame
print("Содержимое DataFrame:")
print(df.head())

# Определение модели для таблицы Gdou в PostgreSQL
Base = declarative_base()

class Gdou(Base):
    __tablename__ = 'Gdou'
    __table_args__ = {'schema': 'Online'}

    id = Column(Integer, primary_key=True, autoincrement=True)
    статья_1 = Column(String)
    лицевой = Column(String)
    ИНН = Column(String)
    ПФХД = Column(String)
    TO = Column(String)
    Косгу = Column(String)
    КЗ = Column(String)
    КВР = Column(Integer)
    n = Column(Integer)
    ноль = Column(String)
    _83 = Column(Integer)
    _90 = Column(Integer)
    в_работе = Column(String)
    зарегистрировано = Column(String)
    ТО2 = Column(String)

# Создаем сессию для работы с базой данных
Session = sessionmaker(bind=engine)
session = Session()

# Попытка загрузить данные в базу данных PostgreSQL
try:
    df.to_sql('Gdou', engine, schema='Online', if_exists='replace', index=False)  # Заменяем существующую таблицу
    print("Загрузка данных в таблицу 'Gdou' успешно завершена.")
except Exception as e:
    print(f"Ошибка при загрузке данных в базу: {e}")

# Закрываем сессию
session.close()
