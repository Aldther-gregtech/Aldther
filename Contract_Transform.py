import pandas as pd
from sqlalchemy import create_engine

engine = create_engine('postgresql+psycopg2://postgres:root@localhost:5432/id_table')

df = pd.read_sql_table('Contracts_TABLE', engine, schema='uploads')

# Агрегация данных по указанным столбцам
grouped_df = df.groupby(['Наименованиезаказчика', 'ИНН', 'Типобеспечения', 
                          'Целеваястатья', 'КодОСГУ', 
                          'Видрасходов', 'Кодзапроса'], as_index=False).agg({
    'СуммаКонтракты': 'sum',
    'Сумма1': 'sum',
    'Сумма2': 'sum'
})


# Запись данных в новую таблицу "ГруппКонтракты" в схеме "Transformations"
grouped_df.to_sql('GroupContracts', engine, schema='Transformations', if_exists='replace', index=False)


print("Таблица 'ГруппКонтракты' успешно создана в схеме 'Transformations'.")
