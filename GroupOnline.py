import pandas as pd
import psycopg2
from sqlalchemy import create_engine

def merge_tables():
    cur = None  # Initialize cur to avoid UnboundLocalError
    try:
        # Создаем SQLAlchemy engine для подключения к PostgreSQL
        engine = create_engine('postgresql+psycopg2://postgres:root@localhost:5432/id_table')

        # Чтение данных из Excel файла с использованием абсолютного пути
        institutions_df = pd.read_excel(r'\\Fbox\Exchange\Плановый отдел\Благов Виталий\SQL\Выгрузки\УЧРЕЖДЕНИЯ\учреждения.xlsx')

        # Загрузка данных из DataFrame в таблицу institutions_table в схеме uploads
        institutions_df.to_sql('institutions_table', con=engine, schema='uploads', if_exists='replace', index=False)

        # Устанавливаем соединение с базой данных PostgreSQL для выполнения других операций
        conn = engine.raw_connection()
        cur = conn.cursor()

        # SQL-запрос для создания схемы Transformations, если она не существует
        create_schema_query = '''
        CREATE SCHEMA IF NOT EXISTS "Transformations";
        '''
        
        # Выполняем запрос на создание схемы
        cur.execute(create_schema_query)

        # SQL-запрос для создания новой таблицы GroupOnline в схеме Transformations
        create_table_query = '''
        DROP TABLE IF EXISTS "Transformations".GroupOnline;
        CREATE TABLE IF NOT EXISTS "Transformations".GroupOnline (
            статья_1 TEXT,
            лицевой TEXT,
            ИНН DOUBLE PRECISION,
            ПФХД DOUBLE PRECISION,
            "TO" TEXT,
            Косгу DOUBLE PRECISION,
            КЗ TEXT,
            КВР DOUBLE PRECISION,
            "00" DOUBLE PRECISION,
            "83" DOUBLE PRECISION,
            "90" DOUBLE PRECISION,
            в_работе DOUBLE PRECISION,
            зарегистрировано DOUBLE PRECISION
        );
        '''
        
        # Выполняем запрос на создание таблицы
        cur.execute(create_table_query)

        # Добавляем новый столбец статья_2, если он не существует
        alter_table_query = '''
        ALTER TABLE "Transformations".GroupOnline ADD COLUMN IF NOT EXISTS статья_2 TEXT;
        '''
        
        # Выполняем запрос на добавление столбца
        cur.execute(alter_table_query)

        # SQL-запрос для объединения данных из трех таблиц с добавлением нового столбца статья_1
        insert_query = '''
        INSERT INTO "Transformations".GroupOnline (статья_1, лицевой, ИНН, ПФХД, "TO", Косгу, КЗ, КВР, "00", "83", "90", в_работе, зарегистрировано)
SELECT 
    CASE 
        WHEN статья_1 IN ('ПД 1413', 'ПД 1500', 'ПД 1501', 'ПД 1600', 'ПД 1609', 
                          'ПД 1702', 'ПД 1703', 'ПД 1704', 'ПД 1709',
                          'ОПД 1413', 'ОПД 1500', 'ОПД 1501', 'ОПД 1600',
                          'ОПД 1609', 'ОПД 1702', 'ОПД 1703', 'ОПД 1704',
                          'ОПД 1709') THEN 
            '0000000000'
        WHEN статья_1 = 'остатки' THEN 
            (SELECT CAST(institutions_table."Целеваястатья" AS TEXT) FROM uploads.institutions_table WHERE ИНН = CAST(institutions_table."Учреждение" AS TEXT))
        ELSE статья_1 
    END AS статья_1,
    лицевой,
    CAST(ИНН AS DOUBLE PRECISION) AS ИНН,
    CAST(ПФХД AS DOUBLE PRECISION) AS ПФХД,
    "TO",
    CAST(Косгу AS DOUBLE PRECISION) AS Косгу,
    КЗ,
    CAST(КВР AS DOUBLE PRECISION) AS КВР,
    SUM(CAST("00" AS DOUBLE PRECISION)) AS "00",
    SUM(CAST("83" AS DOUBLE PRECISION)) AS "83",
    SUM(CAST("90" AS DOUBLE PRECISION)) AS "90",
    SUM(в_работе) AS в_работе,
    SUM(зарегистрировано) AS зарегистрировано
FROM (
    SELECT REPLACE(статья_1, '_', '00') AS статья_1, лицевой, ИНН, CAST(ПФХД AS DOUBLE PRECISION) AS ПФХД, "TO", 
           CAST(Косгу AS DOUBLE PRECISION) AS Косгу,
           КЗ, 
           CAST(КВР AS DOUBLE PRECISION) AS КВР,
           CAST("00" AS DOUBLE PRECISION), 
           CAST("83" AS DOUBLE PRECISION), 
           CAST("90" AS DOUBLE PRECISION),
           в_работе, зарегистрировано 
    FROM "Online"."Gdou"
    
    UNION ALL
    
    SELECT REPLACE(статья_1, '_', '00'), лицевой , ИНН , CAST(ПФХД AS DOUBLE PRECISION) AS ПФХД ,  "TO" , 
           CAST(Косгу AS DOUBLE PRECISION),  
           КЗ , 
           CAST(КВР AS DOUBLE PRECISION),  
           CAST("00" AS DOUBLE PRECISION),
           CAST("83" AS DOUBLE PRECISION),
           CAST("90" AS DOUBLE PRECISION),
           в_работе , зарегистрировано 
    FROM "Online"."Gou"
    
    UNION ALL
    
    SELECT REPLACE(статья_1, '_', '00'), лицевой , ИНН , CAST(ПФХД AS DOUBLE PRECISION) AS ПФХД ,  "TO" , 
           CAST(Косгу AS DOUBLE PRECISION),  
           КЗ , 
           CAST(КВР AS DOUBLE PRECISION),  
           CAST("00" AS DOUBLE PRECISION),
           CAST("83" AS DOUBLE PRECISION),
           CAST("90" AS DOUBLE PRECISION),
           в_работе , зарегистрировано  
    FROM "Online"."Sam"
) AS combined_data
GROUP BY статья_1, лицевой, ИНН, ПФХД, "TO", Косгу, КЗ, КВР;
'''

        # Выполняем запрос на вставку данных
        cur.execute(insert_query)

        # Подтверждаем изменения в базе данных
        conn.commit()
        print('Успешно объединены таблицы онлайна и загружены данные из Excel')

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # Закрываем курсор и соединение только если они были инициализированы
        if cur is not None:
            cur.close()
        if conn is not None:
            conn.close()

# Запускаем функцию для объединения таблиц
merge_tables()
