import psycopg2
import os
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO)

def merge_tables():
    try:
        # Устанавливаем соединение с базой данных PostgreSQL
        conn = psycopg2.connect(
            dbname=os.getenv("DB_NAME", "id_table"),  # Имя базы данных
            user=os.getenv("DB_USER", "postgres"),    # Имя пользователя
            password=os.getenv("DB_PASSWORD", "root"), # Пароль
            host=os.getenv("DB_HOST", "localhost"),    # Хост
            port=os.getenv("DB_PORT", "5432")          # Порт
        )
        
        with conn.cursor() as cur:
            # Создание или замена таблицы planoviy_online
            cur.execute('''
                DROP TABLE IF EXISTS "verification"."planoviy_online";
                CREATE TABLE "verification"."planoviy_online" AS
                SELECT DISTINCT
                    g."Наименованиеучреждения",
                    g."ИНН",
                    g."Типобеспечения",
                    CASE 
                        WHEN o."статья_1" IN (
                            'ПД 1413', 'ПД 1500', 'ПД 1501', 'ПД 1600', 
                            'ПД 1609', 'ПД 1702', 'ПД 1703', 'ПД 1704', 
                            'ПД 1709', 'ОПД 1413', 'ОПД 1500', 'ОПД 1501', 
                            'ОПД 1600', 'ОПД 1609', 'ОПД 1702', 'ОПД 1703', 
                            'ОПД 1704', 'ОПД 1709'
                        ) THEN '0000000000'
                        ELSE o."статья_1"
                    END AS "Целеваястатья",
                    g."ПФХД",
                    g."КОСГУБУ/АУ",
                    g."ВидрасходовБУАУ",
                    g."КЗ",
                    COALESCE(g."СуммаПФХД", 0) AS "СуммаПФХД",
                    COALESCE(o."00", 0) AS "00",
                    COALESCE(o."83", 0) AS "83",
                    COALESCE(o."90", 0) AS "90",
                    (COALESCE(g."СуммаПФХД", 0) - COALESCE(o."00", 0) - COALESCE(o."83", 0) - COALESCE(o."90", 0)) AS "Разница"
                FROM "Transformations"."GroupPFHD" g
                LEFT JOIN "Transformations"."grouponline" o ON 
                    g."Типобеспечения" = o."TO" AND 
                    g."ИНН" = o."ИНН" AND 
                    CASE 
                        WHEN o."статья_1" IN (
                            'ПД 1413', 'ПД 1500', 'ПД 1501', 'ПД 1600',
                            'ПД 1609', 'ПД 1702', 'ПД 1703', 'ПД 1704',
                            'ПД 1709', 'ОПД 1413', 'ОПД 1500', 'ОПД 1501',
                            'ОПД 1600', 'ОПД 1609', 'ОПД 1702', 'ОПД 1703',
                            'ОПД 1704', 'ОПД 1709'
                        ) THEN '0000000000'
                        ELSE o."статья_1"
                    END = g."Целеваястатья" AND 
                    g."КОСГУБУ/АУ" = o."Косгу" AND 
                    g."ВидрасходовБУАУ" = o."КВР" AND 
                    g."КЗ" = o."КЗ" AND
                    g."ПФХД" = o."ПФХД"
            ''')
            
            # Проверка наличия столбца "Ошибки"
            cur.execute('''
                DO $$
                BEGIN
                    IF NOT EXISTS (
                        SELECT 1 FROM information_schema.columns 
                        WHERE table_name='planoviy_online' AND column_name='Ошибки'
                    ) THEN
                        ALTER TABLE "verification"."planoviy_online"
                        ADD COLUMN "Ошибки" TEXT;
                    END IF;
                END $$;
            ''')
            
            # Проверка наличия столбца "Целеваястатья"
            cur.execute('''
                SELECT EXISTS (
                    SELECT column_name 
                    FROM information_schema.columns 
                    WHERE table_name='planoviy_online' AND column_name='Целеваястатья'
                );
            ''')
            exists = cur.fetchone()[0]
            
            if not exists:
                logging.error('Столбец "Целеваястатья" не существует в таблице planoviy_online.')
                return
            
            # Обновление таблицы planoviy_online с отсутствующими записями
            cur.execute('''
                INSERT INTO "verification"."planoviy_online" ("Наименованиеучреждения", "ИНН", "Типобеспечения", "Целеваястатья", "ПФХД", "КОСГУБУ/АУ", "ВидрасходовБУАУ", "КЗ", "СуммаПФХД", "00", "Ошибки")
                SELECT DISTINCT
                    k."Наименованиеучреждения",
                    o."ИНН",
                    o."TO",
                    CASE 
                        WHEN o."статья_1" IN (
                            'ПД 1413', 'ПД 1500', 'ПД 1501', 'ПД 1600',
                            'ПД 1609', 'ПД 1702', 'ПД 1703', 'ПД 1704',
                            'ПД 1709', 'ОПД 1413', 'ОПД 1500', 'ОПД 1501',
                            'ОПД 1600', 'ОПД 1609', 'ОПД 1702', 'ОПД 1703',
                            'ОПД 1704','ОПД 1709'
                        ) THEN '0000000000'
                        ELSE k."Целеваястатья"
                    END AS "статья_1",
                    o."ПФХД",
                    o."Косгу",
                    o."КВР",
                    o."КЗ",
                    NULL::double precision AS "СуммаПФХД",
                    o."00",
                    'Отсутствует в ПФХД'
                FROM "Transformations"."grouponline" o
                INNER JOIN "Transformations"."GroupPFHD" k ON
                        o."ИНН" = k."ИНН"
                WHERE NOT EXISTS (
                    SELECT 1 FROM "verification"."planoviy_online" p
                    WHERE p."ИНН" = o."ИНН" AND p."00" = o."00" AND p."КЗ" = o."КЗ" AND p."КОСГУБУ/АУ" = o."Косгу" AND p."Типобеспечения" = o."TO" AND p."Целеваястатья" = CASE 
                        WHEN o."статья_1" IN (
                            'ПД 1413', 'ПД 1500', 'ПД 1501', 'ПД 1600',
                            'ПД 1609', 'ПД 1702', 'ПД 1703', 'ПД 1704',
                            'ПД 1709', 'ОПД 1413', 'ОПД 1500', 'ОПД 1501',
                            'ОПД 1600', 'ОПД 1609', 'ОПД 1702', 'ОПД 1703',
                            'ОПД 1704','ОПД 1709'
                        ) THEN '0000000000'
                        ELSE o."статья_1"
                    END AND p."ПФХД" = o."ПФХД"
                )
            ''')
            
            # Подтверждение транзакции
            conn.commit()
            
            logging.info("Tables merged successfully.")
            
    except Exception as e:
        logging.error(f'Ошибка: {e}')
    finally:
        if conn:
            conn.close() # Закрытие соединения

# Вызов функции для выполнения операции слияния таблиц
merge_tables()
