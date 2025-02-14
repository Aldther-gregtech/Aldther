import psycopg2

def merge_tables():
    try:
        # Устанавливаем соединение с базой данных PostgreSQL
        conn = psycopg2.connect(
            dbname="",
            user="",
            password="",
            host="",
            port=""
        )
        cur = conn.cursor()

        # Удаляем таблицу, если она существует
        cur.execute("DROP TABLE IF EXISTS \"verification\".\"planoviy_online\";")

        # Создаем новую таблицу
        cur.execute("""
            CREATE TABLE \"verification\".\"planoviy_online\" AS
            SELECT DISTINCT 
                g."Наименованиеучреждения", 
                g."ИНН", 
                g."Типобеспечения", 
                g."Целеваястатья", 
                g."КОСГУБУ/АУ", 
                g."ВидрасходовБУАУ", 
                g."КЗ", 
                g."СуммаПФХД", 
                COALESCE(o."00", 0) AS "00", 
                COALESCE(g."СуммаПФХД", 0) - COALESCE(o."00", 0) AS "Разница"
            FROM "Transformations"."GroupPFHD" g
            INNER JOIN "Transformations"."grouponline" o ON 
                g."Типобеспечения" = o."ТО" AND
                g."ИНН" = o."ИНН" AND
                g."Целеваястатья" = o."статья_1" AND
                g."КОСГУБУ/АУ" = o."Косгу" AND
                g."ВидрасходовБУАУ" = o."КВР" AND
                g."КЗ" = o."КЗ";
        """)

        # Подтверждаем изменения и закрываем соединение
        conn.commit()
        cur.close()
    except Exception as e:
        print(f'Произошла ошибка: {e}')
    finally:
        if conn:
            conn.close()

# Вызов функции для объединения таблиц
merge_tables()
