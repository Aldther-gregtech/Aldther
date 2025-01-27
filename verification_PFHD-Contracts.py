import psycopg2

# Настройки подключения к PostgreSQL
connection_params = {
    "dbname": "id_table",
    "user": "postgres",
    "password": "root",
    "host": "localhost",
    "port": "5432"
}

# SQL-запрос для объединения таблиц с исправленной логикой
query = """
DO $$
BEGIN
    -- Создаем таблицу PFHD-CONTRACTS в схеме verification, если она не существует
    CREATE TABLE IF NOT EXISTS verification."PFHD-CONTRACTS" (
        "Наименованиеучреждения" TEXT,
        "ИНН" TEXT,
        "Типобеспечения" TEXT,
        "Целеваястатья" TEXT,
        "КОСГУБУ/АУ" TEXT,
        "ВидрасходовБУАУ" TEXT,
        "КЗ" TEXT,
        "СуммаПФХД" NUMERIC,
        "СуммаКонтракты" NUMERIC,
        "Разница" NUMERIC,
        "Ошибки" TEXT
    );

    -- Очищаем таблицу перед добавлением новых данных
    TRUNCATE TABLE verification."PFHD-CONTRACTS";

    -- Вставляем строки, которые проходят условия из обеих таблиц
    INSERT INTO verification."PFHD-CONTRACTS" (
        "Наименованиеучреждения", "ИНН", "Типобеспечения", "Целеваястатья", 
        "КОСГУБУ/АУ", "ВидрасходовБУАУ", "КЗ", "СуммаПФХД", "СуммаКонтракты", 
        "Разница", "Ошибки"
    )
    SELECT
        pfhd."Наименованиеучреждения",
        pfhd."ИНН",
        pfhd."Типобеспечения",
        pfhd."Целеваястатья",
        pfhd."КОСГУБУ/АУ",
        pfhd."ВидрасходовБУАУ",
        pfhd."КЗ",
        pfhd."СуммаПФХД",
        contracts."СуммаКонтракты",
        pfhd."СуммаПФХД" - contracts."СуммаКонтракты" AS "Разница",
        NULL::TEXT AS "Ошибки"
    FROM "Transformations"."GroupPFHD" pfhd
    INNER JOIN "Transformations"."GroupContracts" contracts
    ON pfhd."ИНН" = contracts."ИНН"
        AND pfhd."Целеваястатья" = contracts."Целеваястатья"
        AND pfhd."КОСГУБУ/АУ" = contracts."КодОСГУ"
        AND pfhd."ВидрасходовБУАУ" = contracts."Видрасходов"
        AND pfhd."КЗ" = contracts."Кодзапроса"
        AND pfhd."Типобеспечения" = contracts."Типобеспечения";

    -- Вставляем строки из GroupContracts, которых нет в GroupPFHD
    INSERT INTO verification."PFHD-CONTRACTS" (
        "Наименованиеучреждения", "ИНН", "Типобеспечения", "Целеваястатья", 
        "КОСГУБУ/АУ", "ВидрасходовБУАУ", "КЗ", "СуммаПФХД", "СуммаКонтракты", 
        "Разница", "Ошибки"
    )
    SELECT
        NULL AS "Наименованиеучреждения", -- Будет заполнено последним запросом
        contracts."ИНН",
        contracts."Типобеспечения",
        contracts."Целеваястатья",
        contracts."КодОСГУ" AS "КОСГУБУ/АУ",
        contracts."Видрасходов",
        contracts."Кодзапроса" AS "КЗ",
        NULL AS "СуммаПФХД",
        contracts."СуммаКонтракты",
        NULL AS "Разница",
        'Отсутствует в ПФХД' AS "Ошибки"
    FROM "Transformations"."GroupContracts" contracts
    WHERE NOT EXISTS (
        SELECT 1
        FROM "Transformations"."GroupPFHD" pfhd
        WHERE pfhd."ИНН" = contracts."ИНН"
          AND pfhd."Целеваястатья" = contracts."Целеваястатья"
          AND pfhd."КОСГУБУ/АУ" = contracts."КодОСГУ"
          AND pfhd."ВидрасходовБУАУ" = contracts."Видрасходов"
          AND pfhd."КЗ" = contracts."Кодзапроса"
          AND pfhd."Типобеспечения" = contracts."Типобеспечения"
    );

    -- Вставляем строки из GroupPFHD, которых нет в GroupContracts
    INSERT INTO verification."PFHD-CONTRACTS" (
        "Наименованиеучреждения", "ИНН", "Типобеспечения", "Целеваястатья", 
        "КОСГУБУ/АУ", "ВидрасходовБУАУ", "КЗ", "СуммаПФХД", "СуммаКонтракты", 
        "Разница", "Ошибки"
    )
    SELECT
        pfhd."Наименованиеучреждения",
        pfhd."ИНН",
        pfhd."Типобеспечения",
        pfhd."Целеваястатья",
        pfhd."КОСГУБУ/АУ",
        pfhd."ВидрасходовБУАУ",
        pfhd."КЗ",
        pfhd."СуммаПФХД",
        NULL AS "СуммаКонтракты",
        NULL AS "Разница",
        'Отсутствует в контрактах' AS "Ошибки"
    FROM "Transformations"."GroupPFHD" pfhd
WHERE NOT EXISTS (
        SELECT 1
        FROM "Transformations"."GroupContracts" contracts
        WHERE pfhd."ИНН" = contracts."ИНН"
          AND pfhd."Целеваястатья" = contracts."Целеваястатья"
          AND pfhd."КОСГУБУ/АУ" = contracts."КодОСГУ"
          AND pfhd."ВидрасходовБУАУ" = contracts."Видрасходов"
          AND pfhd."КЗ" = contracts."Кодзапроса"
          AND pfhd."Типобеспечения" = contracts."Типобеспечения"
    );

    -- Обновляем "Наименованиеучреждения" на основе ИНН
    UPDATE verification."PFHD-CONTRACTS" target
    SET "Наименованиеучреждения" = COALESCE(
        (SELECT pfhd."Наименованиеучреждения"
         FROM "Transformations"."GroupPFHD" pfhd
         WHERE pfhd."ИНН" = target."ИНН"
         LIMIT 1),
        (SELECT contracts."Наименованиезаказчика"
         FROM "Transformations"."GroupContracts" contracts
         WHERE contracts."ИНН" = target."ИНН"
         LIMIT 1)
    )
    WHERE "Наименованиеучреждения" IS NULL;
END $$;
"""

try:
    # Подключение к базе данных
    with psycopg2.connect(**connection_params) as conn:
        with conn.cursor() as cursor:
            cursor.execute(query)
            print("Таблицы успешно объединены с обновленной логикой.")
except Exception as e:
    print("Ошибка при выполнении скрипта:", e)

