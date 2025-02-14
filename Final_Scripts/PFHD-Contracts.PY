import pandas as pd
from sqlalchemy import create_engine
import re
import psycopg2
import os
import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Progressbar
from datetime import datetime

def run_script():
    # Cоздание подключения к базе данных
    engine = create_engine('postgresql+psycopg2://---------/id_table')

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
                    return "СГЗ ПЛАН"
                elif row['ПФХД'] in ['1200.0', '1209.0']:
                    return "СИЦ ПЛАН"
                elif row['ПФХД'] == '1401.0':
                    return "ОМС"
                elif row['ПФХД'] == '1402.0':
                    return "ОМС"
                elif row['ПФХД'] == '1413.0':
                    return "ПД ПЛАН"
                elif row['ПФХД'] == '1500.0':
                    return "ПД ПЛАН"
                elif row['ПФХД'] == '1501.0':
                    return "ПД ПЛАН"
                elif row['ПФХД'] == '1600.0':
                    return "ПД ПЛАН"
                elif row['ПФХД'] == '1609.0':
                    return "ПД ПЛАН"
                elif row['ПФХД'] == '1702.0':
                    return "ПД ПЛАН"
                elif row['ПФХД'] == '1703.0':
                    return "ПД ПЛАН"
                elif row['ПФХД'] == '1704.0':
                    return "ПД ПЛАН"
                elif row['ПФХД'] == '1709.0':
                    return "ПД ПЛАН"
            elif row['Признакэтапа'].strip() == 'остатки':
                if row['ПФХД'] == '1100.0':
                    return "СГЗ ОСТАТКИ"
                elif row['ПФХД'] in ['1200.0', '1209.0']:
                    return "СИЦ ОСТАТКИ"
                elif row['ПФХД'] == '1401.0':
                    return "ОМС"
                elif row['ПФХД'] == '1402.0':
                    return "ОМС"
                elif row['ПФХД'] == '1413.0':
                    return "ПД ОСТАТКИ"
                elif row['ПФХД'] == '1500.0':
                    return "ПД ОСТАТКИ"
                elif row['ПФХД'] == '1501.0':
                    return "ПД ОСТАТКИ"
                elif row['ПФХД'] == '1600.0':
                    return "ПД ОСТАТКИ"
                elif row['ПФХД'] == '1609.0':
                    return "ПД ОСТАТКИ"
                elif row['ПФХД'] == '1702.0':
                    return "ПД ОСТАТКИ"
                elif row['ПФХД'] == '1703.0':
                    return "ПД ОСТАТКИ"
                elif row['ПФХД'] == '1704.0':
                    return "ПД ОСТАТКИ"
        return "Ошибка"  # В других случаях

    # Добавление нового столбца "Типобеспечения"
    df['Типобеспечения'] = df.apply(determine_type_obespecheniya, axis=1)

    # Функция для преобразования строковых значений в числовой формат
    def convert_to_numeric(value):
        if isinstance(value, str):
            value = value.replace(' ', '').replace(',', '.')
            if value == '':
                return pd.NA  # Возвращаем NA для пустых строк
        return pd.to_numeric(value, errors='coerce')

    # Преобразование столбцов "Сумма", "Сумма1", "Сумма2", "ПФХД" в числовой формат
    df['Сумма'] = df['Сумма'].apply(convert_to_numeric)
    df['Сумма1'] = df['Сумма1'].apply(convert_to_numeric)
    df['Сумма2'] = df['Сумма2'].apply(convert_to_numeric)
    df['ПФХД'] = df['ПФХД'].apply(convert_to_numeric)

    # Попытка загрузить данные в базу данных
    try:
        progress_bar["value"] = 25
        root.update_idletasks()
        df.to_sql('PFHD_TABLE', engine, schema='uploads', if_exists='replace', index=False)
        progress_bar["value"] = 50  # Обновляем прогресс бар
        print("Загрузка ПФХД успешно завершена")
        progress_bar["value"] = 75  # Обновляем прогресс бар
    except Exception as e:
       print(f"Ошибка при загрузке данных в базу: {e}")

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
    
    def determine_type_obespecheniya2(row):
        if isinstance(row['Типобеспечения'], str):
            if row['Типобеспечения'].strip() == 'СГЗ':
                return "СГЗ ПЛАН"
            elif row['Типобеспечения'].strip() == 'ОСГЗ':
                return "СГЗ ОСТАТКИ"
            elif row['Типобеспечения'].strip() == 'ОСИЦ':
                return "СИЦ ОСТАТКИ"
            elif row['Типобеспечения'].strip() == 'СИЦ':
                return "СИЦ ПЛАН"
        return "ПД"  # В других случаях

    # Добавление нового столбца "Типобеспечения"
    df['Типобеспечения2'] = df.apply(determine_type_obespecheniya2, axis=1)

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

    df.to_sql('Contracts_TABLE', engine, schema='uploads',
    if_exists = 'replace', index=False)

    progress_bar["value"] = 100  # Завершаем прогресс бар

    print("Загрузка Контрактов успешно завершена")

    engine = create_engine('postgresql+psycopg2://------/id_table')

    df = pd.read_sql_table('PFHD_TABLE', engine, schema='uploads')

    # Переименование столбцов
    df.rename(columns={
        'Учреждение': 'ИНН',
        'Сумма':'СуммаПФХД'
    }, inplace=True)

    # Агрегация данных по указанным столбцам
    grouped_df = df.groupby(['Наименованиеучреждения', 'ИНН', 'Типобеспечения', 
                            'Целеваястатья', 'ПФХД', 'КОСГУБУ/АУ', 
                            'ВидрасходовБУАУ', 'КЗ'], as_index=False).agg({
        'СуммаПФХД': 'sum',
        'Сумма1': 'sum',
        'Сумма2': 'sum'
    })


    # Запись данных в новую таблицу "ГруппПФХД" в схеме "Transformations"
    grouped_df.to_sql('GroupPFHD', engine, schema='Transformations', if_exists='replace', index=False)

    print("Таблица 'ГруппПФХД' успешно создана в схеме 'Transformations'.")

    engine = create_engine('postgresql+psycopg2://-----------/id_table')

    df = pd.read_sql_table('Contracts_TABLE', engine, schema='uploads')


    # Агрегация данных по указанным столбцам
    grouped_df = df.groupby(['Наименованиезаказчика', 'ИНН', 'Типобеспечения', 
                            'Целеваястатья', 'КодОСГУ', 
                            'Видрасходов', 'Кодзапроса', 'Типобеспечения2'], as_index=False).agg({
        'СуммаКонтракты': 'sum',
        'Сумма1': 'sum',
        'Сумма2': 'sum'
    })


    # Запись данных в новую таблицу "ГруппКонтракты" в схеме "Transformations"
    grouped_df.to_sql('GroupContracts', engine, schema='Transformations', if_exists='replace', index=False)


    print("Таблица 'ГруппКонтракты' успешно создана в схеме 'Transformations'.")

    # Настройки подключения к PostgreSQL
    connection_params = {
        "dbname": "id_table",
        "user": "-------",
        "password": "-----",
        "host": "------",
        "port": "----"
    }

    # SQL-запрос для объединения таблиц с исправленной логикой
    query = """
    DO $$
    BEGIN
        -- Создаем таблицу PFHD-CONTRACTS в схеме verification, если она не существует
        CREATE TABLE IF NOT EXISTS verification."PFHD-CONTRACTS" (
            "Наименованиеучреждения" TEXT,
            "ИНН" double precision,
            "Типобеспечения" TEXT,
            "Целеваястатья" TEXT,
            "КОСГУБУ/АУ" double precision,
            "ВидрасходовБУАУ" double precision,
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
            AND pfhd."Типобеспечения" = contracts."Типобеспечения2";

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
            AND pfhd."Типобеспечения" = contracts."Типобеспечения2"
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
            AND pfhd."Типобеспечения" = contracts."Типобеспечения2"
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

    # Cоздание подключения к базе данных
    engine = create_engine('postgresql+psycopg2://---------/id_table')

    # Запрос для получения данных из таблицы PFHD-CONTRACTS в схеме verification
    query = 'SELECT * FROM verification."PFHD-CONTRACTS"'

    # Чтение данных в DataFrame
    try:
        df = pd.read_sql(query, engine)
        
        # Укажите полный путь для сохранения файла Excel
        output_directory = r'\\Fbox\Exchange\Плановый отдел\Благов Виталий\Upload_verification\PFHD-CONTRACTS'# Можно Заменить на нужный путь
        # Получение текущей даты и времени
        current_time = datetime.now()
        formatted_time = current_time.strftime("%Y-%m-%d_%H-%M-%S")  # Форматирование даты и времени
        excel_file_name = f'PFHD_CONTRACTS_{formatted_time}.xlsx'
        excel_file_path = os.path.join(output_directory, excel_file_name)
        
        # Сохранение DataFrame в Excel файл
        df.to_excel(excel_file_path, index=False)
        
        print(f"Таблица 'pfhd-contracts' из схемы 'verification' успешно сохранена в {excel_file_path}")
    except Exception as e:
        print(f"Ошибка при выполнении запроса: {e}")
# Создание основного окна
root = tk.Tk()
root.title("Крутая Программа Витали")

# Настройка размеров окна
root.geometry("1600x1000")  # Установите размеры окна по вашему усмотрению

# Создание кнопки
button = tk.Button(root, text="Сверить ПФХД с Контрактами", command=run_script, width=40, height=10)
button.pack(pady=20, anchor='w')  # 'w' - выравнивание по левому краю

progress_bar = Progressbar(root, orient="horizontal", length=700, mode="determinate")
progress_bar.pack(side=tk.BOTTOM, pady=40)  # Добавляем индикатор прогресса

# Запуск главного цикла
root.mainloop()
