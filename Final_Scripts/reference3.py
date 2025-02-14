from decimal import Decimal
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter.ttk import Combobox, Treeview
import psycopg2
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import platform
import win32print
import win32ui
import textwrap
from PIL import Image, ImageWin, ImageTk


# --- Проверяем разрядность Windows ---
if platform.architecture()[0] != "64bit":
    messagebox.showerror("Ошибка", "Требуется 64-битная версия Python для корректной работы печати.")
    exit()
# Функция для получения данных из базы

def get_data_from_db(query):
    try:
        conn = psycopg2.connect(
            dbname="id_table",
            user="----",
            password="-----",
            host="-------"
        )
        cursor = conn.cursor()
        cursor.execute(query)
        data = [row[0] for row in cursor.fetchall()]
        cursor.close()
        conn.close()
        return data
    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Ошибка при получении данных: {e}")
        return []

# Получаем данные для выпадающих списков
institution_names = get_data_from_db("SELECT DISTINCT Наименованиеучреждения FROM uploads.institutions_table")
service_names = get_data_from_db("SELECT DISTINCT Наименованиеуслуг FROM uploads.catalog_services")

# Фильтрация значений в Combobox при вводе
# Фильтрация значений в Combobox при вводе
def filter_combobox(event, combobox, data_list):
    input_text = combobox.get()  # Получаем текущий ввод
    cursor_position = combobox.index(tk.INSERT)  # Запоминаем позицию курсора

    filtered_values = [value for value in data_list if value and input_text.lower() in value.lower()]
    
    combobox["values"] = filtered_values  # Обновляем список
    combobox.icursor(cursor_position)  # Возвращаем курсор в прежнее положение

    if filtered_values:  
        combobox["values"] = filtered_values  
       
# Функция для добавления данных в таблицу
def add_to_table(partial=False):
    employee = employee_entry.get()
    institution = institution_combobox.get()
    date = date_entry.get()
    service_name = service_combobox.get()

    # Дополнительные поля
    if partial:
        field_values = [entry.get() for entry in entries]  # Все поля
    else:
        field_values = [entry.get() for entry in entries]  

    data = [employee, institution, date, service_name] + field_values

    if all(data):  
        table.insert("", "end", values=data)  
        for entry in entries:
            if isinstance(entry, Combobox):
                entry.set("")  
            else:
                entry.delete(0, tk.END)  
    else:
        messagebox.showwarning("Ошибка", "Заполните все поля перед добавлением в таблицу!")

# --- Создание таблицы, если её нет ---
def create_table_if_not_exists():
    try:
        conn = psycopg2.connect(
            dbname="id_table",
            user="postgres",
            password="root",
            host="192.168.1.186"
        )
        cursor = conn.cursor()
        
        create_table_query = '''
            CREATE TABLE IF NOT EXISTS reference.save (
                employee VARCHAR(255),  -- Добавлено поле "Сотрудник"
                institution_name VARCHAR(255),
                date DATE,
                service_name VARCHAR(255),
                kvp VARCHAR(255),
                kosgu VARCHAR(255),
                fund VARCHAR(255),
                kz VARCHAR(255),
                subdivision VARCHAR(255),
                target_article VARCHAR(255),
                security_type VARCHAR(255),
                expense_account VARCHAR(255),
                subsidy_volume VARCHAR(255)
            )
        '''
        cursor.execute(create_table_query)
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Не удалось создать таблицу: {e}")

# --- Функция для сохранения данных в базу ---
def send_and_print():
    selected_item = table.selection()
    if not selected_item:
        messagebox.showwarning("Ошибка", "Выберите строку для отправки и печати.")
        return

    selected_record = table.item(selected_item[0])["values"]

    try:
        conn = psycopg2.connect(
            dbname="id_table",
            user="postgres",
            password="root",
            host="192.168.1.186"
        )
        cursor = conn.cursor()

        # Отправляем выбранную строку в БД
        insert_query = '''
        INSERT INTO reference.financing_certificates 
        (employee, institution_name, date, service_name, kvp, kosgu, fund, kz, subdivision, target_article, 
        security_type, expense_account, subsidy_volume)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        '''
        cursor.execute(insert_query, tuple(selected_record))

        conn.commit()
        cursor.close()
        conn.close()

        # Печатаем только выбранную строку
        print_financing_certificate(selected_record)

        # Удаляем строку из таблицы выбора
        table.delete(selected_item[0])

        #messagebox.showinfo("Успех", "Сертификат отправлен и распечатан!",parent=cert_window)

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось отправить сертификат: {e}")


def create_table_without_money():
    try:
        conn = psycopg2.connect(
            dbname="id_table",
            user="-----",
            password="------",
            host="-------"
        )
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reference.save_without_money (
                employee VARCHAR(255),
                institution_name VARCHAR(255),
                date DATE,
                service_name VARCHAR(255),
                kvp VARCHAR(255),
                kosgu VARCHAR(255),
                fund VARCHAR(255),
                kz VARCHAR(255),
                subdivision VARCHAR(255),
                target_article VARCHAR(255),
                security_type VARCHAR(255),
                expense_account VARCHAR(255),
                required_subsidy_volume VARCHAR(255),
                available_subsidy_volume VARCHAR(255)
            )
        ''')

        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Не удалось создать таблицу: {e}")

def send_and_print_partial():
    selected_item = table.selection()
    if not selected_item:
        messagebox.showwarning("Ошибка", "Выберите строку для отправки и печати.")
        return

    selected_record = table.item(selected_item[0])["values"]

    try:
        conn = psycopg2.connect(
            dbname="-------",
            user="------",
            password="------",
            host="-------"
        )
        cursor = conn.cursor()

        create_table_without_money()  # Проверяем, существует ли таблица

        insert_query = '''
        INSERT INTO reference.save_without_money
        (employee, institution_name, date, service_name, kvp, kosgu, fund, kz, subdivision, target_article, 
         security_type, expense_account, required_subsidy_volume, available_subsidy_volume)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        '''
        cursor.execute(insert_query, tuple(selected_record))

        conn.commit()
        cursor.close()
        conn.close()

        table.delete(selected_item[0])
        #messagebox.showinfo("Успех", "Справка с неполным финансированием отправлена!")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось отправить справку: {e}")


# --- Функция для получения списка рабочих мест ---
def get_workplaces():
    try:
        conn = psycopg2.connect(
            dbname="id_table",
            user="postgres",
            password="root",
            host="192.168.1.186"
        )
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT Рабочееместо FROM uploads.catalog_passwords ORDER BY Рабочееместо ASC")  # Добавлена сортировка
        workplaces = [row[0] for row in cursor.fetchall()]
        cursor.close()
        conn.close()
        return workplaces
    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Не удалось получить рабочие места: {e}")
        return []


# --- Функция для проверки пароля ---
def check_password(workplace, password, root, main_root):
    try:
        conn = psycopg2.connect(
            dbname="id_table",
            user="postgres",
            password="root",
            host="192.168.1.186"
        )
        cursor = conn.cursor()
        cursor.execute("SELECT Пароль FROM uploads.catalog_passwords WHERE Рабочееместо = %s", (workplace,))
        result = cursor.fetchone()
        cursor.close()
        conn.close()

        if result and str(result[0]) == password:  # Преобразуем BIGINT в строку перед сравнением
            root.destroy()  # Закрываем окно авторизации
            create_financing_certificate(main_root)  # Передаем root в create_financing_certificate()
        else:
            messagebox.showerror("Ошибка", "Неверный пароль!")

    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Ошибка при проверке пароля: {e}")
# --- Функция для ввода логина и пароля ---
def login_window(root):
    auth_window = tk.Toplevel(root)
    auth_window.title("Авторизация")
    auth_window.geometry("400x250")

    tk.Label(auth_window, text="Выберите рабочее место:", font=("Arial", 12)).pack(pady=5)
    workplaces = get_workplaces()
    workplace_combobox = Combobox(auth_window, font=("Arial", 12), width=30)
    workplace_combobox["values"] = workplaces
    workplace_combobox.pack(pady=5)

    tk.Label(auth_window, text="Введите пароль:", font=("Arial", 12)).pack(pady=5)
    password_entry = tk.Entry(auth_window, font=("Arial", 12), width=30, show="*")
    password_entry.pack(pady=5)

    tk.Button(auth_window, text="Войти", font=("Arial", 12),
              command=lambda: check_password(workplace_combobox.get(), password_entry.get(), auth_window, root)).pack(pady=10)  


    auth_window = tk.Toplevel(root)
    auth_window.title("Авторизация")
    auth_window.geometry("400x250")

    tk.Label(auth_window, text="Выберите рабочее место:", font=("Arial", 12)).pack(pady=5)
    workplaces = get_workplaces()
    workplace_combobox = Combobox(auth_window, font=("Arial", 12), width=30)
    workplace_combobox["values"] = workplaces
    workplace_combobox.pack(pady=5)

    tk.Label(auth_window, text="Введите пароль:", font=("Arial", 12)).pack(pady=5)
    password_entry = tk.Entry(auth_window, font=("Arial", 12), width=30, show="*")
    password_entry.pack(pady=5)

    tk.Button(auth_window, text="Войти", font=("Arial", 12),
              command=lambda: check_password_partial(workplace_combobox.get(), password_entry.get(), auth_window, root)).pack(pady=10)  
    

# --- Функция для печати справки финансирования ---
# --- Функция для печати последней записи ---
def print_financing_certificate(selected_record=None): 
    try: 
        # Если передана конкретная строка, используем её, иначе берём последнюю из таблицы
        if selected_record:
            last_row = [str(item) for item in selected_record]  # Полное сохранение ведущих нулей
        else:
            if not table.get_children(): 
                messagebox.showwarning("Ошибка", "Таблица пуста! Сначала добавьте запись.") 
                return
            last_item = table.get_children()[-1] 
            last_row = [str(item) for item in table.item(last_item)["values"]]  

        printer_name = win32print.GetDefaultPrinter() 
        hprinter = win32print.OpenPrinter(printer_name) 
        win32print.ClosePrinter(hprinter) 

        hdc = win32ui.CreateDC() 
        hdc.CreatePrinterDC(printer_name) 

        hdc.StartDoc("Справка финансирования") 
        hdc.StartPage() 

        # Шрифт заголовка 
        title_font = win32ui.CreateFont({"name": "Arial", "height": 120, "weight": 700}) 
        hdc.SelectObject(title_font) 

        hdc.TextOut(100, 200, "Справка финансирования") 

        institution_name = last_row[1] 
        date_info = last_row[2] 

        institution_y_start = 400 
        hdc.TextOut(100, institution_y_start, "Наименование учреждения:") 

        wrapped_lines = textwrap.wrap(institution_name, width=50) 
        for i, line in enumerate(wrapped_lines): 
            hdc.TextOut(1600, institution_y_start + (i * 100), line)  

        hdc.TextOut(100, institution_y_start + 100 * len(wrapped_lines) + 100, f"Дата: {date_info}") 

        # Шрифт таблицы 
        table_font = win32ui.CreateFont({"name": "Arial", "height": 80, "weight": 400}) 
        hdc.SelectObject(table_font) 

        headers = ["Название услуги", "КОСГУ", "КВР", "Фонд", "КЗ", "Подраздел", 
                   "Целевая статья", "Тип обеспечения", "Счет расходов", "Объем субсидии (руб.)"] 

        x_start = 100 
        column_widths = [800, 300, 300, 300, 200, 450, 700, 500, 500, 700] 
        row_height = 100 
        header_y_start = institution_y_start + 100 * len(wrapped_lines) + 300  

        for i, header in enumerate(headers): 
            max_chars = get_max_chars_per_line(column_widths[i], hdc) 
            wrapped_header_lines = textwrap.wrap(header, width=max_chars)[:2] 
            for j in range(len(wrapped_header_lines)): 
                hdc.TextOut(x_start + sum(column_widths[:i]), header_y_start + j * row_height, wrapped_header_lines[j]) 

        value_start_y = header_y_start + row_height * 2 
        for i, value in enumerate(last_row[3:]): 
            max_chars = get_max_chars_per_line(column_widths[i], hdc) 
            wrapped_value_lines = textwrap.wrap(value, width=max_chars) 
            for j, line in enumerate(wrapped_value_lines): 
                hdc.TextOut(x_start + sum(column_widths[:i]), value_start_y + (j * row_height), line) 

        hdc.TextOut(100, value_start_y + 2000, f"Сотрудник: {last_row[0]}") 
        hdc.TextOut(1200, value_start_y + 2000, "Подпись: __") 

        hdc.EndPage() 
        hdc.EndDoc() 

    except Exception as e: 
        messagebox.showerror("Ошибка печати", f"Не удалось отправить документ на печать: {e}")

def get_max_chars_per_line(width, hdc):
    text_extent = hdc.GetTextExtent("A")  # Получаем размер текста 'A'
    return width // text_extent[0]  # Извлекаем ширину из кортежа

def on_add_rnk_button_click():
        # Открытие окна для выбора записи и добавления РНК
        add_rnk_window()

def apply_styles(root):
        style = ttk.Style()
        style.theme_use("clam")  # Более современный стиль

        # Стили кнопок
        style.configure("TButton", font=("Arial", 14), padding=10, background="#0078D7", foreground="white")
        style.map("TButton", background=[("active", "#005A9E")])

        # Стили Combobox
        style.configure("TCombobox", font=("Arial", 12), padding=5)

        # Стили таблицы (Treeview)
        style.configure("Treeview", font=("Arial", 12), rowheight=25)
        style.configure("Treeview.Heading", font=("Arial", 14, "bold"), background="#0078D7", foreground="white")

# Главный экран
def main_screen():
    root = tk.Tk()
    root.title("Главный экран")
    root.geometry("1920x1080")
    root.state('zoomed')
    root.resizable(True, True)
    root.configure(bg="#F0F0F0")

    def apply_styles(root):
        style = ttk.Style()
        style.theme_use("clam")  # More modern style

        # Button styles
        style.configure("TButton", font=("Arial", 14), padding=10, background="#0078D7", foreground="white")
        style.map("TButton", background=[("active", "#005A9E")])

        # Combobox styles
        style.configure("TCombobox", font=("Arial", 12), padding=5)

        # Table (Treeview) styles
        style.configure("Treeview", font=("Arial", 12), rowheight=25)
        style.configure("Treeview.Heading", font=("Arial", 14, "bold"), background="#0078D7", foreground="white")

    apply_styles(root)

    # --- Load background ---
    bg_image = Image.open("4c6ed832-3b08-5318-b346-857c5905b1be.jpg")
    bg_image = bg_image.resize((1920, 1080), Image.LANCZOS)
    bg_photo = ImageTk.PhotoImage(bg_image)
    bg_label = tk.Label(root, image=bg_photo)
    bg_label.place(relwidth=1, relheight=1)

    # --- Title ---
    title_label = tk.Label(root, text="Справка финансирования", font=("Arial", 36, "bold"), fg="white", bg="#003366")
    title_label.pack(pady=20)

    # --- Frame for buttons ---
    frame = tk.Frame(root, bg="#F0F0F0")
    frame.pack(pady=50)

    btn_width = 40
    ttk.Button(frame, text="Создать справку финансирования", command=lambda: login_window(root), width=btn_width).pack(pady=10)
    ttk.Button(frame, text="Создать справку с неполным финансированием", command=lambda: login_window_partial(root), width=btn_width).pack(pady=10)
    ttk.Button(frame, text="Внести РНК", command=on_add_rnk_button_click, width=btn_width).pack(pady=10)
    ttk.Button(frame, text="Удалить справку", command=delete_certificate, width=btn_width).pack(pady=10)

    root.mainloop()


# Окно справки финансирования
def create_financing_certificate(root):
    global institution_combobox, date_entry, service_combobox, employee_entry, zoom_level, cert_window, canvas, scrollable_frame
   
    zoom_level = 1.0  # Начальный масштаб

    cert_window = tk.Toplevel(root)
    cert_window.title("Справка о финансировании")
    cert_window.geometry("1920x1080")
    cert_window.state('zoomed')  # Разворачиваем на весь экран

    # Создаем Canvas для прокрутки
    canvas = tk.Canvas(cert_window)
    scrollbar_y = tk.Scrollbar(cert_window, orient="vertical", command=canvas.yview)
    scrollbar_x = tk.Scrollbar(cert_window, orient="horizontal", command=canvas.xview)
    
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar_y.grid(row=0, column=2, sticky="ns")
    scrollbar_x.grid(row=23, column=0, sticky="ew")

    cert_window.grid_rowconfigure(0, weight=1)
    cert_window.grid_columnconfigure(0, weight=1)

    

    # --- Функция изменения масштаба ---
    def adjust_zoom(event):
        global zoom_level

        if event.state & 0x4:  # Проверяем, зажат ли Ctrl
            if event.delta > 0:
                zoom_level *= 1.1  # Увеличение масштаба
            else:
                zoom_level /= 1.1  # Уменьшение масштаба
            update_ui_scale()

    cert_window.bind("<MouseWheel>", adjust_zoom)  # Масштабирование через Ctrl + колесо

    # --- Функция обновления UI при изменении масштаба ---
    def update_ui_scale():
        scale = max(0.1, min(zoom_level, 3.0))  # Ограничение масштаба (от 50% до 200%)

        new_font_size = int(14 * scale)
        new_entry_width = int(30 * scale)

        for widget in scrollable_frame.winfo_children():
            if isinstance(widget, tk.Label):
                widget.config(font=("Arial", new_font_size))
            elif isinstance(widget, tk.Entry) or isinstance(widget, Combobox):
                widget.config(font=("Arial", new_font_size), width=new_entry_width)
            elif isinstance(widget, ttk.Button):
                widget.config(style="Zoom.TButton")

        # Обновляем таблицу
        style.configure("Treeview", font=("Arial", new_font_size), rowheight=int(25 * scale))
        style.configure("Treeview.Heading", font=("Arial", new_font_size, "bold"))

        # Обновляем scrollregion
        canvas.configure(scrollregion=canvas.bbox("all"))

    # --- Стили для кнопок ---
    style = ttk.Style()
    style.configure("Zoom.TButton", font=("Arial", 14), padding=10)


# Поля "Наименование учреждения" и "Дата"
    tk.Label(scrollable_frame, text="Наименование учреждения:", font=("Arial", 14)).grid(row=0, column=0, padx=10, pady=5, sticky="w")
    institution_combobox = Combobox(scrollable_frame, font=("Arial", 14), width=50)
    institution_combobox["values"] = institution_names
    institution_combobox.grid(row=0, column=1, padx=10, pady=5)
    institution_combobox.bind("<KeyRelease>", lambda event: filter_combobox(event, institution_combobox, institution_names))

    tk.Label(scrollable_frame, text="Дата:", font=("Arial", 14)).grid(row=1, column=0, padx=10, pady=5, sticky="w")
    date_entry = tk.Entry(scrollable_frame, font=("Arial", 14), width=20)
    date_entry.grid(row=1, column=1, padx=10, pady=5)

# Поле "Название услуги"
    tk.Label(scrollable_frame, text="Название услуги:", font=("Arial", 14)).grid(row=2, column=0, padx=10, pady=5, sticky="w")
    service_combobox = Combobox(scrollable_frame, font=("Arial", 14), width=90)
    service_combobox["values"] = service_names
    service_combobox.grid(row=2, column=1, padx=10, pady=5, sticky="w") 
    service_combobox.bind("<KeyRelease>", lambda event: filter_combobox(event, service_combobox, service_names))

# --- Добавляем поле "Сотрудник" ---
    tk.Label(scrollable_frame, text="Сотрудник:", font=("Arial", 14)).grid(row=3, column=0, padx=10, pady=5, sticky="w")

    employee_entry = tk.Entry(scrollable_frame, font=("Arial", 14), width=50)
    employee_entry.grid(row=3, column=1, padx=10, pady=5)

# Поля для остальных данных
    fields = ["КОСГУ", "КВР", "Фонд", "КЗ", "Подраздел", "Целевая статья", "Тип обеспечения", "Счет расходов", "Объем субсидии (руб.)"]
    combobox_options = {
        "КВР": ["111", "112", "113","119","243","244","247","321","323","340","851","852","853"],
        "КОСГУ": ["211","212","213","214","221","222","223","224","225","226","227","228","229","262","263","264","265","266","291","292","293","295","296","297","299","310","320","341","342","343","344","345","346","347","349"],
        "Фонд": ["'00","'83","'EВ","'81","'Ю6","'90"],
        "КЗ": [ "'з","'5","'е","'р","'0"],
        "Подраздел": ["'0701","'0702","'0703","'0705","'0707","'0709","'1003","'1101","'1102","'1103","'1004"],
        "Целевая статья": ["'021EВ51790","'021Ю650500","'021Ю651790","'021Ю653030","'0250020010","'0250020030","'0250020060","'0250020070","'0250020080","'0250020090","'0250020300","'0250020370","'0250020400","'0250020430","'0250020440","'0250020620","'0250020680","'0250020940","'0250021100","'0250021450","'02500R0500","'02500R3030","'0350040240","'0350040650","'03500R3040","'0450045010","'0450045020","'0450045060","'0450045090","'0450045390","'0850070890","'0950083620","'1750078010","'1750078560","'1750078650","'1750079450","'9950003450"],
        "Тип обеспечения": ["СГЗ", "ОСГЗ", "СИЦ", "ОСИЦ", "ПД", "ОПД"],
        "Счет расходов": ["550360","543148","552578","552555","552556","552524","543131","543134","543132","543133","543135","543138","543136","543137","543139","543142","543140","543141","543145","543146","543147","546940","553286","543149","543150","543151","553062","553086","546939","543152","543153","543154","543156","543157","543158","547692","547700","552980","548038","546941","552843","552846","552399","550695","546942","543143","543144","550357","546943","543159","543160","543161","549010","546944","546947","546945","546946","547671","547004","547005","547006","552904","546766","546767","546768","547942","543155","546769","551361","551362"]
    }
    # Глобальные переменные для нужных Combobox
    global fund_combobox, kz_combobox, subdivision_combobox, target_article_combobox, expense_account_combobox, entries

    entries = []  # Список для хранения всех полей

    for i, field in enumerate(fields):
        tk.Label(scrollable_frame, text=field + ":", font=("Arial", 14)).grid(row=i+4, column=0, padx=10, pady=5, sticky="w")
        
        if field in combobox_options:
            combobox = Combobox(scrollable_frame, font=("Arial", 14), width=30)
            combobox["values"] = combobox_options[field]
            combobox.grid(row=i+4, column=1, padx=10, pady=5, sticky="w")
            
            # Присваиваем переменные
            if field == "Фонд":
                fund_combobox = combobox
            elif field == "КЗ":
                kz_combobox = combobox
            elif field == "Подраздел":
                subdivision_combobox = combobox
            elif field == "Целевая статья":
                target_article_combobox = combobox
            elif field == "Счет расходов":
                expense_account_combobox = combobox
            
            entries.append(combobox)  # Добавляем combobox в entries
        else:
            entry = tk.Entry(scrollable_frame, font=("Arial", 14), width=30)
            entry.grid(row=i+4, column=1, padx=10, pady=5, sticky="w")
            entries.append(entry)


    # Словарь правил для автоматического заполнения "Счет расходов"
    expense_account_rules = {
        ("'021EВ51790", "'0702", "'з", "'83"): "550360",
        ("'021EВ51790", "'0702", "'з", "'EВ"): "543148",
        ("'021Ю650500", "'0709", "'5", "'81"): "552578",
        ("'021Ю651790", "'0702", "'з", "'83"): "552555",
        ("'021Ю651790", "'0702", "'з", "'Ю6"): "552556",
        ("'021Ю653030", "'0702", "'5", "'81"): "552524",
        ("'0250020010", "'0701", "'5", "'00"): "543131",
        ("'0250020010", "'0701", "'е", "'00"): "543134",
        ("'0250020010", "'0701", "'з", "'00"): "543132",
        ("'0250020010", "'0701", "'р", "'00"): "543133",
        ("'0250020030", "'0702", "'5", "'00"): "543135",
        ("'0250020030", "'0702", "'е", "'00"): "543138",
        ("'0250020030", "'0702", "'з", "'00"): "543136",
        ("'0250020030", "'0702", "'р", "'00"): "543137",
        ("'0250020060", "'0702", "'5", "'00"): "543139",
        ("'0250020060", "'0702", "'е", "'00"): "543142",
        ("'0250020060", "'0702", "'з", "'00"): "543140",
        ("'0250020060", "'0702", "'р", "'00"): "543141",
        ("'0250020070", "'0702", "'5", "'00"): "543145",
        ("'0250020070", "'0702", "'з", "'00"): "543146",
        ("'0250020070", "'0702", "'р", "'00"): "543147",
        ("'0250020080", "'0702", "'5", "'00"): "546940",
        ("'0250020090", "'0702", "'5", "'00"): "553286",
        ("'0250020300", "'0703", "'5", "'00"): "543149",
        ("'0250020300", "'0703", "'з", "'00"): "543150",
        ("'0250020300", "'0703", "'р", "'00"): "543151",
        ("'0250020370", "'0709", "'5", "'00"): "553062",
        ("'0250020370", "'0709", "'5", "'00"): "553086",
        ("'0250020400", "'0701", "'5", "'00"): "546939",
        ("'0250020430", "'0709", "'5", "'00"): "543152",
        ("'0250020430", "'0709", "'з", "'00"): "543153",
        ("'0250020430", "'0709", "'р", "'00"): "543154",
        ("'0250020440", "'0709", "'5", "'00"): "543156",
        ("'0250020440", "'0709", "'з", "'00"): "543157",
        ("'0250020440", "'0709", "'р", "'00"): "543158",
        ("'0250020620", "'0705", "'0", "'00"): "547692",
        ("'0250020680", "'0701", "'5", "'00"): "547700",
        ("'0250020940", "'0702", "'5", "'00"): "552980",
        ("'0250020940", "'0709", "'5", "'00"): "548038",
        ("'0250021100", "'0703", "'5", "'00"): "546941",
        ("'0250021450", "'0701", "'5", "'00"): "552843",
        ("'0250021450", "'0702", "'5", "'00"): "552846",
        ("'02500R0500", "'0709", "'5", "'81"): "552399",
        ("'02500R3030", "'0702", "'5", "'81"): "550695",
        ("'0350040240", "'1003", "'0", "'00"): "546942",
        ("'0350040650", "'0702", "'е", "'00"): "543143",
        ("'03500R3040", "'0702", "'е", "'00"): "543144",
        ("'03500R3040", "'0702", "'е", "'83"): "550357",
        ("'0450045010", "'1102", "'0", "'00"): "546943",
        ("'0450045020", "'1101", "'0", "'00"): "543159",
        ("'0450045020", "'1101", "'з", "'00"): "543160",
        ("'0450045020", "'1101", "'р", "'00"): "543161",
        ("'0450045060", "'1101", "'0", "'00"): "549010",
        ("'0450045090", "'1103", "'5", "'00"): "546944",
        ("'0450045090", "'1103", "'е", "'00"): "546947",
        ("'0450045090", "'1103", "'з", "'00"): "546945",
        ("'0450045090", "'1103", "'р", "'00"): "546946",
        ("'0450045390", "'1103", "'0", "'00"): "547671",
        ("'0850070890", "'0703", "'5", "'00"): "547004",
        ("'0850070890", "'0703", "'з", "'00"): "547005",
        ("'0850070890", "'0703", "'р", "'00"): "547006",
        ("'0950083620", "'1004", "'5", "'00"): "552904",
        ("'1750078010", "'0707", "'5", "'00"): "546766",
        ("'1750078010", "'0707", "'з", "'00"): "546767",
        ("'1750078010", "'0707", "'р", "'00"): "546768",
        ("'1750078560", "'0707", "'5", "'00"): "547942",
        ("'1750078650", "'0709", "'5", "'00"): "543155",
        ("'1750079450", "'0707", "'0", "'00"): "546769",
        ("'9950003450", "'0709", "'0", "'00"): "551361",
        ("'9950003450", "'0709", "'0", "'00"): "551362",
    }

    global table
        # Функция для обновления "Счет расходов"
    def update_expense_account(event=None):
        if not (target_article_combobox and subdivision_combobox and kz_combobox and fund_combobox and expense_account_combobox):
            print("Ошибка: Один из Combobox не найден!")
            return

        key = (target_article_combobox.get().strip(), 
            subdivision_combobox.get().strip(), 
            kz_combobox.get().strip(), 
            fund_combobox.get().strip())  

        print(f"Выбранные значения: {key}")  # Проверяем, что приходит в key
        
        if key in expense_account_rules:
            expense_account_combobox.set(expense_account_rules[key])  # Заполняем "Счет расходов"
            print(f"Найдено соответствие: {expense_account_rules[key]}")  
        else:
            expense_account_combobox.set("")  
            print("Совпадение не найдено.")  

    # Привязываем функцию только к нужным Combobox
    if target_article_combobox and subdivision_combobox and kz_combobox and fund_combobox:
        target_article_combobox.bind("<<ComboboxSelected>>", update_expense_account)
        subdivision_combobox.bind("<<ComboboxSelected>>", update_expense_account)
        kz_combobox.bind("<<ComboboxSelected>>", update_expense_account)
        fund_combobox.bind("<<ComboboxSelected>>", update_expense_account)

    
    # Таблица
    table = Treeview(cert_window, columns=["Сотрудник", "Наименование учреждения", "Дата", "Название услуги"] + fields, show="headings", height=10)
    
    table.heading("Сотрудник", text="Сотрудник")  
    table.column("Сотрудник", width=150, anchor="w")
    for column in ["Наименование учреждения", "Дата", "Название услуги"] + fields:
        table.heading(column, text=column)
        table.column(column, width=150, anchor="center")
    table.grid(row=len(fields) + 4, column=0, columnspan=2, padx=10, pady=10)

    def send_and_print():
        selected_item = table.selection()
        if not selected_item:
            messagebox.showwarning("Ошибка", "Выберите строку для отправки и печати.")
            return

        selected_record = table.item(selected_item[0])["values"]

        try:
            conn = psycopg2.connect(
                dbname="id_table",
                user="postgres",
                password="root",
                host="192.168.1.186"
            )
            cursor = conn.cursor()

            # Отправляем выбранную строку в БД
            insert_query = '''
            INSERT INTO reference.save (
                employee, institution_name, date, service_name, kvp, kosgu, fund,
                kz, subdivision, target_article, security_type,
                expense_account, subsidy_volume
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            '''
            cursor.execute(insert_query, tuple(selected_record))

            conn.commit()
            cursor.close()
            conn.close()

            # Печатаем только выбранную строку
            print_financing_certificate(selected_record)

            # Удаляем строку из таблицы выбора
            table.delete(selected_item[0])

            #messagebox.showinfo("Успех", "Сертификат отправлен и распечатан!")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось отправить сертификат: {e}")

    # --- Фрейм для кнопок под таблицей ---
    button_frame = tk.Frame(cert_window)
    button_frame.grid(row=20, column=0, columnspan=2, pady=10)
    # Кнопки
    ttk.Button(button_frame, text="Добавить в таблицу", command=add_to_table).pack(side=tk.LEFT, padx=10)
    ttk.Button(button_frame, text="Отправить и распечатать", command=send_and_print).pack(side=tk.LEFT, padx=10)

    # --- Добавляем кнопку "Печать" между кнопками ---
# print_button = tk.Button(cert_window, text="Печать", command=print_financing_certificate, font=("Arial", 14))
#print_button.grid(row=len(fields) + 6, column=0, pady=20)
    # --- Обновляем UI ---
    update_ui_scale()

    cert_window.mainloop()

# Окно справки финансирования без денег
# --- Функция для ввода логина и пароля для полной справки ---
def login_window(root):
    auth_window = tk.Toplevel(root)
    auth_window.title("Авторизация")
    auth_window.geometry("400x250")

    tk.Label(auth_window, text="Выберите рабочее место:", font=("Arial", 12)).pack(pady=5)
    workplaces = get_workplaces()
    workplace_combobox = Combobox(auth_window, font=("Arial", 12), width=30)
    workplace_combobox["values"] = workplaces
    workplace_combobox.pack(pady=5)

    tk.Label(auth_window, text="Введите пароль:", font=("Arial", 12)).pack(pady=5)
    password_entry = tk.Entry(auth_window, font=("Arial", 12), width=30, show="*")
    password_entry.pack(pady=5)

    tk.Button(auth_window, text="Войти", font=("Arial", 12),
              command=lambda: check_password_full(workplace_combobox.get(), password_entry.get(), auth_window, root)).pack(pady=10)

# --- Функция для проверки пароля для полной справки ---
def check_password_full(workplace, password, root, main_root):
    try:
        conn = psycopg2.connect(
            dbname="-----",
            user="-----",
            password="-----",
            host="------"
        )
        cursor = conn.cursor()
        cursor.execute("SELECT Пароль FROM uploads.catalog_passwords WHERE Рабочееместо = %s", (workplace,))
        result = cursor.fetchone()
        cursor.close()
        conn.close()

        if result and str(result[0]) == password:
            root.destroy()
            create_financing_certificate(main_root)  # Создание окна полной справки
        else:
            messagebox.showerror("Ошибка", "Неверный пароль!")
    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Ошибка при проверке пароля: {e}")

# --- Функция для ввода логина и пароля для частичной справки ---
def login_window_partial(root):
    auth_window = tk.Toplevel(root)
    auth_window.title("Авторизация")
    auth_window.geometry("400x250")

    tk.Label(auth_window, text="Выберите рабочее место:", font=("Arial", 12)).pack(pady=5)
    workplaces = get_workplaces()
    workplace_combobox = Combobox(auth_window, font=("Arial", 12), width=30)
    workplace_combobox["values"] = workplaces
    workplace_combobox.pack(pady=5)

    tk.Label(auth_window, text="Введите пароль:", font=("Arial", 12)).pack(pady=5)
    password_entry = tk.Entry(auth_window, font=("Arial", 12), width=30, show="*")
    password_entry.pack(pady=5)

    tk.Button(auth_window, text="Войти", font=("Arial", 12),
              command=lambda: check_password_partial(workplace_combobox.get(), password_entry.get(), auth_window, root)).pack(pady=10)

# --- Функция для проверки пароля для частичной справки ---
def check_password_partial(workplace, password, root, main_root):
    try:
        conn = psycopg2.connect(
            dbname="-----",
            user="-----",
            password="-----",
            host="-----"
        )
        cursor = conn.cursor()
        cursor.execute("SELECT Пароль FROM uploads.catalog_passwords WHERE Рабочееместо = %s", (workplace,))
        result = cursor.fetchone()
        cursor.close()
        conn.close()

        if result and str(result[0]) == password:
            root.destroy()
            create_partial_financing_certificate(main_root)  # Создание окна частичной справки
        else:
            messagebox.showerror("Ошибка", "Неверный пароль!")
    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Ошибка при проверке пароля: {e}")


# --- Окно справки с неполным финансированием ---
def create_partial_financing_certificate(root):
    global institution_combobox, date_entry, service_combobox, employee_entry, entries, table, zoom_level_partial, cert_window_partial, canvas_partial, scrollable_frame_partial

    zoom_level_partial = 1.0  # Начальный масштаб
    
    # Создаем новое окно
    cert_window = tk.Toplevel(root)
    cert_window.title("Справка о финансировании (неполная)")
    cert_window.geometry("1920x1080")
    cert_window.state('zoomed')

    # --- Создаем Canvas с вертикальной прокруткой ---
    canvas = tk.Canvas(cert_window)
    scrollbar_y = tk.Scrollbar(cert_window, orient="vertical", command=canvas.yview)
    scrollbar_x = tk.Scrollbar(cert_window, orient="horizontal", command=canvas.xview)
    
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar_y.grid(row=0, column=2, sticky="ns")
    scrollbar_x.grid(row=23, column=0, sticky="ew")

    cert_window.grid_rowconfigure(0, weight=1)
    cert_window.grid_columnconfigure(0, weight=1)

    # --- Функция изменения масштаба ---
    def adjust_zoom_partial(event):
        global zoom_level_partial
        if event.state & 0x4:  # Проверяем, зажат ли Ctrl
            if event.delta > 0:
                zoom_level_partial *= 1.1
            else:
                zoom_level_partial /= 1.1
            update_ui_scale_partial()

    cert_window.bind("<MouseWheel>", adjust_zoom_partial)

    # --- Функция обновления UI при изменении масштаба ---
    def update_ui_scale_partial():
        scale = max(0.1, min(zoom_level_partial, 3.0))  # Ограничение масштаба (от 50% до 200%)

        new_font_size = int(14 * scale)
        new_entry_width = int(30 * scale)

        for widget in scrollable_frame.winfo_children():
            if isinstance(widget, tk.Label):
                widget.config(font=("Arial", new_font_size))
            elif isinstance(widget, tk.Entry) or isinstance(widget, Combobox):
                widget.config(font=("Arial", new_font_size), width=new_entry_width)
            elif isinstance(widget, ttk.Button):
                widget.config(style="Zoom.TButton")

        # Обновляем таблицу
        style.configure("Treeview", font=("Arial", new_font_size), rowheight=int(25 * scale))
        style.configure("Treeview.Heading", font=("Arial", new_font_size, "bold"))

        # Обновляем scrollregion
        canvas.configure(scrollregion=canvas.bbox("all"))

    # --- Стили для кнопок ---
    style = ttk.Style()
    style.configure("Zoom.TButton", font=("Arial", 14), padding=10)
    # Поля "Наименование учреждения"
    tk.Label(scrollable_frame, text="Наименование учреждения:", font=("Arial", 14)).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    institution_combobox = Combobox(scrollable_frame, font=("Arial", 14), width=30)  # Уменьшена ширина
    institution_combobox["values"] = institution_names
    institution_combobox.grid(row=0, column=1, padx=5, pady=5)
    institution_combobox.bind("<KeyRelease>", lambda event: filter_combobox(event, institution_combobox, institution_names))

    # Поле "Дата"
    tk.Label(scrollable_frame, text="Дата:", font=("Arial", 14)).grid(row=1, column=0, padx=5, pady=5, sticky="w")
    date_entry = tk.Entry(scrollable_frame, font=("Arial", 14), width=30)  # Уменьшена ширина
    date_entry.grid(row=1, column=1, padx=5, pady=5)

    # Поле "Сотрудник"
    tk.Label(scrollable_frame, text="Сотрудник:", font=("Arial", 14)).grid(row=2, column=0, padx=5, pady=5, sticky="w")
    employee_entry = tk.Entry(scrollable_frame, font=("Arial", 14), width=30)  # Уменьшена ширина
    employee_entry.grid(row=2, column=1, padx=5, pady=5)

    # Поле "Наименование услуги"
    tk.Label(scrollable_frame, text="Наименование услуги:", font=("Arial", 14)).grid(row=3, column=0, padx=5, pady=5, sticky="w")
    service_combobox = Combobox(scrollable_frame, font=("Arial", 14), width=30)  # Уменьшена ширина
    service_combobox["values"] = service_names
    service_combobox.grid(row=3, column=1, padx=5, pady=5)
    service_combobox.bind("<KeyRelease>", lambda event: filter_combobox(event, service_combobox, service_names))


    # --- Настройки для Combobox ---
    fields = ["КВР", "КОСГУ", "Фонд", "КЗ", "Подраздел", "Целевая статья", "Тип обеспечения", "Счет расходов", "Доступный объем субсидии", "Необходимый объем к перераспределению"]
    combobox_options = {
        "КВР": ["111", "112", "113","119","243","244","247","321","323","340","851","852","853"],
        "КОСГУ": ["211","212","213","214","221","222","223","224","225","226","227","228","229","262","263","264","265","266","291","292","293","295","296","297","299","310","320","341","342","343","344","345","346","347","349"],
        "Фонд": ["'00","'83","'EВ","'81","'Ю6","'90"],
        "КЗ": [ "'з","'5","'е","'р","'0"],
        "Подраздел": ["'0701","'0702","'0703","'0705","'0707","'0709","'1003","'1101","'1102","'1103","'1004"],
        "Целевая статья": ["'021EВ51790","'021Ю650500","'021Ю651790","'021Ю653030","'0250020010","'0250020030","'0250020060","'0250020070","'0250020080","'0250020090","'0250020300","'0250020370","'0250020400","'0250020430","'0250020440","'0250020620","'0250020680","'0250020940","'0250021100","'0250021450","'02500R0500","'02500R3030","'0350040240","'0350040650","'03500R3040","'0450045010","'0450045020","'0450045060","'0450045090","'0450045390","'0850070890","'0950083620","'1750078010","'1750078560","'1750078650","'1750079450","'9950003450"],
        "Тип обеспечения": ["СГЗ", "ОСГЗ", "СИЦ", "ОСИЦ", "ПД", "ОПД"],
        "Счет расходов": ["550360","543148","552578","552555","552556","552524","543131","543134","543132","543133","543135","543138","543136","543137","543139","543142","543140","543141","543145","543146","543147","546940","553286","543149","543150","543151","553062","553086","546939","543152","543153","543154","543156","543157","543158","547692","547700","552980","548038","546941","552843","552846","552399","550695","546942","543143","543144","550357","546943","543159","543160","543161","549010","546944","546947","546945","546946","547671","547004","547005","547006","552904","546766","546767","546768","547942","543155","546769","551361","551362"]
    }

    # Словарь правил для автоматического заполнения "Счет расходов"
    expense_account_rules = {
        ("'021EВ51790", "'0702", "'з", "'83"): "550360",
        ("'021EВ51790", "'0702", "'з", "'EВ"): "543148",
        ("'021Ю650500", "'0709", "'5", "'81"): "552578",
        ("'021Ю651790", "'0702", "'з", "'83"): "552555",
        ("'021Ю651790", "'0702", "'з", "'Ю6"): "552556",
        ("'021Ю653030", "'0702", "'5", "'81"): "552524",
        ("'0250020010", "'0701", "'5", "'00"): "543131",
        ("'0250020010", "'0701", "'е", "'00"): "543134",
        ("'0250020010", "'0701", "'з", "'00"): "543132",
        ("'0250020010", "'0701", "'р", "'00"): "543133",
        ("'0250020030", "'0702", "'5", "'00"): "543135",
        ("'0250020030", "'0702", "'е", "'00"): "543138",
        ("'0250020030", "'0702", "'з", "'00"): "543136",
        ("'0250020030", "'0702", "'р", "'00"): "543137",
        ("'0250020060", "'0702", "'5", "'00"): "543139",
        ("'0250020060", "'0702", "'е", "'00"): "543142",
        ("'0250020060", "'0702", "'з", "'00"): "543140",
        ("'0250020060", "'0702", "'р", "'00"): "543141",
        ("'0250020070", "'0702", "'5", "'00"): "543145",
        ("'0250020070", "'0702", "'з", "'00"): "543146",
        ("'0250020070", "'0702", "'р", "'00"): "543147",
        ("'0250020080", "'0702", "'5", "'00"): "546940",
        ("'0250020090", "'0702", "'5", "'00"): "553286",
        ("'0250020300", "'0703", "'5", "'00"): "543149",
        ("'0250020300", "'0703", "'з", "'00"): "543150",
        ("'0250020300", "'0703", "'р", "'00"): "543151",
        ("'0250020370", "'0709", "'5", "'00"): "553062",
        ("'0250020370", "'0709", "'5", "'00"): "553086",
        ("'0250020400", "'0701", "'5", "'00"): "546939",
        ("'0250020430", "'0709", "'5", "'00"): "543152",
        ("'0250020430", "'0709", "'з", "'00"): "543153",
        ("'0250020430", "'0709", "'р", "'00"): "543154",
        ("'0250020440", "'0709", "'5", "'00"): "543156",
        ("'0250020440", "'0709", "'з", "'00"): "543157",
        ("'0250020440", "'0709", "'р", "'00"): "543158",
        ("'0250020620", "'0705", "'0", "'00"): "547692",
        ("'0250020680", "'0701", "'5", "'00"): "547700",
        ("'0250020940", "'0702", "'5", "'00"): "552980",
        ("'0250020940", "'0709", "'5", "'00"): "548038",
        ("'0250021100", "'0703", "'5", "'00"): "546941",
        ("'0250021450", "'0701", "'5", "'00"): "552843",
        ("'0250021450", "'0702", "'5", "'00"): "552846",
        ("'02500R0500", "'0709", "'5", "'81"): "552399",
        ("'02500R3030", "'0702", "'5", "'81"): "550695",
        ("'0350040240", "'1003", "'0", "'00"): "546942",
        ("'0350040650", "'0702", "'е", "'00"): "543143",
        ("'03500R3040", "'0702", "'е", "'00"): "543144",
        ("'03500R3040", "'0702", "'е", "'83"): "550357",
        ("'0450045010", "'1102", "'0", "'00"): "546943",
        ("'0450045020", "'1101", "'0", "'00"): "543159",
        ("'0450045020", "'1101", "'з", "'00"): "543160",
        ("'0450045020", "'1101", "'р", "'00"): "543161",
        ("'0450045060", "'1101", "'0", "'00"): "549010",
        ("'0450045090", "'1103", "'5", "'00"): "546944",
        ("'0450045090", "'1103", "'е", "'00"): "546947",
        ("'0450045090", "'1103", "'з", "'00"): "546945",
        ("'0450045090", "'1103", "'р", "'00"): "546946",
        ("'0450045390", "'1103", "'0", "'00"): "547671",
        ("'0850070890", "'0703", "'5", "'00"): "547004",
        ("'0850070890", "'0703", "'з", "'00"): "547005",
        ("'0850070890", "'0703", "'р", "'00"): "547006",
        ("'0950083620", "'1004", "'5", "'00"): "552904",
        ("'1750078010", "'0707", "'5", "'00"): "546766",
        ("'1750078010", "'0707", "'з", "'00"): "546767",
        ("'1750078010", "'0707", "'р", "'00"): "546768",
        ("'1750078560", "'0707", "'5", "'00"): "547942",
        ("'1750078650", "'0709", "'5", "'00"): "543155",
        ("'1750079450", "'0707", "'0", "'00"): "546769",
        ("'9950003450", "'0709", "'0", "'00"): "551361",
        ("'9950003450", "'0709", "'0", "'00"): "551362",
    }
    # Глобальные переменные для нужных Combobox
    global fund_combobox, kz_combobox, subdivision_combobox, target_article_combobox, expense_account_combobox

    # Список для хранения всех полей
    entries = []
    comboboxes = {}  # Для хранения Combobox

    for i, label in enumerate(fields):
        tk.Label(scrollable_frame, text=label + ":", font=("Arial", 14)).grid(row=i+4, column=0, padx=10, pady=5, sticky="w")

        if label in combobox_options:
            # Создаем Combobox
            combobox = Combobox(scrollable_frame, font=("Arial", 14), width=25)
            combobox["values"] = combobox_options[label]
            combobox.grid(row=i+4, column=1, padx=10, pady=5, sticky="w")

            comboboxes[label] = combobox  # Сохраняем Combobox
            entries.append(combobox)  # Добавляем в список entries
        else:
            # Создаем Entry
            entry = tk.Entry(scrollable_frame, font=("Arial", 14), width=25)
            entry.grid(row=i+4, column=1, padx=10, pady=5, sticky="w")
            entries.append(entry)

    # Назначаем переменные Combobox (после создания всех Combobox)
    fund_combobox = comboboxes.get("Фонд")
    kz_combobox = comboboxes.get("КЗ")
    subdivision_combobox = comboboxes.get("Подраздел")
    target_article_combobox = comboboxes.get("Целевая статья")
    expense_account_combobox = comboboxes.get("Счет расходов")

    # Функция для обновления "Счет расходов"
    def update_expense_account(event=None):
        if not (target_article_combobox and subdivision_combobox and kz_combobox and fund_combobox and expense_account_combobox):
            print("Ошибка: Один из Combobox не найден!")
            return

        # Получаем значения, удостоверившись, что они не None
        target_article_value = target_article_combobox.get().strip() if target_article_combobox.get() else ""
        subdivision_value = subdivision_combobox.get().strip() if subdivision_combobox.get() else ""
        kz_value = kz_combobox.get().strip() if kz_combobox.get() else ""
        fund_value = fund_combobox.get().strip() if fund_combobox.get() else ""

        key = (target_article_value, subdivision_value, kz_value, fund_value)

        print(f"Выбранные значения: {key}")  # Проверяем, что приходит в key

        if key in expense_account_rules:
            expense_account_combobox.set(expense_account_rules[key])  # Заполняем "Счет расходов"
            print(f"Найдено соответствие: {expense_account_rules[key]}")
        else:
            expense_account_combobox.set("")
            print("Совпадение не найдено.")

    # Привязываем функцию только к нужным Combobox
    if target_article_combobox and subdivision_combobox and kz_combobox and fund_combobox:
        target_article_combobox.bind("<<ComboboxSelected>>", update_expense_account)
        subdivision_combobox.bind("<<ComboboxSelected>>", update_expense_account)
        kz_combobox.bind("<<ComboboxSelected>>", update_expense_account)
        fund_combobox.bind("<<ComboboxSelected>>", update_expense_account)
    else:
        print("Предупреждение: Не все Combobox были созданы!")

    # Таблица для отображения данных
    table = Treeview(cert_window, columns=("Сотрудник", "Наименование учреждения", "Дата", "Наименование услуги", "КВР", "КОСГУ", "Фонд", "КЗ", "Подраздел", "Целевая статья", "Тип обеспечения", "Счет расходов", "Доступный объем субсидии", "Необходимый объем к перераспределению"), show="headings")
    for col in table["columns"]:
        table.heading(col, text=col)
        table.column(col, width=140)  # Немного уменьшена ширина столбцов таблицы

    # Сохраняем ссылку на таблицу в глобальной переменной
    global table_data
    table_data = table

    table.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

    # --- Фрейм для кнопок под таблицей ---
    button_frame = tk.Frame(cert_window)
    button_frame.grid(row=2, column=0, columnspan=2, pady=10)

    # Кнопки "Добавить в таблицу" и "Отправить"
    add_button = ttk.Button(button_frame, text="Добавить в таблицу", command=lambda: add_to_table(partial=True))
    add_button.pack(side=tk.LEFT, padx=10)

    send_print_button = ttk.Button(button_frame, text="Отправить и распечатать", command=lambda: send_and_print_partial())
    send_print_button.pack(side=tk.LEFT, padx=10)

    # --- Обновляем интерфейс при загрузке ---
    update_ui_scale_partial()
# --- Функция для печати последней записи (неполное финансирование) ---
def send_and_print_partial():
    global institution_combobox, date_entry, service_combobox, employee_entry, entries, table

    selected_item = table.selection()
    if not selected_item:
        messagebox.showwarning("Ошибка", "Выберите строку для отправки и печати.")
        return

    # Получаем значения из таблицы
    selected_record = table.item(selected_item[0])["values"]

    # Проверяем, что все значения в selected_record не None и не пустые строки
    if not all(selected_record):
        messagebox.showerror("Ошибка", "В выбранной строке таблицы есть пустые значения. Заполните все поля.")
        return

    # Распаковываем значения из selected_record
    employee_name, institution_name, date_info, service_name, kvp_value, kosgu_value, \
    fond_value, kz_value, podrazdel_value, target_article_value, type_obespecheniya_value, \
    schet_rashodov_value, neobhodimiy_obyom_value, dostupniy_obyom_value = selected_record

    try:
        printer_name = win32print.GetDefaultPrinter()
        hprinter = win32print.OpenPrinter(printer_name)
        win32print.ClosePrinter(hprinter)

        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        hdc.StartDoc("Справка финансирования (неполная)")
        hdc.StartPage()

        # Шрифт заголовка
        title_font = win32ui.CreateFont({"name": "Arial", "height": 120, "weight": 700})
        hdc.SelectObject(title_font)
        hdc.TextOut(100, 200, "Справка финансирования (неполная)")

        institution_y_start = 400
        hdc.TextOut(100, institution_y_start, "Наименование учреждения:")
        wrapped_lines = textwrap.wrap(str(institution_name), width=50) # Ensure string conversion

        for i, line in enumerate(wrapped_lines):
            hdc.TextOut(1600, institution_y_start + (i * 100), line)

        hdc.TextOut(100, institution_y_start + 100 * len(wrapped_lines) + 100, f"Дата: {str(date_info)}") # Ensure string conversion

        # Шрифт таблицы
        table_font = win32ui.CreateFont({"name": "Arial", "height": 80, "weight": 400})
        hdc.SelectObject(table_font)

        # Заголовки (обратите внимание на замену "Объем субсидии")
        headers = ["Название услуги", "КОСГУ", "КВР", "Фонд", "КЗ", "Подраздел",
                   "Целевая статья", "Тип обеспечения", "Счет расходов", "Доступный объем субсидии"]

        x_start = 100
        column_widths = [800, 300, 300, 300, 200, 450, 700, 500, 500, 700]
        row_height = 100
        header_y_start = institution_y_start + 100 * len(wrapped_lines) + 300

        for i, header in enumerate(headers):
            max_chars = get_max_chars_per_line(column_widths[i], hdc)
            wrapped_header_lines = textwrap.wrap(header, width=max_chars)[:2]
            for j in range(len(wrapped_header_lines)):
                hdc.TextOut(x_start + sum(column_widths[:i]), header_y_start + j * row_height, wrapped_header_lines[j])

        # Данные из полей слева (из selected_record)
        values = [str(service_name), str(kosgu_value), str(kvp_value), str(fond_value), str(kz_value), str(podrazdel_value),
                  str(target_article_value), str(type_obespecheniya_value), str(schet_rashodov_value), str(neobhodimiy_obyom_value)]

        value_start_y = header_y_start + row_height * 2
        for i, value in enumerate(values):
            max_chars = get_max_chars_per_line(column_widths[i], hdc)
            wrapped_value_lines = textwrap.wrap(value, width=max_chars)
            for j, line in enumerate(wrapped_value_lines):
                hdc.TextOut(x_start + sum(column_widths[:i]), value_start_y + (j * row_height), line)

        # Доступный объем субсидии
        hdc.TextOut(100, value_start_y + 900, f"Финансирование на сумму: {str(dostupniy_obyom_value)}") # Ensure string conversion

        # Сотрудник и подпись
        hdc.TextOut(100, value_start_y + 1700, f"Сотрудник: {employee_name}")
        hdc.TextOut(1200, value_start_y + 1700, "Подпись: __")

        hdc.EndPage()
        hdc.EndDoc()

        # Сохранение и удаление записи
        try:
            conn = psycopg2.connect(
                dbname="-----",
                user="-----",
                password="-----",
                host="------"
            )

            cursor = conn.cursor()
            create_table_without_money()  # Проверяем, существует ли таблица
            insert_query = '''
                INSERT INTO reference.save_without_money
                (employee, institution_name, date, service_name, kvp, kosgu, fund, kz, subdivision, target_article,
                security_type, expense_account, required_subsidy_volume, available_subsidy_volume)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            '''

            data_to_insert = (employee_name, institution_name, date_info, service_name, kvp_value, kosgu_value, fond_value, kz_value, podrazdel_value, target_article_value,
                            type_obespecheniya_value, schet_rashodov_value, neobhodimiy_obyom_value, dostupniy_obyom_value)

            cursor.execute(insert_query, data_to_insert)
            conn.commit()
            cursor.close()
            conn.close()

            # Удаление из таблицы
            table.delete(selected_item[0])

            messagebox.showinfo("Успех", "Справка с неполным финансированием отправлена и распечатана!")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить справку или удалить запись: {e}")

    except Exception as e:
        messagebox.showerror("Ошибка печати", f"Не удалось отправить документ на печать: {e}")

def get_max_chars_per_line(width, hdc):
    text_extent = hdc.GetTextExtent("A")  # Получаем размер текста 'A'
    return width // text_extent[0]  # Извлекаем ширину из кортежа

def create_table_without_money():
    try:
        conn = psycopg2.connect(
            dbname="id_table",
            user="postgres",
            password="root",
            host="192.168.1.186"
        )

        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reference.save_without_money (
                employee VARCHAR(255),
                institution_name VARCHAR(255),
                date DATE,
                service_name VARCHAR(255),
                kvp VARCHAR(255),
                kosgu VARCHAR(255),
                fund VARCHAR(255),
                kz VARCHAR(255),
                subdivision VARCHAR(255),
                target_article VARCHAR(255),
                security_type VARCHAR(255),
                expense_account VARCHAR(255),
                required_subsidy_volume VARCHAR(255),
                available_subsidy_volume VARCHAR(255)
            )
        ''')
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Не удалось создать таблицу: {e}")


# Функция для создания таблицы save_with_reference в PostgreSQL
def create_save_with_reference():
    try:
        conn = psycopg2.connect(
            dbname="id_table",
            user="postgres",
            password="root",
            host="192.168.1.186"
        )
        cursor = conn.cursor()

        create_table_query = '''
        CREATE TABLE IF NOT EXISTS reference.save_with_reference (
            employee VARCHAR(255),
            institution_name VARCHAR(255),
            date DATE,
            service_name VARCHAR(255),
            kvp VARCHAR(50),  -- Делаем текстовым, как в save
            kosgu VARCHAR(50),  -- Делаем текстовым, как в save
            fund VARCHAR(50),
            kz VARCHAR(50),
            subdivision VARCHAR(50),
            target_article VARCHAR(255),
            security_type VARCHAR(100),
            expense_account VARCHAR(50),  -- Делаем текстовым, как в save
            subsidy_volume VARCHAR(50),  -- Делаем текстовым, как в save
            rnk_value VARCHAR(255)
        )
        '''
        cursor.execute(create_table_query)
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Не удалось создать таблицу: {e}")

create_save_with_reference()

# Функция для фильтрации записей из таблицы reference.save
def filter_records(institution_name, target_article, kosgu, service_name):
    try:
        conn = psycopg2.connect(
            dbname="id_table",
            user="postgres",
            password="root",
            host="192.168.1.186"
        )
        cursor = conn.cursor()

        # Загружаем записи из save
        cursor.execute('''
        SELECT employee, institution_name, date, service_name, kvp, kosgu, fund, kz, subdivision, 
               target_article, security_type, expense_account, subsidy_volume
        FROM reference.save
        WHERE institution_name LIKE %s 
              AND target_article LIKE %s
              AND kosgu LIKE %s 
              AND service_name LIKE %s
        ''', ('%' + institution_name + '%', '%' + target_article + '%', 
              '%' + kosgu + '%', '%' + service_name + '%'))
        records = cursor.fetchall()

        # Загружаем записи из save_with_reference с приведением типов к TEXT
        cursor.execute('''
        SELECT employee, institution_name, date, service_name, CAST(kvp AS TEXT), CAST(kosgu AS TEXT), fund, kz, subdivision, 
               target_article, security_type, CAST(expense_account AS TEXT), CAST(subsidy_volume AS TEXT)
        FROM reference.save_with_reference
        ''')
        existing_records = set(cursor.fetchall())

        conn.close()

        # Исключаем строки, которые уже есть в save_with_reference
        filtered_records = [rec for rec in records if rec not in existing_records]

        return filtered_records
    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Не удалось получить данные: {e}")
        return []

# Функция для получения значений для фильтров из таблицы reference.save
def get_filter_values():
    try:
        conn = psycopg2.connect(
            dbname="-----",
            user="-----",
            password="-----",
            host="------"
        )
        cursor = conn.cursor()

        cursor.execute('SELECT DISTINCT institution_name FROM reference.save')
        institutions = [row[0] for row in cursor.fetchall()]

        cursor.execute('SELECT DISTINCT target_article FROM reference.save')
        target_articles = [row[0] for row in cursor.fetchall()]

        cursor.execute('SELECT DISTINCT kosgu FROM reference.save')
        kosgu_values = [row[0] for row in cursor.fetchall()]

        cursor.execute('SELECT DISTINCT service_name FROM reference.save')
        service_names = [row[0] for row in cursor.fetchall()]

        conn.close()
        return institutions, target_articles, kosgu_values, service_names
    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Не удалось загрузить фильтры: {e}")
        return [], [], [], []
    
# Функция для добавления новой записи РНК в таблицу save_with_reference
def add_rnk_record(selected_record):
    try:
        conn = psycopg2.connect(
            dbname="id_table",
            user="postgres",
            password="root",
            host="192.168.1.186"
        )
        cursor = conn.cursor()

        insert_query = '''
        INSERT INTO reference.save_with_reference 
        (employee, institution_name, date, service_name, kvp, kosgu, fund, kz, subdivision, target_article, 
         security_type, expense_account, subsidy_volume, rnk_value)
        VALUES (%s, %s, %s, %s, CAST(%s AS TEXT), CAST(%s AS TEXT), %s, %s, %s, %s, %s, CAST(%s AS TEXT), CAST(%s AS TEXT), %s)
        '''

        cursor.execute(insert_query, tuple(selected_record))
        conn.commit()
        cursor.close()
        conn.close()

    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Не удалось добавить запись РНК: {e}")

        cursor.execute(insert_query, tuple(selected_record))
        conn.commit()
        cursor.close()
        conn.close()

    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Не удалось добавить запись РНК: {e}")

# Окно добавления записи РНК
def add_rnk_window():
    global on_filter, zoom_level_rnk
    zoom_level_rnk = 1.0

    window = tk.Toplevel()
    window.title("Внести РНК")
    
    # Создаем Canvas для прокрутки
    canvas = tk.Canvas(window)
    scrollbar_y = tk.Scrollbar(window, orient="vertical", command=canvas.yview)
    scrollbar_x = tk.Scrollbar(window, orient="horizontal", command=canvas.xview)
    
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar_y.grid(row=0, column=2, sticky="ns")
    scrollbar_x.grid(row=23, column=0, sticky="ew")

    window.grid_rowconfigure(0, weight=1)
    window.grid_columnconfigure(0, weight=1)

    def adjust_zoom_rnk(event):
        global zoom_level_rnk
        if event.state & 0x4:  # Проверяем, зажат ли Ctrl
            if event.delta > 0:
                zoom_level_rnk *= 1.1
            else:
                zoom_level_rnk /= 1.1
            update_ui_scale_rnk()

    window.bind("<MouseWheel>", adjust_zoom_rnk)

    def update_ui_scale_rnk():
        scale = max(0.1, min(zoom_level_rnk, 3.0))
        new_font_size = int(14 * scale)
        new_entry_width = int(30 * scale)

        for widget in scrollable_frame.winfo_children():
            if isinstance(widget, tk.Label):
                widget.config(font=("Arial", new_font_size))
            elif isinstance(widget, tk.Entry) or isinstance(widget, ttk.Combobox):
                widget.config(font=("Arial", new_font_size), width=new_entry_width)
            elif isinstance(widget, ttk.Button):
                widget.config(style="Zoom.TButton")
        
        # Обновляем таблицу
        style.configure("Treeview", font=("Arial", new_font_size), rowheight=int(25 * scale))
        style.configure("Treeview.Heading", font=("Arial", new_font_size, "bold"))

        # Обновляем scrollregion
        canvas.configure(scrollregion=canvas.bbox("all"))

    # --- Стили для кнопок ---
    style = ttk.Style()
    style.configure("Zoom.TButton", font=("Arial", 14), padding=10)



    def on_filter():
        institution_name = institution_combobox.get()
        target_article = target_article_combobox.get()
        kosgu = kosgu_combobox.get()
        service_name = service_name_combobox.get()

        # Получаем отфильтрованные записи
        records = filter_records(institution_name, target_article, kosgu, service_name)

        # Обновляем таблицу для отображения отфильтрованных данных
        for row in tree.get_children():
            tree.delete(row)
        for record in records:
            tree.insert("", "end", values=record)  # Добавляем полную запись

    def on_add_rnk():
        selected_item = tree.selection()
        if selected_item:
            selected_record = tree.item(selected_item[0])["values"]
            rnk_value = entry_rnk.get()

            if rnk_value:
                try:
                    conn = psycopg2.connect(
                        dbname="-----",
                        user="------",
                        password="----",
                        host="-------"
                    )
                    cursor = conn.cursor()

                    # Приводим числовые значения к строковому типу
                    selected_record = list(selected_record)
                    selected_record[4] = str(selected_record[4])  # kvp
                    selected_record[5] = str(selected_record[5])  # kosgu
                    selected_record[11] = str(selected_record[11])  # expense_account
                    selected_record[12] = str(selected_record[12])  # subsidy_volume

                    # Добавляем РНК к строке
                    selected_record.append(rnk_value)

                    # Копируем строку в save_with_reference
                    insert_query = '''
                    INSERT INTO reference.save_with_reference 
                    (employee, institution_name, date, service_name, kvp, kosgu, fund, kz, subdivision, target_article, 
                    security_type, expense_account, subsidy_volume, rnk_value)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    '''
                    cursor.execute(insert_query, tuple(selected_record))

                    # Удаляем строку из save
                    delete_query = '''
                    DELETE FROM reference.save
                    WHERE employee = %s AND institution_name = %s AND date = %s
                        AND service_name = %s 
                        AND CAST(kvp AS TEXT) = CAST(%s AS TEXT)
                        AND CAST(kosgu AS TEXT) = CAST(%s AS TEXT) 
                        AND fund = %s AND kz = %s
                        AND subdivision = %s AND target_article = %s 
                        AND security_type = %s 
                        AND CAST(expense_account AS TEXT) = CAST(%s AS TEXT) 
                        AND CAST(subsidy_volume AS TEXT) = CAST(%s AS TEXT)
                    '''
                    cursor.execute(delete_query, tuple(selected_record[:-1]))  # Без rnk_value

                    conn.commit()
                    cursor.close()
                    conn.close()

                    messagebox.showinfo("Успех", "РНК добавлен! Запись перенесена в save_with_reference.")
                    on_filter()  # Обновляем список

                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось добавить РНК: {e}")
            else:
                messagebox.showwarning("Ошибка", "Введите значение РНК.")
        else:
            messagebox.showwarning("Ошибка", "Выберите запись из таблицы.")

    global tree


    institutions, target_articles, kosgu_values, service_names = get_filter_values()

    filter_frame = tk.Frame(window)
    filter_frame.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    tk.Label(scrollable_frame, text="Наименование учреждения").grid(row=0, column=0, sticky="w")
    institution_combobox = ttk.Combobox(scrollable_frame, width=100, values=institutions)
    institution_combobox.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(scrollable_frame, text="Целевая статья").grid(row=1, column=0, sticky="w")
    target_article_combobox = ttk.Combobox(scrollable_frame, width=100, values=target_articles)
    target_article_combobox.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(scrollable_frame, text="Косгу").grid(row=2, column=0, sticky="w")
    kosgu_combobox = ttk.Combobox(scrollable_frame, width=100, values=kosgu_values)
    kosgu_combobox.grid(row=2, column=1, padx=5, pady=5)

    tk.Label(scrollable_frame, text="Наименование услуги").grid(row=3, column=0, sticky="w")
    service_name_combobox = ttk.Combobox(scrollable_frame, width=100, values=service_names)
    service_name_combobox.grid(row=3, column=1, padx=5, pady=5)

    institution_combobox.bind("<KeyRelease>", lambda event: filter_combobox(event, institution_combobox, institutions))
    service_name_combobox.bind("<KeyRelease>", lambda event: filter_combobox(event, service_name_combobox, service_names))

    global entry_rnk
    # Поле для ввода РНК (увеличено в 2 раза)
    tk.Label(scrollable_frame, text="Введите РНК:").grid(row=5, column=0,padx=10, pady=10, sticky="w")
    entry_rnk = tk.Entry(scrollable_frame, width=60, font=("Arial", 12))  # Увеличено в 2 раза
    entry_rnk.grid(row=5, column=1, padx=10, pady=10)

    # Кнопка для добавления РНК (увеличена в 2 раза)
    add_rnk_button = ttk.Button(window, text="Добавить РНК", command=on_add_rnk)
    add_rnk_button.grid(row=3, column=0, columnspan=2, pady=10)


    filter_button = ttk.Button(scrollable_frame, text="Применить фильтры", command=on_filter)
    filter_button.grid(row=4, column=0, columnspan=2, pady=10)


    tree_frame = tk.Frame(window)
    tree_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

    tree_scroll_x = tk.Scrollbar(tree_frame, orient="horizontal")
    tree_scroll_x.pack(side="bottom", fill="x")

    columns = ("Сотрудник", "Учреждение", "Дата", "Услуга", "КВР", "КОСГУ", "Фонд", "КЗ", "Подразделение",
               "Целевая статья", "Тип обеспечения", "Счет расходов", "Сумма субсидии", "РНК")

    tree = ttk.Treeview(tree_frame, columns=columns, show="headings", xscrollcommand=tree_scroll_x.set)
    tree_scroll_x.config(command=tree.xview)

    column_widths = [80, 130, 80, 130, 50, 50, 50, 50, 80, 100, 100, 80, 80, 80]

    for col, width in zip(columns, column_widths):
        tree.heading(col, text=col)
        tree.column(col, width=width, anchor="center")

    tree.pack(fill=tk.BOTH, expand=True)

    window.mainloop()



# Запуск создания таблицы save_with_reference
    create_save_with_reference()

# Пример вызова окна
# add_rnk_window()

def delete_certificate():
    def check_delete_password():
        if password_entry.get() == "123123":
            auth_window.destroy()
            show_delete_table()  # После успешного ввода пароля показываем таблицу
        else:
            messagebox.showerror("Ошибка", "Неверный пароль!")

    auth_window = tk.Toplevel()
    auth_window.title("Авторизация")
    auth_window.geometry("300x150")

    tk.Label(auth_window, text="Введите пароль для удаления:", font=("Arial", 12)).pack(pady=5)
    password_entry = tk.Entry(auth_window, font=("Arial", 12), width=20, show="*")
    password_entry.pack(pady=5)

    tk.Button(auth_window, text="Подтвердить", font=("Arial", 12), command=check_delete_password).pack(pady=10)

def show_delete_table():
    def on_filter():
        institution_name = institution_combobox.get()
        target_article = target_article_combobox.get()
        kosgu = kosgu_combobox.get()
        service_name = service_name_combobox.get()

        try:
            conn = psycopg2.connect(
                dbname="------",
                user="-----",
                password="----",
                host="-------"
            )
            cursor = conn.cursor()

            # Загружаем данные из save и save_with_reference с указанием источника
            query = '''
            SELECT employee, institution_name, date, service_name, 
                CAST(kvp AS TEXT), CAST(kosgu AS TEXT), fund, kz, subdivision, 
                target_article, security_type, CAST(expense_account AS TEXT), 
                CAST(subsidy_volume AS TEXT), NULL AS rnk_value, 'save' AS source_table
            FROM reference.save
            WHERE institution_name LIKE %s AND target_article LIKE %s
                AND kosgu LIKE %s AND service_name LIKE %s

            UNION ALL

            SELECT employee, institution_name, date, service_name, 
                CAST(kvp AS TEXT), CAST(kosgu AS TEXT), fund, kz, subdivision, 
                target_article, security_type, CAST(expense_account AS TEXT), 
                CAST(subsidy_volume AS TEXT), rnk_value, 'save_with_reference' AS source_table
            FROM reference.save_with_reference
            WHERE institution_name LIKE %s AND target_article LIKE %s
                AND kosgu LIKE %s AND service_name LIKE %s
            '''
            
            cursor.execute(query, ('%' + institution_name + '%', '%' + target_article + '%', 
                                '%' + kosgu + '%', '%' + service_name + '%',
                                '%' + institution_name + '%', '%' + target_article + '%', 
                                '%' + kosgu + '%', '%' + service_name + '%'))
            
            records = cursor.fetchall()
            conn.close()

            # Очищаем таблицу перед обновлением
            for row in tree.get_children():
                tree.delete(row)

            # Добавляем строки с источником
            for record in records:
                tree.insert("", "end", values=record)

        except Exception as e:
            messagebox.showerror("Ошибка базы данных", f"Не удалось получить данные: {e}")

    def on_delete():
        selected_item = tree.selection()
        if selected_item:
            selected_record = tree.item(selected_item[0])["values"]
            source_table = selected_record[-1]  # Последнее поле - источник таблицы

            move_to_deleted(selected_record[:-1], source_table)  # Убираем source_table из передаваемых данных
            tree.delete(selected_item)


    window = tk.Toplevel()
    window.title("Удаление справки")

    filter_frame = tk.Frame(window)
    filter_frame.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    # Загружаем данные для фильтров
    conn = psycopg2.connect(
        dbname="------",
        user="------",
        password="------",
        host="--------"
    )
    cursor = conn.cursor()

    cursor.execute("SELECT DISTINCT institution_name FROM reference.save UNION SELECT DISTINCT institution_name FROM reference.save_with_reference")
    institutions = [row[0] for row in cursor.fetchall()]

    cursor.execute("SELECT DISTINCT target_article FROM reference.save UNION SELECT DISTINCT target_article FROM reference.save_with_reference")
    target_articles = [row[0] for row in cursor.fetchall()]

    cursor.execute("SELECT DISTINCT kosgu FROM reference.save UNION SELECT DISTINCT kosgu FROM reference.save_with_reference")
    kosgu_values = [row[0] for row in cursor.fetchall()]

    cursor.execute("SELECT DISTINCT service_name FROM reference.save UNION SELECT DISTINCT service_name FROM reference.save_with_reference")
    service_names = [row[0] for row in cursor.fetchall()]

    conn.close()

    tk.Label(filter_frame, text="Наименование учреждения").grid(row=0, column=0, sticky="w")
    institution_combobox = ttk.Combobox(filter_frame, width=35, values=institutions)
    institution_combobox.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(filter_frame, text="Целевая статья").grid(row=1, column=0, sticky="w")
    target_article_combobox = ttk.Combobox(filter_frame, width=35, values=target_articles)
    target_article_combobox.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(filter_frame, text="Косгу").grid(row=2, column=0, sticky="w")
    kosgu_combobox = ttk.Combobox(filter_frame, width=35, values=kosgu_values)
    kosgu_combobox.grid(row=2, column=1, padx=5, pady=5)

    tk.Label(filter_frame, text="Наименование услуги").grid(row=3, column=0, sticky="w")
    service_name_combobox = ttk.Combobox(filter_frame, width=35, values=service_names)
    service_name_combobox.grid(row=3, column=1, padx=5, pady=5)

    tk.Button(filter_frame, text="Применить фильтр", command=on_filter).grid(row=4, column=0, columnspan=2, pady=10)

    tree_frame = tk.Frame(window)
    tree_frame.grid(row=1, column=0, padx=10, pady=10)

    tree = ttk.Treeview(tree_frame, columns=("Сотрудник", "Учреждение", "Дата", "Услуга",
                                             "КВР", "КОСГУ", "Фонд", "КЗ", "Подразделение",
                                             "Целевая статья", "Тип обеспечения", "Счет расходов",
                                             "Сумма субсидии", "РНК", "Источник"), show="headings")

    for col in tree["columns"]:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor="center")

    tree.grid(row=0, column=0, sticky="nsew")

    # Принудительное обновление окна перед добавлением кнопки
    window.update_idletasks()

    # Кнопка "Удалить"
    delete_button = tk.Button(window, text="Удалить", command=on_delete, font=("Arial", 14))
    delete_button.grid(row=2, column=0, pady=10)

def move_to_deleted(selected_record, source_table):
    try:
        conn = psycopg2.connect(
            dbname="-------",
            user="------",
            password="------",
            host="-------"
        )
        cursor = conn.cursor()

        # Приводим числовые значения к строковому типу
        selected_record = list(selected_record)
        selected_record[4] = str(selected_record[4])  # kvp
        selected_record[5] = str(selected_record[5])  # kosgu
        selected_record[11] = str(selected_record[11])  # expense_account
        selected_record[12] = str(selected_record[12])  # subsidy_volume

        # Создаём таблицу reference_deleted, если её нет
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS reference.reference_deleted (
            employee VARCHAR(255),
            institution_name VARCHAR(255),
            date DATE,
            service_name VARCHAR(255),
            kvp VARCHAR(50),
            kosgu VARCHAR(50),
            fund VARCHAR(50),
            kz VARCHAR(50),
            subdivision VARCHAR(50),
            target_article VARCHAR(255),
            security_type VARCHAR(100),
            expense_account VARCHAR(50),
            subsidy_volume VARCHAR(50),
            rnk_value VARCHAR(255),
            source_table VARCHAR(50)
        )
        ''')

        # Добавляем запись в reference_deleted
        insert_query = '''
        INSERT INTO reference.reference_deleted 
        (employee, institution_name, date, service_name, kvp, kosgu, fund, kz, subdivision, target_article, 
         security_type, expense_account, subsidy_volume, rnk_value, source_table)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        '''
        cursor.execute(insert_query, tuple(selected_record) + (source_table,))

        # Формируем SQL-запрос в зависимости от наличия rnk_value
        if source_table == "save_with_reference":
            delete_query = f'''
            DELETE FROM reference.{source_table}
            WHERE employee = %s AND institution_name = %s AND date = %s
                  AND service_name = %s 
                  AND CAST(kvp AS TEXT) = CAST(%s AS TEXT)
                  AND CAST(kosgu AS TEXT) = CAST(%s AS TEXT) 
                  AND fund = %s AND kz = %s
                  AND subdivision = %s AND target_article = %s 
                  AND security_type = %s 
                  AND CAST(expense_account AS TEXT) = CAST(%s AS TEXT) 
                  AND CAST(subsidy_volume AS TEXT) = CAST(%s AS TEXT) 
                  AND (rnk_value IS NOT DISTINCT FROM %s OR (rnk_value IS NULL AND %s IS NULL))
            '''
            cursor.execute(delete_query, tuple(selected_record))
        else:  # Удаляем из save (без rnk_value)
            delete_query = f'''
            DELETE FROM reference.{source_table}
            WHERE employee = %s AND institution_name = %s AND date = %s
                  AND service_name = %s 
                  AND CAST(kvp AS TEXT) = CAST(%s AS TEXT)
                  AND CAST(kosgu AS TEXT) = CAST(%s AS TEXT) 
                  AND fund = %s AND kz = %s
                  AND subdivision = %s AND target_article = %s 
                  AND security_type = %s 
                  AND CAST(expense_account AS TEXT) = CAST(%s AS TEXT) 
                  AND CAST(subsidy_volume AS TEXT) = CAST(%s AS TEXT)
            '''
            cursor.execute(delete_query, tuple(selected_record[:-1]))  # rnk_value не передаём

        conn.commit()
        cursor.close()
        conn.close()
        messagebox.showinfo("Успех", "Справка успешно удалена!")

    except Exception as e:
        messagebox.showerror("Ошибка базы данных", f"Не удалось удалить справку: {e}")

# Запуск
main_screen()

