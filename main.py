from docxtpl import DocxTemplate
from tkinter import Tk, Label, Button, filedialog, messagebox, Text, END, StringVar
import os
import platform
import subprocess
from tkinter import ttk
import json
from tkinter.font import Font

# Константы для файла хранения данных
DATA_FILE = "company_data.json"

# Настройки шрифтов
LARGE_FONT = ('Calibri', 14)
MEDIUM_FONT = ('Calibri', 12)
BUTTON_FONT = ('Calibri', 12, 'bold')


def load_companies():
    """Загружает список компаний из файла"""
    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {"senders": [], "receivers": []}


def save_companies(data):
    """Сохраняет список компаний в файл"""
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def add_company(company_type, name, callback):
    """Добавляет новую компанию и обновляет список"""
    data = load_companies()
    if name and name not in data[company_type]:
        data[company_type].append(name)
        save_companies(data)
        messagebox.showinfo("Успех", f"Компания '{name}' добавлена!")
        callback()
    else:
        messagebox.showwarning("Предупреждение", "Такая компания уже существует или название пустое")


def open_file(path):
    """Открывает файл в зависимости от ОС"""
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])
    except Exception as e:
        messagebox.showwarning("Предупреждение", f"Не удалось открыть файл: {e}")


def fill_template():
    try:
        def preserve_newlines(text):
            return text.replace('\n', '<w:br/>')

        context = {
            'data': data_entry.get(),
            'number': number_entry.get(),
            'sender': sender_var.get(),
            'receiver': receiver_var.get(),
            'products': preserve_newlines(products_text.get("1.0", END).strip()),
            'counts': preserve_newlines(counts_text.get("1.0", END).strip()),
            'transporter': transporter_entry.get(),
            'driver': driver_entry.get(),
            'car': car_entry.get(),
            'car_number': car_number_entry.get(),
            'from': from_entry.get(),
            'adress_from': adress_from_entry.get(),
            'signature': signature_entry.get(),
            'adress_to': adress_to_entry.get()
        }

        doc = DocxTemplate("template.docx")
        doc.render(context, autoescape=True)

        default_filename = f"Накладная_{context['number']}.docx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Files", "*.docx")],
            initialfile=default_filename
        )

        if save_path:
            doc.save(save_path)
            messagebox.showinfo("Готово!", f"Накладная сохранена:\n{save_path}")

            if messagebox.askyesno("Открыть", "Открыть созданный файл?"):
                open_file(save_path)

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")


def create_company_popup(company_type, callback):
    popup = Tk()
    popup.title(f"Добавить нового {company_type}")
    popup.geometry("400x200")

    Label(popup, text="Название компании:", font=LARGE_FONT).pack(pady=10)
    entry = ttk.Entry(popup, width=30, font=MEDIUM_FONT)
    entry.pack(pady=10)

    def save_and_close():
        add_company(company_type, entry.get(), callback)
        popup.destroy()

    Button(popup, text="Добавить", command=save_and_close,
           font=BUTTON_FONT, bg='#4CAF50', fg='black').pack(pady=10)


def update_comboboxes():
    data = load_companies()
    sender_combobox['values'] = data['senders']
    receiver_combobox['values'] = data['receivers']


# Основное окно
root = Tk()
root.title("Генератор транспортных накладных")
root.geometry("750x900")

# Стилизация
style = ttk.Style()
style.configure('TEntry', font=MEDIUM_FONT, padding=8)
style.configure('TCombobox', font=MEDIUM_FONT, padding=8)
style.map('TCombobox',
          fieldbackground=[('readonly', 'white'), ('active', '#f0f0f0')],
          selectbackground=[('readonly', '#e6f2ff')])

# Поля ввода
fields = [
    ("Дата (ДД.ММ.ГГГГ):", "data_entry"),
    ("Номер накладной:", "number_entry"),
    ("Перевозчик:", "transporter_entry"),
    ("Водитель:", "driver_entry"),
    ("Машина:", "car_entry"),
    ("Номер машины:", "car_number_entry"),
    ("8. Прием груза:", "from_entry"),
    ("Место погрузки:", "adress_from_entry"),
    ("Подпись:", "signature_entry"),
    ("Место доставки:", "adress_to_entry")
]

for i, (label_text, var_name) in enumerate(fields):
    Label(root, text=label_text, font=LARGE_FONT).grid(row=i, column=0, padx=10, pady=1, sticky="e")
    entry = ttk.Entry(root, width=35, font=MEDIUM_FONT)
    entry.grid(row=i, column=1, padx=10, pady=1, sticky="w")
    globals()[var_name] = entry

# Выпадающие списки
sender_var = StringVar()
receiver_var = StringVar()

Label(root, text="Грузоотправитель:", font=LARGE_FONT).grid(row=len(fields), column=0, padx=10, pady=1, sticky="e")
sender_combobox = ttk.Combobox(root, textvariable=sender_var, width=35, font=MEDIUM_FONT)
sender_combobox.grid(row=len(fields), column=1, padx=10, pady=1, sticky="w")

add_sender_btn = Button(root, text="+ Добавить", command=lambda: create_company_popup("senders", update_comboboxes),
                        font=MEDIUM_FONT, bg='#2196F3', fg='black')
add_sender_btn.grid(row=len(fields), column=2, padx=10, pady=1)

Label(root, text="Грузополучатель:", font=LARGE_FONT).grid(row=len(fields) + 1, column=0, padx=10, pady=1, sticky="e")
receiver_combobox = ttk.Combobox(root, textvariable=receiver_var, width=35, font=MEDIUM_FONT)
receiver_combobox.grid(row=len(fields) + 1, column=1, padx=10, pady=1, sticky="w")

add_receiver_btn = Button(root, text="+ Добавить", command=lambda: create_company_popup("receivers", update_comboboxes),
                          font=MEDIUM_FONT, bg='#2196F3', fg='black')
add_receiver_btn.grid(row=len(fields) + 1, column=2, padx=10, pady=1)

# Многострочные поля
row_offset = len(fields) + 2
Label(root, text="Груз:", font=LARGE_FONT).grid(row=row_offset, column=0, padx=10, pady=1, sticky="ne")
products_text = Text(root, height=8, width=50, wrap="word", font=MEDIUM_FONT)
products_text.grid(row=row_offset, column=1, padx=10, pady=1, sticky="w", columnspan=2)

Label(root, text="Количество, маркировка:", font=LARGE_FONT).grid(row=row_offset + 1, column=0, padx=10, pady=8,
                                                                  sticky="ne")
counts_text = Text(root, height=8, width=50, wrap="word", font=MEDIUM_FONT)
counts_text.grid(row=row_offset + 1, column=1, padx=10, pady=1, sticky="w", columnspan=2)

# Кнопка создания
Button(root, text="Создать накладную", command=fill_template,
       font=BUTTON_FONT, bg='#4CAF50', fg='black', height=2, width=20) \
    .grid(row=row_offset + 2, column=0, columnspan=3, pady=5)

update_comboboxes()
root.mainloop()
