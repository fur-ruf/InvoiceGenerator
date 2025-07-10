from docxtpl import DocxTemplate, RichText
from tkinter import Tk, Label, Button, filedialog, messagebox, Text, END, StringVar, Frame, Canvas, Scrollbar, VERTICAL, \
    RIGHT, Y, BOTH, NW
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
        # Функция для сохранения переносов строк через RichText
        def format_text(text):
            rt = RichText()
            lines = text.split('\n')
            for i, line in enumerate(lines):
                if i > 0:  # Добавляем перенос для всех строк кроме первой
                    rt.add('\n')
                rt.add(line)
            return rt

        context = {
            'data': data_entry.get(),
            'number': number_entry.get(),
            'sender': sender_var.get(),
            'receiver': receiver_var.get(),
            'products': format_text(products_text.get("1.0", END).strip()),
            'counts': format_text(counts_text.get("1.0", END).strip()),
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
        doc.render(context)

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


# Основное окно с полноэкранной прокруткой
root = Tk()
root.title("Генератор транспортных накладных")
root.geometry("850x600")  # Стартовый размер окна
root.minsize(800, 500)  # Минимальный размер

# Создаем основной контейнер с прокруткой
container = Frame(root)
container.pack(fill=BOTH, expand=True)

# Создаем Canvas и Scrollbar
canvas = Canvas(container)
scrollbar = Scrollbar(container, orient=VERTICAL, command=canvas.yview)
scrollable_frame = Frame(canvas)

# Привязка прокрутки
scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)

canvas.create_window((0, 0), window=scrollable_frame, anchor=NW)
canvas.configure(yscrollcommand=scrollbar.set)

# Размещаем элементы
canvas.pack(side="left", fill=BOTH, expand=True)
scrollbar.pack(side="right", fill=Y)


# Настройка прокрутки колесиком мыши
def _on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


canvas.bind_all("<MouseWheel>", _on_mousewheel)

# Стилизация
style = ttk.Style()
style.configure('TEntry', font=MEDIUM_FONT, padding=5)
style.configure('TCombobox', font=MEDIUM_FONT, padding=5)
style.map('TCombobox',
          fieldbackground=[('readonly', 'white'), ('active', '#f0f0f0')],
          selectbackground=[('readonly', '#e6f2ff')])

# Поля ввода (обычные однострочные)
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
    Label(scrollable_frame, text=label_text, font=LARGE_FONT).grid(row=i, column=0, padx=10, pady=5, sticky="e")
    entry = ttk.Entry(scrollable_frame, width=35, font=MEDIUM_FONT)
    entry.grid(row=i, column=1, padx=10, pady=5, sticky="w")
    globals()[var_name] = entry

# Выпадающие списки для компаний
sender_var = StringVar()
receiver_var = StringVar()

row_start = len(fields)
Label(scrollable_frame, text="Грузоотправитель:", font=LARGE_FONT).grid(row=row_start, column=0, padx=10, pady=5,
                                                                        sticky="e")
sender_combobox = ttk.Combobox(scrollable_frame, textvariable=sender_var, width=32, font=MEDIUM_FONT)
sender_combobox.grid(row=row_start, column=1, padx=10, pady=5, sticky="w")

add_sender_btn = Button(scrollable_frame, text="+ Добавить",
                        command=lambda: create_company_popup("senders", update_comboboxes),
                        font=MEDIUM_FONT, bg='#2196F3', fg='black')
add_sender_btn.grid(row=row_start, column=2, padx=5, pady=5)

Label(scrollable_frame, text="Грузополучатель:", font=LARGE_FONT).grid(row=row_start + 1, column=0, padx=10, pady=5,
                                                                       sticky="e")
receiver_combobox = ttk.Combobox(scrollable_frame, textvariable=receiver_var, width=32, font=MEDIUM_FONT)
receiver_combobox.grid(row=row_start + 1, column=1, padx=10, pady=5, sticky="w")

add_receiver_btn = Button(scrollable_frame, text="+ Добавить",
                          command=lambda: create_company_popup("receivers", update_comboboxes),
                          font=MEDIUM_FONT, bg='#2196F3', fg='black')
add_receiver_btn.grid(row=row_start + 1, column=2, padx=5, pady=5)

# Многострочные поля
row_offset = row_start + 2
Label(scrollable_frame, text="Груз:", font=LARGE_FONT).grid(row=row_offset, column=0, padx=10, pady=5, sticky="ne")
products_text = Text(scrollable_frame, height=6, width=50, wrap="word", font=MEDIUM_FONT)
products_text.grid(row=row_offset, column=1, padx=10, pady=5, sticky="w", columnspan=2)

Label(scrollable_frame, text="Количество, маркировка:", font=LARGE_FONT).grid(row=row_offset + 1, column=0, padx=10,
                                                                              pady=5, sticky="ne")
counts_text = Text(scrollable_frame, height=6, width=50, wrap="word", font=MEDIUM_FONT)
counts_text.grid(row=row_offset + 1, column=1, padx=10, pady=5, sticky="w", columnspan=2)

# Кнопка создания
Button(scrollable_frame, text="Создать накладную", command=fill_template,
       font=BUTTON_FONT, bg='#4CAF50', fg='black', height=1, width=20) \
    .grid(row=row_offset + 2, column=0, columnspan=3, pady=15)

update_comboboxes()

# Добавление поддержки Ctrl+V
def setup_paste_support():
    # Функция для вставки в Combobox
    def paste_to_combobox(event):
        widget = event.widget
        if widget.selection_present():
            widget.delete(widget.selection_first(), widget.selection_last())
        widget.insert("insert", root.clipboard_get())

    # Обработчики для всех полей
    for entry in [data_entry, number_entry, transporter_entry, driver_entry,
                  car_entry, car_number_entry, from_entry, adress_from_entry,
                  signature_entry, adress_to_entry]:
        entry.bind("<Control-v>", lambda e: e.widget.event_generate("<<Paste>>"))

    for text in [products_text, counts_text]:
        text.bind("<Control-v>", lambda e: e.widget.event_generate("<<Paste>>"))

    for combobox in [sender_combobox, receiver_combobox]:
        combobox.bind("<Control-v>", paste_to_combobox)

setup_paste_support()

root.mainloop()
