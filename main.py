from docxtpl import DocxTemplate
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox, Text, END
import os
import platform
import subprocess
from tkinter import ttk


def open_file(path):
    """Открывает файл в зависимости от ОС"""
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", path])
        else:  # Linux и другие
            subprocess.run(["xdg-open", path])
    except Exception as e:
        messagebox.showwarning("Предупреждение", f"Не удалось открыть файл: {e}")


def fill_template():
    try:
        # Получаем данные из полей ввода
        context = {
            'data': data_entry.get(),
            'number': number_entry.get(),
            'sender': sender_entry.get(),
            'receiver': receiver_entry.get(),
            'products': products_text.get("1.0", END).strip(),
            'counts': counts_text.get("1.0", END).strip(),
            'transporter': transporter_entry.get(),
            'driver': driver_entry.get(),
            'car': car_entry.get(),
            'car_number': car_number_entry.get(),
            'from': from_entry.get(),
            'adress_from': adress_from_entry.get(),
            'signature': signature_entry.get(),
            'adress_to': adress_to_entry.get()
        }

        # Загружаем шаблон
        doc = DocxTemplate("template.docx")

        # Заполняем шаблон
        doc.render(context)

        # Сохраняем новый документ
        default_filename = f"Накладная_{context['number']}.docx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Files", "*.docx")],
            initialfile=default_filename
        )

        if save_path:
            doc.save(save_path)
            messagebox.showinfo("Готово!", f"Накладная сохранена:\n{save_path}")

            # Открываем файл (опционально)
            if messagebox.askyesno("Открыть", "Открыть созданный файл?"):
                open_file(save_path)

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")


# Создаем графический интерфейс
root = Tk()
root.title("Генератор транспортных накладных")
root.geometry("800x800")

# Стиль для полей ввода
style = ttk.Style()
style.configure('TEntry', padding='5 1 5 1')

# Поля ввода (обычные однострочные)
fields = [
    ("Дата (ДД.ММ.ГГГГ):", "data_entry"),
    ("Номер накладной:", "number_entry"),
    ("Грузоотправитель:", "sender_entry"),
    ("Грузополучатель:", "receiver_entry"),
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
    Label(root, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky="e")
    entry = ttk.Entry(root, width=40)
    entry.grid(row=i, column=1, padx=5, pady=5, sticky="w")
    globals()[var_name] = entry

# Многострочные поля

Label(root, text="Груз:").grid(row=len(fields) + 1, column=0, padx=5, pady=5, sticky="ne")
products_text = Text(root, height=6, width=40, wrap="word")
products_text.grid(row=len(fields) + 1, column=1, padx=5, pady=5, sticky="w")

Label(root, text="Количество, маркировка:").grid(row=len(fields) + 2, column=0, padx=5, pady=5, sticky="ne")
counts_text = Text(root, height=6, width=40, wrap="word")
counts_text.grid(row=len(fields) + 2, column=1, padx=5, pady=5, sticky="w")

# Кнопка "Создать"
Button(root, text="Создать накладную", command=fill_template).grid(
    row=len(fields) + 3, column=0, columnspan=2, pady=20)

root.mainloop()