import tkinter as tk
from tkinter import filedialog

# Создание основного окна [1](https://www.iditect.com/faq/python/choosing-a-file-in-python-with-simple-dialog.html)
root = tk.Tk()
root.withdraw()

# Создание окна выбора файла [1](https://www.iditect.com/faq/python/choosing-a-file-in-python-with-simple-dialog.html)
file_path = filedialog.askopenfilename(title="Выберите файл")
root.destroy()
# Проверка, выбран ли файл [1](https://www.iditect.com/faq/python/choosing-a-file-in-python-with-simple-dialog.html)
if file_path:
    print("Выбранный файл:", file_path)
else:
    print("Файл не выбран")
# Закрытие основного окна (необязательно) [1](https://www.iditect.com/faq/python/choosing-a-file-in-python-with-simple-dialog.html)
