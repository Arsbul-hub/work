import os
import random
from tkinter import filedialog, Tk

# import pandas as pd
import pyexcel as pe

"""
pyinstaller --noconfirm --onefile --console --hidden-import "pyexcel-xlsx" --hidden-import "pyexcel-xls" --hidden-import "openpyxl" --hidden-import "et-xmlfile" --hidden-import "xlwt" --hidden-import "xlrd" --hidden-import "lml" --hidden-import "pyexcel-io" --hidden-import "texttable" --add-data "C:/Users/arsbu/Desktop/work/.venv/Lib/site-packages/lml;lml/" --add-data "C:/Users/arsbu/Desktop/work/.venv/Lib/site-packages/et_xmlfile;et_xmlfile/" --add-data "C:/Users/arsbu/Desktop/work/.venv/Lib/site-packages/pyexcel;pyexcel/" --add-data "C:/Users/arsbu/Desktop/work/.venv/Lib/site-packages/pyexcel_io;pyexcel_io/" --add-data "C:/Users/arsbu/Desktop/work/.venv/Lib/site-packages/pyexcel_xls;pyexcel_xls/" --add-data "C:/Users/arsbu/Desktop/work/.venv/Lib/site-packages/pyexcel_xlsx;pyexcel_xlsx/" --add-data "C:/Users/arsbu/Desktop/work/.venv/Lib/site-packages/xlrd;xlrd/" --add-data "C:/Users/arsbu/Desktop/work/.venv/Lib/site-packages/xlwt;xlwt/" --hidden-import "urllib" --collect-all "pyexcel" --hidden-import "urllib.request" --hidden-import "pyexcel" --hidden-import "pkgutil" --hidden-import "json" --hidden-import "__future__" --hidden-import "openpyxl" --collect-all "openpyxl" --add-data "C:/Users/arsbu/Desktop/work/.venv/Lib/site-packages/openpyxl;openpyxl/" --hidden-import "xml" --hidden-import "xml.etree" --hidden-import "xml.etree.ElementTree" --collect-all "xml"  "C:/Users/arsbu/Desktop/work/main.py"
"""


def check_collide(output_list, ship_name, people_list, min_distance):
    if not output_list:
        return False
    if ship_name == output_list[-1][0]:
        return True
    if set(output_list[-1][1]) & set(people_list):
        return True
    for i in range(len(output_list) - 1, -1, -1):
        ship, pp, number = output_list[i]
        pp = pp.split(", ")
        rad = len(output_list) - i - 1
        if ship == ship_name:
            if rad < min_distance:
                return True
        if set(pp) & set(people_list) and rad < min_distance:
            return True
    return False


def put_to_end(data_list, index):
    data_to_put = data_list[index]
    data_list.pop(index)
    data_list.append(data_to_put)


def process(original_data):
    # original_data = data.values.tolist()

    data = original_data.copy()
    min_rad = 1
    verified_rads = {}
    original_min_rad = min_rad
    while True:
        out = []
        permutation_count = 0
        no_vars = 0
        while len(out) < len(original_data):
            ship_name, people, number = data[0]
            people_list = people.split(", ")
            if permutation_count > len(data):
                put_to_end(data, 0)
                permutation_count = 0
                no_vars += 1
            if no_vars > 3:
                data = original_data.copy()
                put_to_end(data, 0)
                out.clear()
                min_rad += 1
                break

            is_colliding = check_collide(out, ship_name, people_list, min_rad)
            if is_colliding:
                put_to_end(data, 0)
                permutation_count += 1
                continue

            out.append(data[0])
            data.pop(0)
            permutation_count = 0

        if len(out) == len(original_data):
            if min_rad not in verified_rads:
                verified_rads[min_rad] = out.copy()

        if min_rad > len(original_data) and original_min_rad:
            break
        data = original_data.copy()
        put_to_end(data, 0)

        min_rad += 1
    return verified_rads


def get_file_path():
    root = Tk()
    root.attributes('-topmost', True)
    root.withdraw()

    file_path = filedialog.askopenfilename(title="Выберите файл", filetypes=[("Excel файлы", "*.xlsx")])

    root.destroy()

    if file_path:
        print("Выбранный файл:", file_path)
        return file_path
    else:
        print("Файл не выбран")
        return None


def select_save_folder():
    root = Tk()
    root.attributes('-topmost', True)
    root.withdraw()

    folder_path = filedialog.askdirectory(
        title="Выберите папку для сохранения таблицы"
    )

    root.destroy()

    if folder_path:
        print(f"Выбрана папка для сохранения: {folder_path}")
        return folder_path
    else:
        print("Сохранение отменено")
        return None


run = True

while run:
    ships_column = None
    people_column = None
    table_filepath = None
    save_dir_path = None
    tabel_list_name = None
    number = None
    header = None
    while table_filepath is None:
        print("Выберите действие (укажите цифру):")
        print("1 - Открыть файл таблицы")
        print("Или напишите 'exit' для перезапуска програмы")

        control = input("Выберите действие: ")
        if control == "1":
            print("Ожидайте открытия окна выбора файла...")
            file_path = get_file_path()
            if file_path is not None:
                table_filepath = file_path
                print("Файл загружен!")
                break
            else:
                print("Ошибка загрузки файла повторите попытку!")
                continue
        elif control == "exit":
            print("\n" * 2)
            break
        print("Ошибка корректности вводимых данных")
        print()
    if table_filepath is None:
        print("\n" * 2)
        continue

    print()
    while tabel_list_name is None:
        print("Укажите название листа таблицы для анализа")
        print("Или напишите 'exit' для перезапуска програмы")
        control = input("Укажите просто название (регистр имеет значение): ")
        if control == "exit":
            print("\n" * 2)
            continue
        else:
            try:
                tabel_list_name = control
                sheet = pe.get_sheet(file_name=table_filepath, sheet_name=tabel_list_name)
                header = sheet.row[0]
                print("Сохранено!")
            except ValueError:
                print(f"Не удалось найти лист '{control}' в таблице!")
                print()
                continue
            except Exception:
                print(f"Ошибка открытия таблицы: {str(e)}")
                print()
                continue

        print()
    while ships_column is None:
        print("Укажите название столбца, где записаны названия суден")
        print("Или напишите 'exit' для перезапуска програмы")
        control = input("Укажите просто название (регистр имеет значение): ")
        if control == "exit":
            print("\n" * 2)
            continue
        if control in header:
            ships_column = control
        else:
            print(f"На листе '{tabel_list_name}' нет колонки '{control}'")
        print()

    while people_column is None:
        print("Укажите название столбца, где записаны имена участников\n"
              "(Формат данных столбца: разделитель - запятая и пробел, например 'Иван Иванов, Дарья Сергеева, Василий Андреев')")
        print("Или напишите 'exit' для перезапуска програмы")
        control = input("Укажите просто название (регистр имеет значение): ")
        if control == "exit":
            print("\n" * 2)
            continue
        if control in header:
            people_column = control
            print("Сохранено!")
        else:
            print(f"На листе '{tabel_list_name}' нет колонки '{control}'")
        print()
    while number is None:
        print("Укажите название столбца, где записан номер команды")
        print("Или напишите 'exit' для перезапуска програмы")
        control = input("Укажите просто название (регистр имеет значение): ")
        if control == "exit":
            print("\n" * 2)
            continue
        if control in header:
            number = control
            print("Сохранено!")
        else:
            print(f"На листе '{tabel_list_name}' нет колонки '{control}'")
        print()

    while save_dir_path is None:
        print("Выберите действие (укажите цифру):")
        print("1 - Выбрать папку сохранения таблицы")
        print("Или напишите 'exit' для перезапуска програмы")

        control = input("Выберите действие: ")
        if control == "1":
            print("Ожидайте открытия окна выбора файла...")
            dir_path = select_save_folder()
            if dir_path is not None:
                save_dir_path = dir_path
                print("Папка успешно выбрана!")
                print()
                break

        elif control == "exit":
            print("\n" * 2)
            break
        print("Ошибка корректности вводимых данных")
        print()
    if save_dir_path is None:
        continue
    print("Анализ таблицы....")

    # try:
    #     # Чтение данных с помощью pyexcel
    #     data = pe.get_array(file_name=table_filepath, sheet_name=tabel_list_name)
    #
    #     # Получаем индексы колонок
    #     header = data[0]
    #     ship_idx = header.index(ships_column)
    #     people_idx = header.index(people_column)
    #     number_idx = header.index(number)
    #
    #     # Извлекаем нужные колонки (пропускаем заголовок)
    #     extracted_data = []
    #     for row in data[1:]:
    #         extracted_data.append((row[ship_idx], row[people_idx], row[number_idx]))
    #
    # except Exception as e:
    #     print(f"Произошла ошибка открытия таблицы: {str(e)}")
    #     print("Попробуйте закрыть программу, в которой она открыта")
    #     continue

    try:
        sheet = pe.get_sheet(file_name=table_filepath, sheet_name=tabel_list_name)
        ship_idx = header.index(ships_column)
        people_idx = header.index(people_column)
        number_idx = header.index(number)

        extracted_data = []
        for row in sheet:
            if row == header:
                continue
            extracted_data.append((row[ship_idx], row[people_idx], row[number_idx]))

    except Exception as e:
        print(f"Ошибка открытия таблицы: {str(e)}")
        continue

    try:
        verified_rads = process(extracted_data)
        format_data = {}

        if not verified_rads:
            print("ПРОГРАММА НЕ СМОГЛА НАЙТИ НИ ОДНОГО ВАРИАНТА!")
            print("\n" * 2)
            continue

        book_dict = {}
        for rad, items in verified_rads.items():
            sheet_data = [["Участники", "Судно", "Номер"]]
            for ship, people, num in items:
                sheet_data.append([people, ship, num])
            book_dict[f"Минимальное расстояние - {rad}"] = sheet_data

        full_filepath = f"{save_dir_path}/Результат составления протакола {random.randint(0, 1000000)}.xlsx"
        pe.save_book_as(bookdict=book_dict, dest_file_name=full_filepath)

        print(f"Результат анализа успешно сохранён в файл {full_filepath}")
        print("#####################\n" * 2)

    except Exception as e:
        print(f"Ошибка обработки: {str(e)}")
