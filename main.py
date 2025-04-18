import random
import time
from tkinter import filedialog, Tk

import pandas as pd


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


def process(data):
    original_data = data.values.tolist()

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
    print("Укажите название листа таблицы для анализа")
    print("Или напишите 'exit' для перезапуска програмы")
    control = input("Укажите просто название (регистр имеет значение): ")
    if control:
        tabel_list_name = control
        print("Сохранено!")
    elif control == "exit":
        print("\n" * 2)
        continue
    print()
    print("Укажите название столбца, где записаны названия суден")
    print("Или напишите 'exit' для перезапуска програмы")
    control = input("Укажите просто название (регистр имеет значение): ")
    if control:
        ships_column = control
    elif control == "exit":
        print("\n" * 2)
        continue
    print()
    print("Укажите название столбца, где записаны имена участников\n"
          "(Формат данных столбца: разделитель - запятая и пробел, например 'Иван Иванов, Дарья Сергеева, Василий Андреев')")
    print("Или напишите 'exit' для перезапуска програмы")
    control = input("Укажите просто название (регистр имеет значение): ")
    if control:
        people_column = control
        print("Сохранено!")
    elif control == "exit":
        print("\n" * 2)
        continue
    print()
    print("Укажите название столбца, где записан номер команды")
    print("Или напишите 'exit' для перезапуска програмы")
    control = input("Укажите просто название (регистр имеет значение): ")
    if control:
        number = control
        print("Сохранено!")
    elif control == "exit":
        print("\n" * 2)
        continue
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
    try:
        sheets = pd.read_excel(table_filepath, sheet_name=tabel_list_name)
        data = pd.DataFrame(sheets, columns=[ships_column, people_column, number])

    except:
        print("Произошла ошибка открытия таблицы (попробуйте закрыть програму, в которой она открыта)")
        continue
    try:
        verified_rads = process(data)
        format_data = {}
    except:
        print("Ошибка анализа данных.")
        continue
    try:
        for rad, data in verified_rads.items():
            for ship, people, number in data:
                if rad not in format_data:
                    format_data[rad] = []
                format_data[rad].append((people, ship, number))
        full_filepath = f"{save_dir_path}/Результат составления протакола {random.randint(0, 1000000)}.xlsx"
        with pd.ExcelWriter(full_filepath) as writer:
            for rad, data in format_data.items():
                df = pd.DataFrame([(p, s, n) for p, s, n in data],
                                  columns=["Участники", "Судно", "Номер"])

                visible = True if min(format_data.keys()) == rad else False

                df.to_excel(writer, sheet_name=f"Минимальное растояние - {rad}", index=False)
            print(f"Результат анализа успешно сохранён в файл {full_filepath}")
            print("#####################\n" * 2)
    except:
        print("Ошибка сохранения данных в таблицу")
        print("\n" * 2)
        continue
