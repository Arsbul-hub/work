print("Анализ таблицы....")
try:
    # Чтение данных с помощью pyexcel
    data = pe.get_array(file_name=table_filepath, sheet_name=tabel_list_name)

    # Получаем индексы колонок
    header = data[0]
    ship_idx = header.index(ships_column)
    people_idx = header.index(people_column)
    number_idx = header.index(number)

    # Извлекаем нужные колонки (пропускаем заголовок)
    extracted_data = []
    for row in data[1:]:
        extracted_data.append((row[ship_idx], row[people_idx], row[number_idx]))

except Exception as e:
    print(f"Произошла ошибка открытия таблицы: {str(e)}")
    print("Попробуйте закрыть программу, в которой она открыта")
    continue

try:
    # Чтение данных
    sheet = pe.get_sheet(file_name=table_filepath, sheet_name=tabel_list_name)

    # Получаем индексы колонок
    header = sheet.row[0]
    ship_idx = header.index(ships_column)
    people_idx = header.index(people_column)
    number_idx = header.index(number)

    # Извлекаем нужные колонки
    extracted_data = []
    for row in sheet:
        if row == header:  # Пропускаем заголовок
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

    # Формируем данные для сохранения
    book_dict = {}
    for rad, items in verified_rads.items():
        sheet_data = [["Участники", "Судно", "Номер"]]
        for ship, people, num in items:
            sheet_data.append([people, ship, num])
        book_dict[f"Дистанция {rad}"] = sheet_data

    # Сохраняем в Excel
    full_filepath = os.path.join(save_dir_path,
                                 f"Результат составления протакола {random.randint(0, 1000000)}.xlsx")
    pe.save_book_as(bookdict=book_dict, dest_file_name=full_filepath)

    print(f"Результат анализа успешно сохранён в файл {full_filepath}")
    print("#####################\n" * 2)

except Exception as e:
    print(f"Ошибка обработки: {str(e)}")