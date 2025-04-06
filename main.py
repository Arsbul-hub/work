import pandas as pd
from openpyxl import load_workbook
import random

sheets = pd.read_excel('data.xlsx', sheet_name=None)
out = []


def check_collide(output_list, ship_name, people_list, min_distance):
    for i in range(len(output_list) - 1, -1, -1):
        ship, pp = output_list[i]
        if ship == ship_name:
            rad = len(output_list) - i - 1
            if rad <= min_distance:
                return True
        for p in pp:
            if p in people_list:
                rad = len(output_list) - i - 1
                if rad <= min_distance:
                    return True

def put_to_end(data_list, index):
    data_list.pop(index)
    data_list.append(name)

for s, df in sheets.items():
    # data_head = df.head()
    # data_row = df.iloc()
    # d = data_head.values
    # ship_names = list(map(lambda a: a[0], data_head.values.tolist()))
    # ships_people = data_head["ФИО уастников"].to_list()
    ship_names_first = ["Лодка 1",
                        "Лодка 2",
                        "Лодка 5",
                        "Лодка 2",
                        "Лодка 2",
                        "Лодка 3",
                        "Лодка 4",
                        "Лодка 3",
                        "Лодка 4",
                        ]

    ships_people_first = ["Андрей Андреев;Иван Иванович;Кирил Кириллович",
                          "Михаил Михайлович;Арсений Арсеньеви;Илья Ильич",
                          "Вадим Вадимович;Иван Иванович",
                          "Андрей Андреев;Иван Иванович;Кирил Кириллович",
                          "Андрей Андреев;Иван Иванович;Кирил Кириллович",
                          "Михаил Михайлович;Асений Арсеньеви;Илья Ильич",
                          "Андрей Андреев;Иван Иванович;Кирил Кириллович",
                          "Михаил Михайлович;Арсений Арсеньеви;Илья Ильич",
                          "Михаил Михайлович;Асений Арсеньеви;Илья Ильич",
                          ]
    ship_names = ship_names_first.copy()
    ships_people = ships_people_first.copy()
    min_rad = 1
    permutation_count = 0
    while len(out) < len(ship_names):
        ship_name = ship_names[0]
        people = ships_people[0]
        people_list = people.split(";")
        if permutation_count > len(ship_names):
            ship_names = ship_names_first.copy()
            ships_people = ships_people_first.copy()
            put_to_end(ship_names, 0)
            put_to_end(ships_people, 0)
            permutation_count = 0
        is_colliding = check_collide(out, ship_name, people_list, min_rad)
        if is_colliding:
            permutation_count += 1
            continue

        out.append((ship_name, people_list))
        ship_names.remove(ship_name)
        ships_people.remove(people)
        permutation_count = 0

print(out)