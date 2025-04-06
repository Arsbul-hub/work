import pandas as pd
from openpyxl import load_workbook
import random

sheets = pd.read_excel('data.xlsx', sheet_name=None)
out = []
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
    c = len(ship_names)
    # first_index = random.randint(0, len(ship_names) - 1)
    cant_use = []
    pers = 0
    while len(out) < c:
        ship_name = ship_names[0]
        people = ships_people[0]

        people_list = people.split(";")
        if pers > len(ship_names):
            ship_names = ship_names_first.copy()
            ships_people = ships_people_first.copy()
            name = ship_names[0]
            sp = ships_people[0]
            ship_names.pop(0)
            ship_names.append(name)
            ships_people.pop(0)
            ships_people.append(sp)
            pers = 0
        flag = False
        for i in range(len(out) - 1, -1, -1):
            ship, pp = out[i]
            if ship == ship_name:
                rad = len(out) - i - 1
                if rad < 2:
                    ship_names.remove(ship_name)
                    ship_names.append(ship_name)
                    ships_people.remove(people)
                    ships_people.append(people)
                    flag = True
                    break
        if flag:
            pers += 1
            continue

        out.append((ship_name, people_list))
        ship_names.remove(ship_name)
        ships_people.remove(people)
        pers = 0

print(out)
