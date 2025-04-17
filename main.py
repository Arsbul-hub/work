import pandas as pd
from openpyxl import load_workbook
import random
from numba import jit

sheets = pd.read_excel('data.xlsx')
data = pd.DataFrame(sheets, columns=["Название судна", 'ФИО уастников'])
original_data = data.values.tolist()
out = []


def check_collide(output_list, ship_name, people_list, min_distance):
    if not output_list:
        return False
    for i in range(len(output_list) - 1, -1, -1):
        ship, pp = output_list[i]
        pp = pp.split(";")
        if ship == ship_name:
            rad = len(output_list) - i - 1
            if rad < min_distance:
                return True
        for p in pp:
            if p in people_list:
                rad = len(output_list) - i - 1
                if rad < min_distance:
                    return True
    return False


def put_to_end(data_list, index):
    data_to_put = data_list[index]
    data_list.pop(index)
    data_list.append(data_to_put)


for s, df in sheets.items():
    data = original_data.copy()
    min_rad = 1
    permutation_count = 0
    no_vars = False
    while True:
        while len(out) < len(original_data):
            ship_name, people = data[0]
            people_list = people.split(";")
            if permutation_count > len(data):

                # data = original_data.copy()
                put_to_end(data, 0)
                permutation_count = 0
                no_vars += 1
            if no_vars > 3:

                data = original_data.copy()
                put_to_end(data, 0)
                out.clear()
                permutation_count = 0
                min_rad -= 1
                break
            is_colliding = check_collide(out, ship_name, people_list, min_rad)
            if is_colliding:
                put_to_end(data, 0)
                permutation_count += 1
                continue

            out.append(data[0])
            data.pop(0)
            permutation_count = 0
        if len(out) == len(original_data) or min_rad < 1:
            break
        data = original_data.copy()
        put_to_end(data, 0)
        out.clear()
        permutation_count = 0
        min_rad -= 1

print(out)
