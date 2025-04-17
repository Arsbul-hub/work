import pandas as pd


def check_collide(output_list, ship_name, people_list, min_distance):
    if not output_list:
        return False
    if ship_name == output_list[-1][0]:
        return True
    if set(output_list[-1][1]) & set(people_list):
        return True
    for i in range(len(output_list) - 1, -1, -1):
        ship, pp = output_list[i]
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
    out = []
    data = original_data.copy()
    min_rad = 2
    verified_rads = {}
    original_min_rad = min_rad
    permutation_count = 0
    no_vars = 0
    work_with_lower_rad = False
    while True:
        if min_rad == 3 and verified_rads:
            pass

        while len(out) < len(original_data):
            ship_name, people = data[0]
            people_list = people.split(", ")
            if permutation_count > len(data):
                put_to_end(data, 0)
                permutation_count = 0
                no_vars += 1
            if no_vars > 3:
                data = original_data.copy()
                put_to_end(data, 0)
                out.clear()
                permutation_count = 0
                no_vars = 0
                if work_with_lower_rad:
                    min_rad -= 1
                else:
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
            if not min_rad in verified_rads:
                verified_rads[min_rad] = out.copy()
            data = original_data.copy()
            put_to_end(data, 0)
            out.clear()
            permutation_count = 0
            no_vars = 0
            if work_with_lower_rad:
                min_rad -= 1
            else:
                min_rad += 1
        if min_rad <= 0 and work_with_lower_rad and original_min_rad:
            break
        elif min_rad > len(original_data):
            work_with_lower_rad = True
    return verified_rads


sheets = pd.read_excel('data_original.xlsx', sheet_name="Лист1")
data = pd.DataFrame(sheets, columns=["Судно", "Экипаж"])
original_data = data.values.tolist()
verified_rads = process(original_data)
format_data = {}
for rad, data in verified_rads.items():
    for ship, people in data:
        if rad not in format_data:
            format_data[rad] = {}
        format_data[rad][ship] = people

try:
    with pd.ExcelWriter("multiple.xlsx") as writer:
        for rad, data in format_data.items():
            df = pd.DataFrame([(k, v) for k, v in data.items()],
                              columns=['Судно', 'Участники'])

            df = df.sort_values('Судно')

            visible = True if min(format_data.keys()) == rad else False

            df.to_excel(writer, sheet_name=f"Минимально растояние - {rad}", index=False)
except PermissionError:
    print("Закройте файл и повторите попытку!")
