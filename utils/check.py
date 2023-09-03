import pandas as pd

# проверка есть ли все дивизионы по колонке 'Назн.:Звд', check_value_in_list_and_set_value устанавливает 'нет значения' в 'див' если не нашел
# row - исходная таблица
# row_name - название столбца
# items_list - список столбцов для которых нужно сделать новые значения
# default_value - значение если в списке не найдей ключ
def check_value_in_list_and_set_value(row, row_name, items_list, default_value=None):
    for key, value in items_list.items():
        if row[row_name] == key:
            return value
    if default_value:
        return default_value
    else:
        return 'нет значения'

def check_value_in_row_and_set_value(row, row_name_1, row_name_2, row_value_name, default_value=None):
    if row[row_name_1] in row[row_name_2]:
        return row[row_value_name]
    if default_value:
        return default_value
    else:
        return 'нет значения'        