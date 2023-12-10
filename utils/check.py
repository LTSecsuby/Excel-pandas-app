import pandas as pd
import sys
import os
# для работы import utils нужно подтянуть пути проекта
PROJECT_ROOT = os.path.abspath(os.path.join(
                  os.path.dirname(__file__), 
                  os.pardir)
)
sys.path.append(PROJECT_ROOT)
import utils
import json

# проверка есть ли все дивизионы по колонке 'Назн.:Звд', check_value_in_list_and_set_value устанавливает 'нет значения' в 'див' если не нашел
# row - исходная таблица
# row_name - название столбца
# items_list - список столбцов для которых нужно сделать новые значения
# default_value - значение если в списке не найдей ключ
def check_value_in_list_and_set_value_lists(row, row_name, items_list_key, items_list_value, default_value=None):
    for index, value in enumerate(items_list_key):
        if row[row_name] == value:
            return items_list_value[index]
    if default_value:
        return default_value
    else:
        return 'нет значения'

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

def check_null_pointer_in_table_value(table, find_column_name, value_column_name, null_pointer_name):
    unknowns = table.loc[table[find_column_name] == null_pointer_name, value_column_name].tolist()
    if len(unknowns) > 0:
        error_json = utils.createEnvPath('SAVED_ERRPR_PATH', 'unknowns_division')
        output_file_json = os.path.splitext(error_json)[0] + '.json'
        error = pd.DataFrame()
        error['error'] = unknowns
        error['name'] = find_column_name
        error = error.drop_duplicates(subset='error')
        with open(output_file_json, 'w', encoding='utf-8') as file:
            error.to_json(output_file_json, force_ascii=False)
        return True
    else:
        return False