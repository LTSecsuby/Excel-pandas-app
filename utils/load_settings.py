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
import pandas as pd

# возвращает список значений указанной настройки и указанного в ней столбца
# settings_name - название таблицы
# column_name - название столбца
def load_settings_table_column_values(settings_name, column_name):
    json_file = utils.createEnvPath('SAVED_SETTINGS_FILES_PATH', settings_name)
    column_values = []
    with open(json_file, encoding="utf-8") as f:
        load_json = json.load(f)
        for table in load_json['table']:
            if table['key'] == column_name:
                column_values = table['values']

    return column_values

def load_file_obj(files):
    file_name = files[0]
    res = {}
    tables = []
    for file in files:
        excel_file = utils.createEnvPath('SAVED_FILES_PATH', file)
        current_data = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl')
        tables.append(current_data)

    output_file_excel = utils.createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
    output_file_html = os.path.splitext(output_file_excel)[0] + '.html'
    res["output_file_excel"] = output_file_excel
    res["output_file_html"] = output_file_html
    res["files"] = tables

    return res