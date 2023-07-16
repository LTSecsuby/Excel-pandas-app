import sys
import os
PROJECT_ROOT = os.path.abspath(os.path.join(
                  os.path.dirname(__file__), 
                  os.pardir)
)
sys.path.append(PROJECT_ROOT)
import utils

import json
import openpyxl
import pandas as pd
import numpy as np
from dotenv import load_dotenv
load_dotenv()
pd.options.mode.chained_assignment = None

def createEnvPath(env_path, last = None):
    if os.getenv('MODE') == 'production':
        if last:
            return os.path.join(os.getcwd(), 'dist', os.getenv(env_path), last)
        return os.path.join(os.getcwd(), 'dist', os.getenv(env_path))
    else:
        if last:
            return os.path.join(os.getcwd(), os.getenv(env_path), last)
    return os.path.join(os.getcwd(), os.getenv(env_path))

def check_div(row, items_list):
    for key, value in items_list.items():
        if row['Наименование завода'] == key:
            return value
    return 'нет значения'

def run_script(file_name):
    excel_file = createEnvPath('SAVED_FILES_PATH', file_name)

    Sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')
    encoding = 'utf-8'

    Sheet1 = Sheet1.dropna(subset=['Документ сбыта'])

    Sheet1 = Sheet1.sort_values(by='Фактическая дата', kind="mergesort")

    Sheet1 = Sheet1.drop_duplicates(subset='Документ сбыта')
    

    json_file = createEnvPath('SAVED_SETTINGS_FILES_PATH', 'удалить.json')
    values_to_drop = []
    with open(json_file, encoding="utf-8") as f:
        load_json = json.load(f)
        values_to_drop = load_json['table'][0]['values']
    Sheet1 = Sheet1[~Sheet1['Наименование завода'].isin(values_to_drop)]

    # divisions = pd.DataFrame()
    json_file = createEnvPath('SAVED_SETTINGS_FILES_PATH', 'дивизионы.json')
    values_to_add_rp = []
    values_to_add_div = []
    with open(json_file, encoding="utf-8") as f:
        load_json = json.load(f)
        values_to_add_rp = load_json['table'][0]['values']
        values_to_add_div = load_json['table'][1]['values']
        items = dict(zip(values_to_add_rp, values_to_add_div))
        Sheet1['Дивизион'] = Sheet1.apply(check_div, axis=1, items_list=items)
        Sheet1 = Sheet1.loc[Sheet1['Дивизион'] != 'Ритейл']
        # divisions['РП'] = values_to_add_rp
        # divisions['Дивизион'] = values_to_add_div

    unknowns = Sheet1.loc[Sheet1['Дивизион'] == 'нет значения', 'Наименование завода'].tolist()
    if len(unknowns) > 0:
        error_json = createEnvPath('SAVED_ERRPR_PATH', 'unknowns_division')
        output_file_json = os.path.splitext(error_json)[0] + '.json'
        error = pd.DataFrame()
        error['error'] = unknowns
        with open(output_file_json, 'w', encoding='utf-8') as file:
            error.to_json(output_file_json, force_ascii=False)
        print('unknowns_division')
    else:
        Sheet1['Отклонение'] = Sheet1.apply(lambda x: x['Фактическая дата'] - x['Плановая дата ПО'] if not pd.isnull(x['Плановая дата ПО']) and not pd.isnull(x['Фактическая дата']) else pd.NaT, axis=1)
        Sheet1['Отклонение'] = Sheet1['Отклонение'].apply(lambda x: x.total_seconds() / 86400 if not pd.isnull(x) else np.nan)
        Sheet1['Нарушение'] = Sheet1.apply(lambda x: 'да' if (pd.notnull(x['Отклонение']) and x['Отклонение'] > 0) else 'нет', axis=1)
        Sheet1['Отклонение'] = Sheet1['Отклонение'].fillna(value='нет данных')
        Sheet1.loc[Sheet1['Отклонение'] == 'нет данных', 'Нарушение'] = 'да'

        Sheet2 = pd.DataFrame()
        Sheet2 = pd.pivot_table(Sheet1,
                index=['Дивизион', 'Наименование завода'],
                columns='Нарушение',
                values='Назв. вида поставки',
                aggfunc='count')

        Sheet2['да'] = Sheet2['да'].fillna(0)
        Sheet2['нет'] = Sheet2['нет'].fillna(0)

        Sheet2 = Sheet2.reset_index()
        
        Sheet3 = utils.add_total_by_field(Sheet2, 'Дивизион', ['да', 'нет'])

        total_fact = Sheet2['да'].sum()
        total_cost = Sheet2['нет'].sum()
        total_count = total_fact + total_cost
        total_row = pd.DataFrame({'Наименование завода': [''],
                                'Дивизион': ['Общий итог'],
                                'да': [total_fact],
                                'нет': [total_cost],
                                'Общий итог': [total_count]})
        Sheet3 = pd.concat([Sheet3, total_row])

        sum_column = Sheet3['да'] + Sheet3['нет']
        Sheet3['Общий итог'] = sum_column

        percentage_column = (Sheet3['нет'] / Sheet3['Общий итог']) * 100
        Sheet3['Процент %'] = percentage_column

        output_file_excel = createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
        output_file_html = os.path.splitext(output_file_excel)[0] + '.html'

        with pd.ExcelWriter(output_file_excel) as writer:
            Sheet1.to_excel(writer, sheet_name="Sheet1", index=False)
            Sheet3.to_excel(writer, sheet_name="Sheet2", index=False)

        Sheet3.to_html(output_file_html, index=False)

        print(True)

if len(sys.argv) < 2:
    # нет файлов
    print(False)
else:
    # sys.argv[1] - загрузим первый файл, если их несколько то нужно загружать их в цикле for arg in sys.argv[1:]:
    excel_file = createEnvPath('SAVED_FILES_PATH', sys.argv[1])

    Sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')

    if 'Назв. вида поставки' in Sheet1.columns:
        if Sheet1['Назв. вида поставки'].iloc[0] == 'Приём груза':
            run_script(sys.argv[1])
        else:
            print(False)
    else:
        print(False)
