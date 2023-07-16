import sys
import os
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

    Sheet1 = Sheet1.drop_duplicates()

    Sheet1 = Sheet1.dropna(subset=['Документ сбыта'])

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
        Sheet1['Нарушение предварительно'] = Sheet1.apply(lambda x: 'да' if (pd.notnull(x['Отклонение']) and x['Отклонение'] > 0) else 'нет', axis=1)
        Sheet1['Отклонение'] = Sheet1['Отклонение'].fillna(value='нет данных')
        Sheet1.loc[Sheet1['Отклонение'] == 'нет данных', 'Нарушение предварительно'] = 'да'

        Sheet2 = pd.DataFrame()

        docs = []
        if Sheet1.loc[Sheet1['Нарушение предварительно'] == 'да'].shape[0] > 0:
            docs = Sheet1.loc[Sheet1['Нарушение предварительно'] == 'да', 'Документ сбыта'].tolist()
            Sheet2['Торговый документ'] = docs

        Sheet2['Торговый документ подрядчик'] = None

        Sheet3 = pd.DataFrame()

        # values_list = Sheet2['Торговый документ'].tolist()
        # Sheet1 = Sheet1[~Sheet1['Наименование завода'].isin(values_list)]

        # output_file = createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
        # Sheet1.to_excel(output_file, index=False)

        output_file_excel = createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
        output_file_html = os.path.splitext(output_file_excel)[0] + '.html'

        with pd.ExcelWriter(output_file_excel) as writer:
            Sheet1.to_excel(writer, sheet_name="Sheet1", index=False)
            Sheet2.to_excel(writer, sheet_name="Sheet2", index=False)
            Sheet3.to_excel(writer, sheet_name="Sheet3", index=False)

        Sheet1.to_html(output_file_html, index=False)

        # Перезапись листа
        # with pd.ExcelWriter("path_to_file.xlsx", mode="a", engine="openpyxl") as writer:
        # df.to_excel(writer, sheet_name="Sheet3")  

        # pivot = pd.pivot_table(
        #     data,
        #     values=['summ'],
        #     index=['division', 'name'], 
        #     aggfunc='sum',
        #     columns='periods'
        # )
        # pivot.loc[('Total', 'Total'), :] = pivot.sum(axis=0)
        # pivot.loc[('Total', 'Total'), 'summ'] = 'Total'
        # df = pivot.apply(sum_all. axis=1)

        print(True)

if len(sys.argv) < 2:
    # нет файлов
    print(False)
else:
    # sys.argv[1] - загрузим первый файл, если их несколько то нужно загружать их в цикле for arg in sys.argv[1:]: 
    excel_file = createEnvPath('SAVED_FILES_PATH', sys.argv[1])
    Sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')

    if 'Назв. вида поставки' in Sheet1.columns:
        if Sheet1['Назв. вида поставки'].iloc[0] == 'Отпуск груза':
            run_script(sys.argv[1])
        else:
            print(False)
    else:
        print(False)
