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
import openpyxl
import xlsxwriter
import pandas as pd
import numpy as np
from dotenv import load_dotenv
load_dotenv()
pd.options.mode.chained_assignment = None

def run_script(current_data, output_file_excel, output_file_html):
    Sheet1 = current_data

    Sheet1 = Sheet1.drop_duplicates()

    Sheet1 = Sheet1.dropna(subset=['Документ сбыта'])

    values_to_drop = utils.load_settings_table_column_values('удалить.json', 'Доп склад')
    Sheet1 = Sheet1[~Sheet1['Наименование завода'].isin(values_to_drop)]

    values_to_add_rp = utils.load_settings_table_column_values('дивизионы.json', 'РП')
    values_to_add_div = utils.load_settings_table_column_values('дивизионы.json', 'Дивизион')
    items = dict(zip(values_to_add_rp, values_to_add_div))
    Sheet1['Дивизион'] = Sheet1.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наименование завода', items_list=items)

    Sheet1 = Sheet1.loc[Sheet1['Дивизион'] != 'Ритейл']

    unknowns = Sheet1.loc[Sheet1['Дивизион'] == 'нет значения', 'Наименование завода'].tolist()
    if len(unknowns) > 0:
        error_json = utils.createEnvPath('SAVED_ERRPR_PATH', 'unknowns_division')
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

load_obj = utils.load_file_obj(sys.argv[1:])
output_file_excel = load_obj["output_file_excel"]
output_file_html = load_obj["output_file_html"]
files = load_obj["files"]

isExistData = False
current_data = None

for file in files:
    for sheet_name, sheet in file.items():
        if 'Назв. вида поставки' in sheet.columns:
            if sheet['Назв. вида поставки'].iloc[0] == 'Отпуск груза':
                isExistData = True
                current_data = sheet

if isExistData:
    run_script(current_data, output_file_excel, output_file_html)
else:
    print(False)
