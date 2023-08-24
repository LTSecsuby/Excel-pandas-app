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

def run_script(file_name):
    excel_file = utils.createEnvPath('SAVED_FILES_PATH', file_name)

    Sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')
    encoding = 'utf-8'

    Sheet1 = Sheet1.dropna(subset=['Документ сбыта'])

    Sheet1 = Sheet1.sort_values(by='Фактическая дата', kind="mergesort")

    Sheet1 = Sheet1.drop_duplicates(subset='Документ сбыта')
    

    json_file = utils.createEnvPath('SAVED_SETTINGS_FILES_PATH', 'удалить.json')
    values_to_drop = []
    with open(json_file, encoding="utf-8") as f:
        load_json = json.load(f)
        values_to_drop = load_json['table'][0]['values']
    Sheet1 = Sheet1[~Sheet1['Наименование завода'].isin(values_to_drop)]

    # divisions = pd.DataFrame()
    json_file = utils.createEnvPath('SAVED_SETTINGS_FILES_PATH', 'дивизионы.json')
    values_to_add_rp = []
    values_to_add_div = []
    with open(json_file, encoding="utf-8") as f:
        load_json = json.load(f)
        values_to_add_rp = load_json['table'][2]['values']
        values_to_add_div = load_json['table'][3]['values']
        items = dict(zip(values_to_add_rp, values_to_add_div))

        Sheet1['Дивизион'] = Sheet1.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наименование завода', items_list=items, default_value='Пустой дивизион')

        Sheet1 = Sheet1.loc[Sheet1['Дивизион'] != 'Ритейл']
        # divisions['РП'] = values_to_add_rp
        # divisions['Дивизион'] = values_to_add_div

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
        
        # Sheet3 = utils.add_total_by_field(Sheet2, 'Наименование завода', 'Дивизион', ['да', 'нет'])

        Sheet3 = pd.DataFrame(columns=Sheet2.columns)
        new_rows = []

        previous_division = Sheet2.iloc[0]['Дивизион']

        division_list = []
        division_list.append(previous_division)

        last = None
        previous_sum1 = 0
        previous_sum2 = 0

        for index, row in Sheet2.iterrows():
            current_division = row['Дивизион']
            last = row['Дивизион']

            if current_division != previous_division:
                new_row = {'Дивизион': None, 
                        'Наименование завода': previous_division, 
                        'да': previous_sum1,
                        'нет': previous_sum2 }
                new_rows.append(new_row)
                division_list.append(current_division)
                
                previous_division = current_division
                previous_sum1 = 0
                previous_sum2 = 0
                previous_sum1 += int(row['да'])
                previous_sum2 += int(row['нет'])
            else:
                previous_sum1 += int(row['да'])
                previous_sum2 += int(row['нет'])

        last_row_dict = {'Дивизион': None, 
                    'Наименование завода': last, 
                    'да': previous_sum1,
                    'нет': previous_sum2 }
        new_rows.append(last_row_dict)
        division_list.append(last)

        new_rows_index = 0

        Sheet3.loc[len(Sheet3)] = new_rows[new_rows_index]
        new_rows_index = new_rows_index + 1
        previous_division = Sheet2.iloc[0]['Дивизион']

        for index, row in Sheet2.iterrows():
            current_division = row['Дивизион']

            if current_division != previous_division:
                Sheet3.loc[len(Sheet3)] = new_rows[new_rows_index]
                previous_division = current_division
                new_rows_index = new_rows_index + 1
                Sheet3.loc[len(Sheet3)] = row
            else:
                Sheet3.loc[len(Sheet3)] = row

        total_fact = Sheet2['да'].sum()
        total_cost = Sheet2['нет'].sum()
        total_count = total_fact + total_cost
        total_row = pd.DataFrame({'Наименование завода': ['Общий итог'],
                                'Дивизион': [''],
                                'да': [total_fact],
                                'нет': [total_cost],
                                'Общий итог': [total_count]})
        Sheet3 = pd.concat([Sheet3, total_row])

        sum_column = Sheet3['да'] + Sheet3['нет']
        Sheet3['Общий итог'] = sum_column

        percentage_column = (Sheet3['нет'] / Sheet3['Общий итог'])
        Sheet3['Процент %'] = percentage_column.round(4)

        Sheet3['да'] = Sheet3['да'].apply(lambda x: round(x)).astype(int)
        Sheet3['нет'] = Sheet3['нет'].apply(lambda x: round(x)).astype(int)
        Sheet3['Общий итог'] = Sheet3['Общий итог'].apply(lambda x: round(x)).astype(int)

        division_list.append('Общий итог')

        Sheet3 = Sheet3.drop(['Дивизион'], axis=1)

        output_file_excel = utils.createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
        output_file_html = os.path.splitext(output_file_excel)[0] + '.html'

        with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
            book = writer.book
            percent_format = book.add_format({'num_format': '0.00%'})
            bold_format = book.add_format({'bold': True})

            Sheet1.to_excel(writer, sheet_name="Sheet1", index=False)
            Sheet3.to_excel(writer, sheet_name="Sheet2", index=False)
            worksheet = writer.sheets["Sheet2"]

            worksheet.set_column('E:E', 18, percent_format)
            worksheet.set_column('A:A', 28)
            worksheet.set_column('B:D', 18)

            column_values = Sheet3['Наименование завода'].values.tolist()

            for row_num, value in enumerate(column_values):
                if value in division_list:
                    worksheet.write(row_num + 1, 0, value, bold_format)
                else:
                    worksheet.write(row_num + 1, 0, value)

        Sheet3.to_html(output_file_html, index=False)

        print(True)

if len(sys.argv) < 2:
    # нет файлов
    print(False)
else:
    # sys.argv[1] - загрузим первый файл, если их несколько то нужно загружать их в цикле for arg in sys.argv[1:]:
    excel_file = utils.createEnvPath('SAVED_FILES_PATH', sys.argv[1])

    Sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')

    if 'Назв. вида поставки' in Sheet1.columns:
        if Sheet1['Назв. вида поставки'].iloc[0] == 'Приём груза':
            run_script(sys.argv[1])
        else:
            print(False)
    else:
        print(False)
