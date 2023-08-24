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
    encoding = 'utf-8'

    Sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')
    Sheet2 = pd.read_excel(excel_file, sheet_name='Sheet2', engine='openpyxl')
    Sheet3 = pd.DataFrame()

    if not Sheet2['Торговый документ подрядчик'].isnull().values.all():
        docs = Sheet2['Торговый документ подрядчик'].dropna().tolist()
    else:
        contractors_table = pd.read_excel(excel_file, sheet_name='Sheet3', engine='openpyxl')
        contractors_table = contractors_table.drop_duplicates(subset='Торговый документ')
        filtered_rows = contractors_table.loc[contractors_table['Доставка подрядчиком'] == 'ИСТИННО']
        docs = filtered_rows['Торговый документ'].dropna().tolist()

    Sheet2['Торговый документ подрядчик'] = pd.Series(docs)

    Sheet1['Доставка подрядчиком'] = Sheet1['Документ сбыта'].apply(lambda x: 'Доставка подрядчиком' if x in docs else None)

    Sheet1['Нарушение итог'] = 'нет'
    mask = (Sheet1['Доставка подрядчиком'].isna()) & (Sheet1['Нарушение предварительно'] == 'да')
    Sheet1.loc[mask, 'Нарушение итог'] = 'да'

    Sheet3 = pd.pivot_table(Sheet1,
        index=['Дивизион', 'Наименование завода'],
        columns='Нарушение итог',
        values='Назв. вида поставки',
        aggfunc='count')

    Sheet3['да'] = Sheet3['да'].fillna(0)
    Sheet3['нет'] = Sheet3['нет'].fillna(0)

    Sheet3 = Sheet3.reset_index()
    
    # Sheet4 = utils.add_total_by_field(Sheet3, 'Дивизион', ['да', 'нет'])

    Sheet4 = pd.DataFrame(columns=Sheet3.columns)
    new_rows = []

    previous_division = Sheet3.iloc[0]['Дивизион']

    division_list = []
    division_list.append(previous_division)

    last = None
    previous_sum1 = 0
    previous_sum2 = 0

    for index, row in Sheet3.iterrows():
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

    Sheet4.loc[len(Sheet4)] = new_rows[new_rows_index]
    new_rows_index = new_rows_index + 1
    previous_division = Sheet3.iloc[0]['Дивизион']

    for index, row in Sheet3.iterrows():
        current_division = row['Дивизион']

        if current_division != previous_division:
            Sheet4.loc[len(Sheet4)] = new_rows[new_rows_index]
            previous_division = current_division
            new_rows_index = new_rows_index + 1
            Sheet4.loc[len(Sheet4)] = row
        else:
            Sheet4.loc[len(Sheet4)] = row

    total_fact = Sheet3['да'].sum()
    total_cost = Sheet3['нет'].sum()
    total_count = total_fact + total_cost
    total_row = pd.DataFrame({'Наименование завода': ['Общий итог'],
                            'Дивизион': [''],
                            'да': [total_fact],
                            'нет': [total_cost],
                            'Общий итог': [total_count]})
    Sheet4 = pd.concat([Sheet4, total_row])

    sum_column = Sheet4['да'] + Sheet4['нет']
    Sheet4['Общий итог'] = sum_column

    percentage_column = (Sheet4['нет'] / Sheet4['Общий итог'])
    Sheet4['Процент %'] = percentage_column.round(4)

    Sheet4['да'] = Sheet4['да'].apply(lambda x: round(x)).astype(int)
    Sheet4['нет'] = Sheet4['нет'].apply(lambda x: round(x)).astype(int)
    Sheet4['Общий итог'] = Sheet4['Общий итог'].apply(lambda x: round(x)).astype(int)

    division_list.append('Общий итог')

    Sheet4 = Sheet4.drop(['Дивизион'], axis=1)

    # Sheet3.columns = [' '.join(col).strip() for col in Sheet3.columns.values]
    # Sheet3 = Sheet3.reset_index()

    output_file_excel = utils.createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
    output_file_html = os.path.splitext(output_file_excel)[0] + '.html'

    with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
        book = writer.book
        percent_format = book.add_format({'num_format': '0.00%'})
        bold_format = book.add_format({'bold': True})

        Sheet1.to_excel(writer, sheet_name="Sheet1", index=False)
        Sheet2.to_excel(writer, sheet_name="Sheet2", index=False)
        Sheet4.to_excel(writer, sheet_name="Sheet3", index=False)
        worksheet = writer.sheets["Sheet3"]

        worksheet.set_column('E:E', 18, percent_format)
        worksheet.set_column('A:A', 28)
        worksheet.set_column('B:D', 18)

        column_values = Sheet4['Наименование завода'].values.tolist()

        for row_num, value in enumerate(column_values):
            if value in division_list:
                worksheet.write(row_num + 1, 0, value, bold_format)
            else:
                worksheet.write(row_num + 1, 0, value)

    Sheet4.to_html(output_file_html, index=False)

    print(True)

if len(sys.argv) < 2:
    # нет файлов
    print(False)
else:
    # sys.argv[1] - загрузим первый файл, если их несколько то нужно загружать их в цикле for arg in sys.argv[1:]: 
    excel_file = utils.createEnvPath('SAVED_FILES_PATH', sys.argv[1])
    Sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')

    if 'Назв. вида поставки' in Sheet1.columns:
        if Sheet1['Назв. вида поставки'].iloc[0] == 'Отпуск груза':
            run_script(sys.argv[1])
        else:
            print(False)
    else:
        print(False)
