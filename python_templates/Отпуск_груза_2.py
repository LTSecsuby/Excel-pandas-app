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

def run_script(table_data, contractor_data, output_file_excel, output_file_html):
    Sheet1 = table_data
    Sheet2 = contractor_data

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

    # utils.create_pivot_table_and_get_div_list - создать сводную типа нарушение да нет процент итого и возвращает ее и список дивизионов в итоговой таблице
    # data_table - исходная таблица с данными
    # violation_col - столбц с данными по нарушениями
    # div_col_name - название столбца с дивизионами
    # rp_col_name - название столбца с заводами
    # values_col_name - название столбца значений
    result = utils.create_pivot_table_and_get_div_list(Sheet1, 'Нарушение итог', 'Дивизион', 'Наименование завода', 'Назв. вида поставки')
    Sheet4 = result['table']
    division_list = result['div']

    # Sheet3.columns = [' '.join(col).strip() for col in Sheet3.columns.values]
    # Sheet3 = Sheet3.reset_index()

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

load_obj = utils.load_file_obj(sys.argv[1:])
output_file_excel = load_obj["output_file_excel"]
output_file_html = load_obj["output_file_html"]
files = load_obj["files"]

table_data = None
contractor_data = None
isExistData = False
isExistContractor = False

for file in files:
    for sheet_name, sheet in file.items():
        if 'Назв. вида поставки' in sheet.columns:
            if sheet['Назв. вида поставки'].iloc[0] == 'Отпуск груза':
                isExistData = True
                table_data = sheet
        elif 'Торговый документ подрядчик' in sheet.columns:
            isExistContractor = True
            contractor_data = sheet

if isExistData and isExistContractor:
    run_script(table_data, contractor_data, output_file_excel, output_file_html)
else:
    print(False)
