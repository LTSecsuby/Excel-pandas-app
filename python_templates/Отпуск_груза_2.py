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

def run_script(file_name):
    excel_file = createEnvPath('SAVED_FILES_PATH', file_name)
    encoding = 'utf-8'

    Sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')
    Sheet2 = pd.read_excel(excel_file, sheet_name='Sheet2', engine='openpyxl')
    Sheet3 = pd.DataFrame()

    docs = Sheet2['Торговый документ подрядчик'].tolist()

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
    
    Sheet4 = utils.add_total_by_field(Sheet3, 'Дивизион', ['да', 'нет'])

    total_fact = Sheet3['да'].sum()
    total_cost = Sheet3['нет'].sum()
    total_count = total_fact + total_cost
    total_row = pd.DataFrame({'Наименование завода': [''],
                            'Дивизион': ['Общий итог'],
                            'да': [total_fact],
                            'нет': [total_cost],
                            'Общий итог': [total_count]})
    Sheet4 = pd.concat([Sheet4, total_row])

    sum_column = Sheet4['да'] + Sheet4['нет']
    Sheet4['Общий итог'] = sum_column

    percentage_column = (Sheet4['нет'] / Sheet4['Общий итог']) * 100
    Sheet4['Процент %'] = percentage_column

    # Sheet3.columns = [' '.join(col).strip() for col in Sheet3.columns.values]
    # Sheet3 = Sheet3.reset_index()

    output_file_excel = createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
    output_file_html = os.path.splitext(output_file_excel)[0] + '.html'

    with pd.ExcelWriter(output_file_excel) as writer:
        Sheet1.to_excel(writer, sheet_name="Sheet1", index=False)
        Sheet2.to_excel(writer, sheet_name="Sheet2", index=False)
        Sheet4.to_excel(writer, sheet_name="Sheet3", index=False)

    Sheet4.to_html(output_file_html, index=False)

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
