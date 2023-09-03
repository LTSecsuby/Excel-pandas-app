import sys
import os
import json
import pandas as pd
import numpy as np
import math

# для работы import utils нужно подтянуть пути проекта
PROJECT_ROOT = os.path.abspath(os.path.join(
                  os.path.dirname(__file__), 
                  os.pardir)
)
sys.path.append(PROJECT_ROOT)
import utils

from dotenv import load_dotenv
from datetime import datetime, timedelta
load_dotenv()
pd.options.mode.chained_assignment = None


def run_script(file_name):
    excel_file = utils.createEnvPath('SAVED_FILES_PATH', file_name)

    Sheet1 = pd.read_excel(excel_file, sheet_name='Лист1', engine='openpyxl')
    Sheet2 = pd.read_excel(excel_file, sheet_name='zimg', engine='openpyxl')

    Sheet1['акт'] = Sheet1['МЛ'].apply(lambda x: x if x in Sheet2['ввв'].values else 'Пустой акт')

    Sheet1.loc[Sheet1['акт'] != 'Пустой акт', 'Нарушение'] = 'нет'

    output_file = utils.createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
    output_file_html = os.path.splitext(output_file)[0] + '.html'
    Sheet1.to_excel(output_file, index=False)

    # utils.create_pivot_table_and_get_div_list - создать сводную типа нарушение/да/нет/%/Общий итог и возвращает ее и список дивизионов в итоговой таблице
    # data_table - исходная таблица с данными
    # violation_col - столбц с данными по нарушениями
    # div_col_name - название столбца с дивизионами
    # rp_col_name - название столбца с заводами
    # values_col_name - название столбца значений
    result = utils.create_pivot_table_and_get_div_list(Sheet1, 'Нарушение', 'див', 'РП разгрузки ТС', 'Трка/этап')
    Sheet4 = result['table']
    division_list1 = result['div']

    # utils.create_pivot_table_and_get_div_list2 - создать сводную типа 1.0-4 часа/2.4-12 часов/3.Более 12 часов/4.Не завершена/%/Общий итог и возвращает ее и список дивизионов в итоговой таблице
    # data_table - исходная таблица с данными
    # violation_col - столбц с данными по нарушениями
    # div_col_name - название столбца с дивизионами
    # rp_col_name - название столбца с заводами
    # values_col_name - название столбца значений
    result2 = utils.create_pivot_table_and_get_div_list_2(Sheet1, 'Период выгрузки', 'див', 'РП разгрузки ТС', 'Трка/этап')
    Sheet5 = result2['table']
    division_list2 = result['div']

    # сохраняем
    output_file_excel = utils.createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
    output_file_html = os.path.splitext(output_file_excel)[0] + '.html'

    with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
        book = writer.book
        percent_format = book.add_format({'num_format': '0.00%'})
        bold_format = book.add_format({'bold': True})

        Sheet1.to_excel(writer, sheet_name="выгрузка из SAP", index=False)
        Sheet2.to_excel(writer, sheet_name="zimg", index=False)
        Sheet4.to_excel(writer, sheet_name="сводн", index=False)
        Sheet5.to_excel(writer, sheet_name="сводн2", index=False)

        worksheet1 = writer.sheets["сводн"]

        worksheet1.set_column('E:E', 18, percent_format)
        worksheet1.set_column('A:A', 28)
        worksheet1.set_column('B:D', 18)

        column_values1 = Sheet4['РП разгрузки ТС'].values.tolist()

        for row_num, value in enumerate(column_values1):
            if value in division_list1:
                worksheet1.write(row_num + 1, 0, value, bold_format)
            else:
                worksheet1.write(row_num + 1, 0, value)
        
        worksheet2 = writer.sheets["сводн2"]

        worksheet2.set_column('A:A', 28)
        worksheet2.set_column('B:J', 18)

        column_values2 = Sheet5['РП разгрузки ТС'].values.tolist()

        for row_num, value in enumerate(column_values2):
            if value in division_list2:
                worksheet2.write(row_num + 1, 0, value, bold_format)
            else:
                worksheet2.write(row_num + 1, 0, value)

        worksheet2.set_column('C:C', 18, percent_format)
        worksheet2.set_column('E:E', 18, percent_format)
        worksheet2.set_column('G:G', 18, percent_format)
        worksheet2.set_column('I:I', 18, percent_format)

    Sheet5.to_html(output_file_html, index=False)

    print(True)

if len(sys.argv) < 2:
    # нет файлов
    print(False)
else:
    # sys.argv[1] - загрузим первый файл, если их несколько то нужно загружать их в цикле for arg in sys.argv[1:]: 
    excel_file = utils.createEnvPath('SAVED_FILES_PATH', sys.argv[1])
    Sheet1 = pd.read_excel(excel_file, sheet_name='Лист1', engine='openpyxl')
    Sheet2 = pd.read_excel(excel_file, sheet_name='zimg', engine='openpyxl')

    if '№ транспортировки' in Sheet1.columns and 'ввв' in Sheet2.columns:
        run_script(sys.argv[1])
    else:
        print(False)
