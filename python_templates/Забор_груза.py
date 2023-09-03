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

        # utils.create_pivot_table_and_get_div_list - создать сводную типа нарушение да нет процент итого и возвращает ее и список дивизионов в итоговой таблице
        # data_table - исходная таблица с данными
        # violation_col - столбц с данными по нарушениями
        # div_col_name - название столбца с дивизионами
        # rp_col_name - название столбца с заводами
        # values_col_name - название столбца значений
        result = utils.create_pivot_table_and_get_div_list(Sheet1, 'Нарушение', 'Дивизион', 'Наименование завода', 'Назв. вида поставки')
        Sheet3 = result['table']
        division_list = result['div']

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
