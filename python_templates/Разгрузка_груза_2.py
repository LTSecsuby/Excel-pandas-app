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

from datetime import datetime, timedelta

from dotenv import load_dotenv
load_dotenv()
pd.options.mode.chained_assignment = None


def run_script(file_name):
    return

if len(sys.argv) < 2:
    # нет файлов
    print(False)
else:
    # sys.argv[1] - загрузим первый файл, если их несколько то нужно загружать их в цикле for arg in sys.argv[1:]:
    res_list = []
    num_list = 1
    for arg in sys.argv[1:]:
        # нулевой не трогаем
        file_name = sys.argv[1]
        if sys.argv[0] != arg:
            excel_file = utils.createEnvPath('SAVED_FILES_PATH', arg)
            # прочитаем самую первую страницу
            current_sheet = pd.read_excel(excel_file, engine='openpyxl')

            if 'Наименование типа документа' in current_sheet.columns:
                current_sheet['Время размещения фотографий/документов'] = pd.to_timedelta(current_sheet['Время размещения фотографий/документов'].astype(str))
                current_sheet['ДатаВремя'] = (pd.to_datetime(current_sheet['Дата размещения фотографий/документов']) + current_sheet['Время размещения фотографий/документов'])
                for index, row in current_sheet.iterrows():
                    key = row['Ключ объекта']
                    if pd.isna(key):
                        continue
                    rp = row['Завод пользователя']
                    if isinstance(rp, float):
                        rp = int(round(rp))
                        rp = str(rp)
                        if len(rp) < 4:
                            rp = '0' + rp
                    time = row['ДатаВремя']
                    object = {'key': key, 'rp': rp, 'time': time}
                    res_list.append(object)

    Sheet1 = pd.DataFrame(columns=['Дата/время размещения фотографий/документов','Номер документа-основания= Ключ объекта', 'Завод пользователя'])

    for i, object in enumerate(res_list):
        Sheet1.at[i, 'Дата/время размещения фотографий/документов'] = object['time']
        Sheet1.at[i, 'Номер документа-основания= Ключ объекта'] = object['key']
        Sheet1.at[i, 'Завод пользователя'] = object['rp']

    Sheet1['Номер документа-основания= Ключ объекта'] = Sheet1['Номер документа-основания= Ключ объекта'].astype(float)

    Sheet1 = Sheet1.sort_values(by='Дата/время размещения фотографий/документов')

    # divisions = pd.DataFrame()
    json_file = utils.createEnvPath('SAVED_SETTINGS_FILES_PATH', 'дивизионы.json')
    values_to_add_rp = []
    values_to_add_rp_num = []
    with open(json_file, encoding="utf-8") as f:
        load_json = json.load(f)
        values_to_add_rp_num = load_json['table'][1]['values']
        values_to_add_rp = load_json['table'][2]['values']

        items_num_rp = dict(zip(values_to_add_rp_num, values_to_add_rp))

        Sheet1['Наименован завода польз'] = Sheet1.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Завод пользователя', items_list=items_num_rp)

    Sheet1['ввв'] = Sheet1['Номер документа-основания= Ключ объекта'].astype(str).str.slice(0, 10) + Sheet1['Наименован завода польз'].astype(str)

    unknowns_div = Sheet1.loc[Sheet1['Наименован завода польз'] == 'нет значения', 'Завод пользователя'].tolist()
    if len(unknowns_div) > 0:
        error_json = utils.createEnvPath('SAVED_ERRPR_PATH', 'unknowns_division')
        output_file_json = os.path.splitext(error_json)[0] + '.json'
        error = pd.DataFrame()
        error['error'] = unknowns_div
        with open(output_file_json, 'w', encoding='utf-8') as file:
            error.to_json(output_file_json, force_ascii=False)
        print('unknowns_division')
    else:
        output_file = utils.createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
        output_file_html = os.path.splitext(output_file)[0] + '.html'
        Sheet1.to_excel(output_file, index=False)

        # тут можно накинуть стилей в уже сохраненные листы файла, сохранить нужные листы и тд (примеры ниже)
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            book = writer.book
            num_format = book.add_format({'num_format': '0'})
            wrap_format = book.add_format({'bold': True})
            wrap_format.set_text_wrap()

            Sheet1.to_excel(writer, sheet_name="zimg", index=False)
            worksheet = writer.sheets["zimg"]

            worksheet.set_column('A:A', 30, num_format)
            worksheet.set_column('B:B', 30, num_format)
            worksheet.set_column('C:C', 18, num_format)
            worksheet.set_column('D:D', 30, num_format)
            worksheet.set_column('E:E', 30, num_format)
            worksheet.write(0, 0, "Дата/время размещения фотографий/документов", wrap_format)
            worksheet.write(0, 1, "Номер документа-основания= Ключ объекта", wrap_format)
            worksheet.write(0, 2, "Завод пользователя", wrap_format)
            worksheet.write(0, 3, "Наименован завода польз", wrap_format)
            worksheet.write(0, 4, "ввв", wrap_format)

        Sheet1.to_html(output_file_html, index=False)        
        print(True)