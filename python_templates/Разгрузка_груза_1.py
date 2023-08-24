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

def calculate_duration(value):
    if value == 'Пусто':
        return 'Пусто'
    else:
        timedelta = pd.to_timedelta(value) - pd.to_timedelta('0h')
        return timedelta / pd.Timedelta(days=1)

# функция для форматирования timedelta в виде "0:00:00"
def format_timedelta(timedelta):
    total_seconds = timedelta.total_seconds()
    if pd.isnull(timedelta):
        return "Пусто"
    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    seconds = int(total_seconds % 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

# функция для форматирования timedelta в seconds
def format_float(decimal_time):
    hours = math.ceil(decimal_time)
    minutes = int((decimal_time % 1) * 60)
    seconds = int((decimal_time % 1) * 60 % 1 * 60)
    return hours * 3600 + minutes * 60 + seconds

# Определение функции для вычисления значений в новой колонке
def calculate_normative_time(row):
    if row['Категория склада'] == '1':
        return format_float(0.16666667)
    elif row['Категория склада'] == '2':
        return format_float(0.16666667)
    elif row['Категория склада'] == '3':
        return format_float(0.25)
    elif row['Категория склада'] == '4':
        return format_float(0.25)
    elif row['Категория склада'] == '5':
        return format_float(0.33333333)
    elif row['Категория склада'] == '6':
        return format_float(0.33333333)
    elif row['Категория склада'] == '1Д':
        return format_float(0.33333333)
    elif row['Категория склада'] == '2Д':
        return format_float(0.33333333)
    elif row['Категория склада'] == '3Д':
        return format_float(0.33333333)
    elif row['Категория склада'] == '4Д':
        return format_float(0.33333333)
    else:
        return "Пусто"

def calculate_normative_time_format(row):
    if row['Категория склада'] == '1':
        return '04:00:00'
    elif row['Категория склада'] == '2':
        return '04:00:00'
    elif row['Категория склада'] == '3':
        return '06:00:00'
    elif row['Категория склада'] == '4':
        return '06:00:00'
    elif row['Категория склада'] == '5':
        return '08:00:00'
    elif row['Категория склада'] == '6':
        return '08:00:00'
    elif row['Категория склада'] == '1Д':
        return '08:00:00'
    elif row['Категория склада'] == '2Д':
        return '08:00:00'
    elif row['Категория склада'] == '3Д':
        return '08:00:00'
    elif row['Категория склада'] == '4Д':
        return '08:00:00'
    else:
        return "Пусто"

def calculate_normative_time_period(value):
    if value == 'Пусто':
        return "4.Не завершена"
    elif value < 0.16668:
        return '1.0-4 часа'
    elif value == 0.16668:
        return '1.0-4 часа'
    elif value < 0.5:
        return '2.4-12 часов'
    elif value == 0.5:
        return '2.4-12 часов'
    elif value > 0.5:
        return '3.Более 12 часов'
    else:
        return "4.Не завершена"

def run_script(file_name):
    excel_file = utils.createEnvPath('SAVED_FILES_PATH', file_name)

    Sheet1 = pd.read_excel(excel_file, sheet_name='Лист1', engine='openpyxl')

    Sheet1 = Sheet1.loc[Sheet1['СрМакс%Заг'] != 0]

    # тут сделать сначало все столбцы из отчета потом уже ниже

    # Sheet1['Эт:ПлнВрПр'] = pd.to_datetime(Sheet1['Эт:ПлнВрПр'])
    # Sheet1['Эт:ФктВрПр'] = pd.to_datetime(Sheet1['Эт:ФктВрПр'])

    # 1)	если  ЭтФактДатП1 и Эт:ФактДатР -пусто( машина не заезжала в рп, исключить из отчета)-удалить строки
    Sheet1 = Sheet1.dropna(subset=['Эт:ФктДатП_1', 'Эт:ФктДатР'], how='all')

    Sheet1['Эт:ПлнВрПр'] = pd.to_timedelta(Sheet1['Эт:ПлнВрПр'].astype(str))
    Sheet1['Эт:ФктВрПр'] = pd.to_timedelta(Sheet1['Эт:ФктВрПр'].astype(str))

    Sheet1['Эт:ФактДатаВремяПр'] = (pd.to_datetime(Sheet1['Эт:ФктДатП_1']) + Sheet1['Эт:ФктВрПр'])
    Sheet1['Эт:ПланДатаВремяПр'] = (pd.to_datetime(Sheet1['Эт:ПлнДатП']) + Sheet1['Эт:ПлнВрПр'])

    Sheet1['Очередь на выгрузку= простой на выгрузку'] = Sheet1['Эт:ФактДатаВремяПр'] - Sheet1['Эт:ПланДатаВремяПр']

    Sheet1['Очередь на выгрузку= простой на выгрузку значение'] = Sheet1['Очередь на выгрузку= простой на выгрузку'].map(calculate_duration)

    Sheet1['Очередь на выгрузку= простой на выгрузку'] = Sheet1['Очередь на выгрузку= простой на выгрузку'].apply(format_timedelta)

    # Sheet1['Очередь на выгрузку= простой на выгрузку'] = Sheet1['Очередь на выгрузку= простой на выгрузку'].apply(lambda x: x.total_seconds() / 3600)

    # Sheet1['Очередь на выгрузку= простой на выгрузку'] = Sheet1['Очередь на выгрузку= простой на выгрузку'].apply(lambda x: str(x).split()[-1])

    Sheet1 = Sheet1[~Sheet1['Узел отправки'].str.contains('Шенкер')]
    Sheet1 = Sheet1[~Sheet1['Узел отправки'].str.contains('Нарьян-Мар')]
    Sheet1 = Sheet1[~Sheet1['Узел отправки'].str.contains('Москва-Щербинка')]

    # divisions = pd.DataFrame()
    json_file = utils.createEnvPath('SAVED_SETTINGS_FILES_PATH', 'дивизионы.json')
    values_to_add_rp = []
    values_to_add_div = []
    values_to_add_stock = []
    values_to_add_rp_num = []
    with open(json_file, encoding="utf-8") as f:
        load_json = json.load(f)    
        values_to_add_stock = load_json['table'][0]['values']
        values_to_add_rp_num = load_json['table'][1]['values']
        values_to_add_rp = load_json['table'][2]['values']
        values_to_add_div = load_json['table'][3]['values']

        items_num_rp = dict(zip(values_to_add_rp_num, values_to_add_rp))
        items_div = dict(zip(values_to_add_rp, values_to_add_div))
        items_stock = dict(zip(values_to_add_rp, values_to_add_stock))

        Sheet1['РП разгрузки ТС'] = Sheet1.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Назн.:Звд', items_list=items_num_rp, default_value='Пустой завод')

        Sheet1['див'] = Sheet1.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='РП разгрузки ТС', items_list=items_div, default_value='Пустой див')

    # проверка есть ли все дивизионы по колонке 'Назн.:Звд', check_value_in_list_and_set_value устанавливает 'нет значения' в 'див' если не нашел
    unknowns = []
    for index, row in Sheet1.iterrows():
        if row['див'] == 'нет значения':
            unknowns.append(row['Назн.:Звд'] + ', ')

    if len(unknowns) > 0:
        error_json = utils.createEnvPath('SAVED_ERRPR_PATH', 'unknowns_division')
        output_file_json = os.path.splitext(error_json)[0] + '.json'
        error = pd.DataFrame()
        error['error'] = unknowns
        with open(output_file_json, 'w', encoding='utf-8') as file:
            error.to_json(output_file_json, force_ascii=False)
        print('unknowns_division')
    else:
        Sheet1['Категория склада'] = Sheet1.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='РП разгрузки ТС', items_list=items_stock, default_value='Пустая категория')

        Sheet1['Нормативное Время'] = Sheet1.apply(calculate_normative_time_format, axis=1)

        Sheet1['Нормативное Время значение'] = Sheet1['Нормативное Время'].map(calculate_duration)

        # divisions['РП'] = values_to_add_rp
        # divisions['Дивизион'] = values_to_add_div

        # удалить рп и див по списку в инструкции
        Sheet1 = Sheet1.loc[Sheet1['РП разгрузки ТС'] != 'Москва-Запад']
        Sheet1 = Sheet1.loc[Sheet1['РП разгрузки ТС'] != 'Нарьян-Мар']
        Sheet1 = Sheet1.loc[Sheet1['РП разгрузки ТС'] != 'Москва-Щербинка']

        Sheet1 = Sheet1.loc[Sheet1['див'] != 'Ритейл']
        Sheet1 = Sheet1.loc[Sheet1['див'] != 'Аэропорт']

        # 2) если Эт:ФктДатП_1 пусто, но Эт:ФактДатР-есть дата- в столбце Нарушение заменить на  ДА ( машина прибыла на разгрузку, но не поставили дату(случайно или умышленно) 
        cond1 = Sheet1.apply(lambda row: 'да' if row['Очередь на выгрузку= простой на выгрузку значение'] > row['Нормативное Время значение'] else 'нет', axis=1)
        Sheet1['Нарушение'] = Sheet1.apply(lambda row: 'да' if (pd.isnull(row['Эт:ФктДатП_1']) and not pd.isnull(row['Эт:ФктДатР'])) or (cond1[row.name] == 'да') else 'нет', axis=1)

        Sheet1['Эт:ФктВрРа'] = pd.to_timedelta(Sheet1['Эт:ФктВрРа'].astype(str))

        Sheet1['Эт:ФактДатаВремяП'] = (pd.to_datetime(Sheet1['Эт:ФктДатП_1']) + Sheet1['Эт:ФктВрПр'])
        Sheet1['Эт:ФактДатаВремяР'] = (pd.to_datetime(Sheet1['Эт:ФктДатР']) + Sheet1['Эт:ФктВрРа'])

        Sheet1['Продолжительность выгрузки'] = Sheet1['Эт:ФактДатаВремяР'] - Sheet1['Эт:ФактДатаВремяП']

        Sheet1['Продолжительность выгрузки'] = Sheet1['Продолжительность выгрузки'].apply(format_timedelta)

        Sheet1['Продолжительность выгрузки значение'] = Sheet1['Продолжительность выгрузки'].map(calculate_duration)

        Sheet1['Период выгрузки'] = Sheet1['Продолжительность выгрузки значение'].apply(lambda x: calculate_normative_time_period(x))

        Sheet1['МЛ'] = Sheet1['№ транспортировки'].astype(str).str[:10] + Sheet1['РП разгрузки ТС'].astype(str)

        Sheet2 = pd.DataFrame()

        Sheet2['Номер документа-основания= Ключ объекта'] = None

        Sheet2['Завод пользователя'] = None

        output_file = utils.createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
        output_file_html = os.path.splitext(output_file)[0] + '.html'
        Sheet1.to_excel(output_file, index=False)

        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            book = writer.book
            time_format = book.add_format({'num_format': 'hh:mm:ss'})
            wrap_format = book.add_format({'bold': True})
            wrap_format.set_text_wrap()

            Sheet1.to_excel(writer, sheet_name="Лист1", index=False)
            Sheet2.to_excel(writer, sheet_name="zimg", index=False)
            worksheet1 = writer.sheets["Лист1"]

            worksheet1.set_column('AR:AR', 18, time_format)
            worksheet1.set_column('AV:AV', 18, time_format)
            worksheet1.set_column('AZ:AZ', 18, time_format)

            worksheet2 = writer.sheets["zimg"]

            worksheet2.set_column('A:A', 30, wrap_format)
            worksheet2.set_column('B:B', 18, wrap_format)
            worksheet2.write(0, 0, "Номер документа-основания= Ключ объекта", wrap_format)
            worksheet2.write(0, 1, "Завод пользователя", wrap_format)

        Sheet1.to_html(output_file_html, index=False)
        print(True)

if len(sys.argv) < 2:
    # нет файлов
    print(False)
else:
    # sys.argv[1] - загрузим первый файл, если их несколько то нужно загружать их в цикле for arg in sys.argv[1:]: 
    excel_file = utils.createEnvPath('SAVED_FILES_PATH', sys.argv[1])
    Sheet1 = pd.read_excel(excel_file, sheet_name='Лист1', engine='openpyxl')

    if '№ транспортировки' in Sheet1.columns:
        run_script(sys.argv[1])
    else:
        print(False)
