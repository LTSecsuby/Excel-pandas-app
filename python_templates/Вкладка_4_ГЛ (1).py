# чтобы работать с настройками (не удалять)
import json
# библиотека pandas (не удалять)
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import numpy as np
from numpy import sqrt, abs, round
# для работы с аргументами передаваемыми в скрипт (не удалять)
import sys
# для правильной работы в продакшен версии нужно подтянуть переменные среды (не удалять)
from dotenv import load_dotenv
load_dotenv()
# для работы import utils нужно подтянуть пути проекта (не удалять)
import os
PROJECT_ROOT = os.path.abspath(os.path.join(
                  os.path.dirname(__file__), 
                  os.pardir))
sys.path.append(PROJECT_ROOT)
# в utils будут общие функции, которые можно будет использовать для облегчения создания скриптов (будет пополняться)
import utils


def run_script(data, output_file_excel, output_file_html):

    values_to_add_rp = utils.load_settings_table_column_values('дивизионы.json', 'РП')
    values_to_add_div = utils.load_settings_table_column_values('дивизионы.json', 'Дивизион')
    items_div = dict(zip(values_to_add_rp, values_to_add_div))

    values_to_drop = utils.load_settings_table_column_values('удалить.json', 'Доп склад')
    data = data[~data['Наименование завода'].isin(values_to_drop)]

    values_to_add_Rp = utils.load_settings_table_column_values('Часовой пояс.json', 'Наименование завода')
    values_to_add_hp = utils.load_settings_table_column_values('Часовой пояс.json', 'Часовой_пояс')
    values_to_add_zn = utils.load_settings_table_column_values('Часовой пояс.json', 'Значение')

    items_times = dict(zip(values_to_add_Rp, values_to_add_hp))
    items_times_value = dict(zip(values_to_add_Rp, values_to_add_zn))

    data['Дивизион'] = data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наименование завода', items_list=items_div, default_value='Пустой дивизион')
    data['Часовой_пояс'] = data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наименование завода', items_list=items_times, default_value='Пустой Часовой_пояс')
    data['Значение'] = data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наименование завода', items_list=items_times_value, default_value='Пустое Значение')

    error_div = utils.check_null_pointer_in_table_value(data, 'Дивизион', 'Наименование завода', 'Пустой дивизион')
    if error_div:
        print('unknowns_division')
        return
    error_times = utils.check_null_pointer_in_table_value(data, 'Часовой_пояс', 'Наименование завода', 'Пустой Часовой_пояс')
    if error_times:
        print('unknowns_division')
        return
    error_times_value = utils.check_null_pointer_in_table_value(data, 'Значение', 'Наименование завода', 'Пустое Значение')
    if error_times_value:
        print('unknowns_division')
        return        

    data['Часовой_пояс'] = data['Часовой_пояс'].astype(int)

    data = data.drop_duplicates(subset='№ ЭР')      
    data['time'] = data['Дата создания'].astype(str) + " " + data['Время'].astype(str)
    data['time'] = pd.to_datetime(data['time'])
    def time_lp(data):
        if data['Значение'] == 'прибавить':
            return data['time'] + relativedelta(hours=data['Часовой_пояс'])
        else:
            return data['time'] - relativedelta(hours=data['Часовой_пояс'])
    data['дата и время создания с учетом час пояса'] = data.apply(time_lp, axis=1)
    data['Дата создания с час поясом'] = pd.to_datetime(data['дата и время создания с учетом час пояса']).dt.date
    data['Время создания с час поясом'] = pd.to_datetime(data['дата и время создания с учетом час пояса']).dt.time
    data['День недели'] = pd.to_datetime(data['Дата создания с час поясом']).dt.dayofweek
    data['Время создания с час поясом'] = pd.to_datetime(data['дата и время создания с учетом час пояса']).dt.hour
    def time_lp_1(data):
        if (data['День недели'] == 4 and data['Время создания с час поясом'] > 12):
            return data['дата и время создания с учетом час пояса'] + relativedelta(days=+3)
        elif data['День недели'] == 5:
            return data['дата и время создания с учетом час пояса'] + relativedelta(days=+2)
        elif data['День недели'] == 6:
            return data['дата и время создания с учетом час пояса'] + relativedelta(days=+1)
        else:
            return data['Дата создания с час поясом']
    data['Дата создания с учетом выходных'] = data.apply(time_lp_1, axis=1)
    data['Дата создания с учетом выходных'] = pd.to_datetime(data['Дата создания с учетом выходных']).dt.date
    data['Плановая дата забора'] = data['Плановая дата забора'].astype(str)
    def time_lp_2(data):
        if data['Плановая дата забора'] == 'NaT':
            return 'убрать'
        else:
            return ''
    data['убрать нет план даты'] = data.apply(time_lp_2, axis=1)
    data['Дата статуса  "Груз принят"'] = data['Дата статуса  "Груз принят"'].astype(str)
    data['Плановая дата забора'] = pd.to_datetime(data['Плановая дата забора'])
    def time_lp_3(data):
        if (data['Дата статуса  "Груз принят"'] == 'NaT'):
            if (data['Плановая дата забора'] > datetime.now()):
                return 'убрать'
        else:
            return ''
    data['убрать план дата больше даты создания отчета'] = data.apply(time_lp_3, axis=1)
    data['Дата статуса  "Груз принят"'] = pd.to_datetime(data['Дата статуса  "Груз принят"'])
    data['Срок забора без выходных'] = (data['Дата статуса  "Груз принят"'] - data['Дата создания'])
    data['Срок забора без выходных'] = data['Срок забора без выходных'].dt.days
    data['Срок забора с выходными'] = (data['Дата статуса  "Груз принят"'] - pd.to_datetime(data['Дата создания с учетом выходных']))
    data['Срок забора с выходными'] = data['Срок забора с выходными'].dt.days
    data['Срок забора с выходными'] = data['Срок забора с выходными'].replace(-3, 0)
    data['Срок забора с выходными'] = data['Срок забора с выходными'].replace(-2, 1)
    data['Срок забора с выходными'] = data['Срок забора с выходными'].replace(-1, 2)
    data['меньшее'] = np.minimum(data['Срок забора без выходных'], data['Срок забора с выходными'])
    data['меньшее'] = data['меньшее'].fillna(-45000)
    data['убрать нет план даты'] = data['убрать нет план даты'].astype(str)
    data['убрать план дата больше даты создания отчета'] = data['убрать план дата больше даты создания отчета'].astype(str)
    def minimum(data):
        if data['меньшее'] == 0:
            return 'Забор день в день'
        elif data['меньшее'] == 1:
            return 'Забор на следующий день'
        elif data['меньшее'] == 2:
            return 'Забор через 2 дня'
        elif data['меньшее'] == 3:
            return 'Забор через 3 дня'
        elif data['меньшее'] == 4:
            return 'Забор через 4-5 дней'
        elif data['меньшее'] == 5:
            return 'Забор через 4-5 дней'
        elif data['меньшее'] >= 6:
            return 'Забор через 6 дней и более'
        else:
            return 'Забор груза не состоялся (возможно отсутствует услуга забора)'
    data['периоды'] = data.apply(minimum, axis=1)
    data = data_end[data_end['убрать нет план даты'] != 'убрать']
    data = data_end[data_end['убрать план дата больше даты создания отчета'] != 'убрать']
    table = pd.pivot_table(data_end, values='№ ЭР', index='Дивизион', aggfunc='count', columns='периоды', margins_name='Итог количество', margins=True, dropna=True)
    table = table.reset_index()
    table = table.rename(columns={'Дивизион' : 'Див_Завод'})
    table['% от количества_1'] = (table['Забор груза не состоялся (возможно отсутствует услуга забора)'] / table['Итог количество']*100).round(2)
    table['% от количества_2'] = (table['Забор день в день'] / table['Итог количество']*100).round(2)
    table['% от количества_3'] = (table['Забор на следующий день'] / table['Итог количество']*100).round(2)
    table['% от количества_4'] = (table['Забор через 2 дня'] / table['Итог количество']*100).round(2)
    table['% от количества_5'] = (table['Забор через 3 дня'] / table['Итог количество']*100).round(2)
    table['% от количества_6'] = (table['Забор через 4-5 дней'] / table['Итог количество']*100).round(2)
    table['% от количества_7'] = (table['Забор через 6 дней и более'] / table['Итог количество']*100).round(2)
    table['% от количества_sum'] = (table['Итог количество'] / table['Итог количество']*100).round(2)
    table_2 = pd.pivot_table(data_end, values='№ ЭР', index=['Наименование завода отправителя'], aggfunc='count', columns='периоды', margins_name='Итог количество', margins=True, dropna=True)
    table_2.reset_index()
    table_2['% от количества_1'] = (table_2['Забор груза не состоялся (возможно отсутствует услуга забора)'] / table_2['Итог количество']*100).round(2)
    table_2['% от количества_2'] = (table_2['Забор день в день'] / table_2['Итог количество']*100).round(2)
    table_2['% от количества_3'] = (table_2['Забор на следующий день'] / table_2['Итог количество']*100).round(2)
    table_2['% от количества_4'] = (table_2['Забор через 2 дня'] / table_2['Итог количество']*100).round(2)
    table_2['% от количества_5'] = (table_2['Забор через 3 дня'] / table_2['Итог количество']*100).round(2)
    table_2['% от количества_6'] = (table_2['Забор через 4-5 дней'] / table_2['Итог количество']*100).round(2)
    table_2['% от количества_7'] = (table_2['Забор через 6 дней и более'] / table_2['Итог количество']*100).round(2)
    table_2['% от количества_sum'] = (table_2['Итог количество'] / table_2['Итог количество']*100).round(2)
    table_2 = table_2[['Див_Завод', 'Забор груза не состоялся (возможно отсутствует услуга забора)', '% от количества_1',
            'Забор день в день', '% от количества_2', 'Забор на следующий день', '% от количества_3',
            'Забор через 2 дня', '% от количества_4', 'Забор через 3 дня', '% от количества_5', 'Забор через 4-5 дней',
            '% от количества_6', 'Забор через 6 дней и более', '% от количества_7', 'Итог количество', '% от количества_sum']]
    table_2 = table_2.reset_index()
    table_2 = table_2.rename(columns={'Наименование завода отправителя' : 'Див_Завод'})
    frames = [table, table_2]
    result = pd.concat(frames)
    result = result.reset_index()
    result = result[['Див_Завод', 'Забор груза не состоялся (возможно отсутствует услуга забора)', '% от количества_1',
            'Забор день в день', '% от количества_2', 'Забор на следующий день', '% от количества_3',
            'Забор через 2 дня', '% от количества_4', 'Забор через 3 дня', '% от количества_5', 'Забор через 4-5 дней',
            '% от количества_6', 'Забор через 6 дней и более', '% от количества_7', 'Итог количество', '% от количества_sum']]
    
    result_table = pd.DataFrame()        

    # !!!! ТУТ идет сохранение полученного результата, чтобы отправить обратно в приложение и скачать
    # тут можно накинуть стилей в уже сохраненные листы файла, сохранить нужные листы и тд (примеры ниже)
    with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
        # пересохраняем нужные листы
        data.to_excel(writer, sheet_name="Общая", index=False)       
        # table_itog.to_excel(writer, sheet_name="Итоговая", index=False)          

    # сохраним html для вывода таблицы в приложении в браузере
    data.to_html(output_file_html, index=False)
    # print(True) - заканчивает выполнение скрипта и выходит в приложение
    print(True)


load_obj = utils.load_file_obj(sys.argv[1:])
output_file_excel = load_obj["output_file_excel"]
output_file_html = load_obj["output_file_html"]
files = load_obj["files"]

data = None
isExistData = False

# сюда попадаем если файлов больше одного, то есть их несколько, значит нужно считывать их в цикле for file in files:
# в file названия загруженных файлов (выгрузок)
for file in files:
    for sheet_name, sheet in file.items():
        if 'Дата создания' in sheet.columns and 'Время' in sheet.columns and 'Плановая дата забора' in sheet.columns:
            data = sheet
            isExistData = True

if isExistData:
    run_script(data, output_file_excel, output_file_html)
else:
    print(False)
