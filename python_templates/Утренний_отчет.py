# чтобы работать с настройками (не удалять)
import json
# библиотека pandas (не удалять)
import pandas as pd
# для работы с аргументами передаваемыми в скрипт (не удалять)
from datetime import datetime
import datetime as dt
import sys
import numpy as np
# для правильной работы в продакшен версии нужно подтянуть переменные среды (не удалять)
from dotenv import load_dotenv
load_dotenv()
# для работы import utils нужно подтянуть пути проекта (не удалять)
import os
PROJECT_ROOT = os.path.abspath(os.path.join(
                  os.path.dirname(__file__), 
                  os.pardir)
)
sys.path.append(PROJECT_ROOT)
# в utils будут общие функции, которые можно будет использовать для облегчения создания скриптов (будет пополняться)
import utils


def run_script(data, output_file_excel, output_file_html):

    values_to_add_rp = utils.load_settings_table_column_values('дивизионы.json', 'РП')
    values_to_add_div = utils.load_settings_table_column_values('дивизионы.json', 'Дивизион')
    items_div = dict(zip(values_to_add_rp, values_to_add_div))

    data['Дивизион'] = data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наименование завода', items_list=items_div, default_value='Пустой дивизион')    
    data['Завод'] = data['Завод'].apply(lambda x: str(x).zfill(4))
    data['Партия'] = data['Партия'].apply(lambda x: str(x).zfill(10))

    error = utils.check_null_pointer_in_table_value(data, 'Дивизион', 'Наименование завода', 'Пустой дивизион')
    if error:
        print('unknowns_division')
    else:
        data_filter = data['План дата прибытия'].min()
        data['Груз к выдаче'] = data['Груз к выдаче'].astype(str)
        data['ДатаСписания'] = data['ДатаСписания'].astype(str)
        data['ДатаПогрузки'] = data['ДатаПогрузки'].astype(str)                

        v_1 = data.loc[(data['План дата прибытия'] >= data_filter) & (data['Груз к выдаче'] == 'NaT') & (data['ДатаСписания'] == 'NaT') & (data['ДатаПогрузки'] == 'NaT')]


        '''КОД ДЛЯ ВРЕМЕНИ'''
        # Преобразование столбцов с датой и временем в соответствующие типы
        v_1.loc[:, 'ДатаПоступл'] = pd.to_datetime(v_1['ДатаПоступл'], format='%d.%m.%Y')
        v_1.loc[:, 'Время проводки'] = pd.to_datetime(v_1['Время проводки'], errors='coerce').dt.time


        # Получение текущей даты
        current_date = datetime.now().date()

        # Фильтрация данных
        min_date = v_1['ДатаПоступл'].min()
        yesterday_filtered = v_1[(v_1['ДатаПоступл'] == min_date) & (v_1['Время проводки'] <= pd.to_datetime('08:30:00').time())]

        # Удаление отфильтрованных строк из исходного DataFrame
        v_1 = v_1.drop(yesterday_filtered.index)
        today_filtered = v_1[(v_1['ДатаПоступл'] == current_date) & (v_1['Время проводки'] > pd.to_datetime('08:30:00').time())]
        v_1 = v_1.drop(today_filtered.index)
        '''КОД ДЛЯ ВРЕМЕНИ'''
        



        v_1 = v_1.sort_values(by='ДатаПоступл', ascending=False)
        v_1 = v_1.drop_duplicates(['Партия', 'Завод'])

        v_1 = v_1[v_1["Наименование завода"].str.contains("Ритейл")== False]
        v_1 = v_1[v_1["Наименование завода"].str.contains("Москва-Щербинка")== False]
        v_1['Kit_otpr'] = v_1['Наим отправителя'].str[:4]
        v_1['Kit_poluch'] = v_1['Наим получателя'].str[:4]
        v_1 = v_1[v_1['Kit_otpr'] != 'КИТ ']
        v_1 = v_1[v_1['Kit_poluch'] != 'КИТ ']
        v_1 = v_1[v_1["Наим отправителя"].str.contains("GTD")== False]
        v_1 = v_1[v_1["Наим отправителя"].str.contains("НЕВОСТРЕБ")== False]
        v_1 = v_1[v_1["Наим отправителя"].str.contains("Невостреб")== False]
        v_1 = v_1[v_1["Наим отправителя"] != 'Временный тех. дебитор']
        v_1 = v_1[v_1["Наим получателя"].str.contains("GTD")== False]
        v_1 = v_1[v_1["Наим получателя"].str.contains("НЕВОСТРЕБ")== False]
        v_1 = v_1[v_1["Наим получателя"].str.contains("Невостреб")== False]
        v_1 = v_1[v_1["Наим получателя"] != 'Временный тех. дебитор']
        v_1 = v_1[v_1['Наименование завода'] != v_1['Наим завод куда']]
        v_1 = v_1[v_1['Наим завод куда'] != v_1['Наим завод откуда']]

        nedostachy = v_1[((v_1['Тек запас на складе'] == 0) & (v_1['Услуга'] == 'GR'))]
        nd = nedostachy[['Партия', 'Завод']]

        v_drobl = data[data['ЧастОтпр'] == 'X']

        v_drobl = v_drobl.sort_values(by='ДатаПоступл', ascending=False)
        v_drobl = v_drobl.drop_duplicates(['Партия', 'Завод'])

        v_drobl = v_drobl[v_drobl["Наименование завода"].str.contains("Ритейл")== False]
        v_drobl['Kit_otpr'] = v_drobl['Наим отправителя'].str[:4]
        v_drobl['Kit_poluch'] = v_drobl['Наим получателя'].str[:4]
        v_drobl = v_drobl[v_drobl['Kit_otpr'] != 'КИТ ']
        v_drobl = v_drobl[v_drobl['Kit_poluch'] != 'КИТ ']
        v_drobl = v_drobl[v_drobl["Наим отправителя"].str.contains("GTD")== False]
        v_drobl = v_drobl[v_drobl["Наим отправителя"].str.contains("НЕВОСТРЕБ")== False]
        v_drobl = v_drobl[v_drobl["Наим отправителя"].str.contains("Невостреб")== False]
        v_drobl = v_drobl[v_drobl["Наим отправителя"] != 'Временный тех. дебитор']
        v_drobl = v_drobl[v_drobl["Наим получателя"].str.contains("GTD")== False]
        v_drobl = v_drobl[v_drobl["Наим получателя"].str.contains("НЕВОСТРЕБ")== False]
        v_drobl = v_drobl[v_drobl["Наим получателя"].str.contains("Невостреб")== False]
        v_drobl = v_drobl[v_drobl["Наим получателя"] != 'Временный тех. дебитор']
        v_drobl = v_drobl[v_drobl['Наименование завода'] != v_drobl['Наим завод куда']]
        v_drobl = v_drobl[v_drobl['Наим завод куда'] != v_drobl['Наим завод откуда']]



        '''КОД ДЛЯ ВРЕМЕНИ'''
        # Преобразование столбцов с датой и временем в соответствующие типы
        v_1.loc[:, 'ДатаПоступл'] = pd.to_datetime(v_1['ДатаПоступл'], format='%d.%m.%Y')
        v_1.loc[:, 'Время проводки'] = pd.to_datetime(v_1['Время проводки'], errors='coerce').dt.time


        # Получение текущей даты
        current_date = datetime.now().date()

        # Фильтрация данных
        min_date = data['ДатаПоступл'].min()
        
        yesterday_filtered = data[(data['ДатаПоступл'] == min_date) & (data['Время проводки'] < pd.to_datetime('08:30:00').time())]
        data = data.drop(yesterday_filtered.index)

        today_filtered = data[(data['ДатаПоступл'] == current_date) & (data['Время проводки'] > pd.to_datetime('08:30:00').time())]
        data = data.drop(today_filtered.index)
        '''КОД ДЛЯ ВРЕМЕНИ'''


        data = data.sort_values('ДатаПоступл', ascending=False)
        data = data.drop_duplicates(['Партия', 'Завод'])

        data = data[data["Наименование завода"].str.contains("Ритейл")== False]
        data = data[data["Наименование завода"].str.contains("Москва-Щербинка")== False]
        data['Kit_otpr'] = data['Наим отправителя'].str[:4]
        data['Kit_poluch'] = data['Наим получателя'].str[:4]
        data = data[data['Kit_otpr'] != 'КИТ ']
        data = data[data['Kit_poluch'] != 'КИТ ']
        data = data[data["Наим отправителя"].str.contains("GTD")== False]
        data = data[data["Наим отправителя"].str.contains("НЕВОСТРЕБ")== False]
        data = data[data["Наим отправителя"].str.contains("Невостреб")== False]
        data = data[data["Наим отправителя"] != 'Временный тех. дебитор']
        data = data[data["Наим получателя"].str.contains("GTD")== False]
        data = data[data["Наим получателя"].str.contains("НЕВОСТРЕБ")== False]
        data = data[data["Наим получателя"].str.contains("Невостреб")== False]
        data = data[data["Наим получателя"] != 'Временный тех. дебитор']
        data = data[data['Наим завод куда'] != data['Наим завод откуда']]
        data = data[data['Дивизион'] != "Ритейл"]
        data = data[data['Дивизион'] != "Аэропорт"]
        data = data[data['Дивизион'] != "Удалить"]

        ezhedn_otchet = pd.pivot_table(data, values='ЭР', index='Дивизион', aggfunc='count')
        ezhedn_otchet = ezhedn_otchet.reset_index()
        msk_er = data[data["Наименование завода"] == 'Москва'].agg({'ЭР' : 'count'})

        result_table = pd.DataFrame()               

        # !!!! ТУТ идет сохранение полученного результата, чтобы отправить обратно в приложение и скачать
        # тут можно накинуть стилей в уже сохраненные листы файла, сохранить нужные листы и тд (примеры ниже)
        with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:

            v_1['ДатаПоступл'] = v_1['ДатаПоступл'].dt.strftime('%d.%m.%Y')
            v_drobl['ДатаПоступл'] = v_drobl['ДатаПоступл'].dt.strftime('%d.%m.%Y')
            data['ДатаПоступл'] = data['ДатаПоступл'].dt.strftime('%d.%m.%Y')


            v_1['План дата прибытия'] = v_1['План дата прибытия'].dt.strftime('%d.%m.%Y')
            v_drobl['План дата прибытия'] = v_drobl['План дата прибытия'].dt.strftime('%d.%m.%Y')
            data['План дата прибытия'] = data['План дата прибытия'].dt.strftime('%d.%m.%Y')

            
            # пересохраняем нужные листы
            v_1.to_excel(writer, sheet_name="Выгрузка_1", index=False)
            nd.to_excel(writer, sheet_name="Недостачи", index=False)
            v_drobl.to_excel(writer, sheet_name="Дробление", index=False)
            ezhedn_otchet.to_excel(writer, sheet_name="Ежедневный_отчет_общая", index=False)
            msk_er.to_excel(writer, sheet_name="Москва_количество_ЭР", index=False) 
            data.to_excel(writer, sheet_name="Общая_выгрузка", index=False)



            # Получите объекты 'workbook' и 'worksheet' из ExcelWriter для первого листа
            workbook = writer.book
            worksheet = writer.sheets['Выгрузка_1']
                        
            # Получите объекты 'workbook' и 'worksheet' из ExcelWriter для первого листа
            workbook_2 = writer.book
            worksheet_2 = writer.sheets['Дробление']


                                    
            # Получите объекты 'workbook' и 'worksheet' из ExcelWriter для первого листа
            workbook_3 = writer.book
            worksheet_3 = writer.sheets['Общая_выгрузка']


            # Установите автоширину для каждого столбца, основываясь на его содержании для первого листа
            for col_num, col in enumerate(v_1.columns):
                max_len = max(v_1[col].astype(str).apply(len).max() + 1, len(str(col)))
                worksheet.set_column(col_num, col_num, max_len) 


            # Установите автоширину для каждого столбца, основываясь на его содержании для первого листа
            for col_num, col in enumerate(v_drobl.columns):
                max_len = max(v_drobl[col].astype(str).apply(len).max() + 1, len(str(col)))
                worksheet_2.set_column(col_num, col_num, max_len)      

            # Установите автоширину для каждого столбца, основываясь на его содержании для первого листа
            for col_num, col in enumerate(data.columns):
                max_len = max(data[col].astype(str).apply(len).max() + 1, len(str(col)))
                worksheet_3.set_column(col_num, col_num, max_len)   

            # сохраним html для вывода таблицы в приложении в браузере
            result_table.to_html(output_file_html, index=False)
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
        if 'План дата прибытия' in sheet.columns and 'ЧастОтпр' in sheet.columns and 'Тек запас на складе' in sheet.columns:
            data = sheet
            isExistData = True

if isExistData:
    run_script(data, output_file_excel, output_file_html)
else:
    print(False)
