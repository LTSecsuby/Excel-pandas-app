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

def run_script(data, output_file_excel, output_file_html):
    # Здесь оставляем data вместо potok_df

    selected_columns = ['Завод', 'ЭР', 'Наим операции', 'ДатаПоступл', 'Партия', 'Услуга', 
                        'Наименование завода', 'Наим завод куда', 'Вес', 'Объем ЭР', 'НаправПВГ']
    
    data = data[selected_columns]

    data['Доп ПВГ']= ''


    data['Партия'] = data['Партия'].astype(str)

    # Забор из города переименовать в Получение от клиента
    data['Наим операции'].replace('Забор из города', 'Получение от клиента', inplace=True)

    # Удаляем дубликаты
    data = data.drop_duplicates(['Завод', 'ЭР'])

    # Нужны только грузы
    data = data[data['Услуга'] == 'GR']


    # Делаем сводную таблицу и сбрасываем индекс, чтобы повторять подписи элементов, нужно найти сумму Веса и Объёма по двум операциям
    pivot_table = pd.pivot_table(data, values=['Вес', 'Объем ЭР'], index=['Наименование завода', 'Наим завод куда', 'НаправПВГ' ,'ДатаПоступл', 'Наим операции'], aggfunc='sum').reset_index()
    

    # Из сводной таблицы делаем два Frame, чтобы потом их объеднить вместе и создать новый DataFrame
    potok_client = pivot_table[pivot_table['Наим операции'] == 'Получение от клиента']

    potok_transit = pivot_table[pivot_table['Наим операции'] == 'Принятие транзитного груза']

    # Объединяем Frame-мы
    potok_client_transit = potok_client.merge(potok_transit, on=['Наименование завода', 'Наим завод куда', 'НаправПВГ','ДатаПоступл'], suffixes=('_клиент', '_транзит'), how='outer')


    # Словарь для замены английских названий дней недели на русские
    day_translation = {
        'Monday': 'пн',
        'Tuesday': 'вт',
        'Wednesday': 'ср',
        'Thursday': 'чт',
        'Friday': 'пт',
        'Saturday': 'сб',
        'Sunday': 'вс'
    }

    # Функция для определения дня недели
    def get_weekday(date_str):
        date = pd.to_datetime(date_str, format='%d.%m.%Y', errors='coerce')
        if not pd.isnull(date):
            english_weekday = date.strftime('%A')
            return day_translation.get(english_weekday, english_weekday)
        return None

    # Создание нового столбца День недели
    potok_client_transit['День недели'] = potok_client_transit['ДатаПоступл'].apply(get_weekday)


    # Сортировка
    potok_client_transit = potok_client_transit.sort_values(['Наименование завода', 'Наим завод куда','НаправПВГ','ДатаПоступл', 'День недели'], ascending=[True, True, True, True, True])


    # Дату поступления переводим в строковый формат
    potok_client_transit['ДатаПоступл'] = potok_client_transit['ДатаПоступл'].dt.strftime('%d.%m.%Y')
    data['ДатаПоступл'] = data['ДатаПоступл'].dt.strftime('%d.%m.%Y')

    # Группируем данные по нужным столбцам и просуммируем 'Вес_клиент' и 'Вес_транзит' для каждого дня недели
    grouped_data = potok_client_transit.groupby(['Наименование завода', 'Наим завод куда', 'НаправПВГ', 'День недели'])[['Вес_клиент', 'Вес_транзит', 'Объем ЭР_клиент', 'Объем ЭР_транзит']].sum().reset_index()

    # Объединяем результаты группировки с исходным DataFrame
    potok_client_transit = potok_client_transit.merge(grouped_data, on=['Наименование завода', 'Наим завод куда', 'День недели'], how='left', suffixes=('', '_итог'))

    # Создаем новый столбец 'Вес итог', сложив значения 'Вес_клиент' и 'Вес_транзит'
    potok_client_transit['Вес итог'] = potok_client_transit['Вес_клиент'] + potok_client_transit['Вес_транзит']

    # Создаем новый столбец 'Вес итог', сложив значения 'Вес_клиент' и 'Вес_транзит'
    potok_client_transit['Объём итог'] = potok_client_transit['Объем ЭР_клиент'] + potok_client_transit['Объем ЭР_транзит']

    # Убираем лишние столбцы
    potok_client_transit = potok_client_transit.drop(['Вес_клиент_итог', 'Вес_транзит_итог', 'НаправПВГ_итог', 'Объем ЭР_клиент_итог', 'Объем ЭР_транзит_итог'], axis=1)


    # Группируем данные по 'Наименование завода', 'Наим завод куда' и 'День недели', а затем находим среднее значение 'Вес итог'
    result = potok_client_transit.groupby(['Наименование завода', 'Наим завод куда', 'НаправПВГ', 'День недели'])[['Вес итог', 'Объём итог']].mean().reset_index()

        # Сортировка данных
    result = result.groupby(['Наименование завода', 'Наим завод куда'], group_keys=False).apply(
        lambda x: x.sort_values('День недели', key=lambda y: pd.Categorical(y, categories=['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс'], ordered=True))    
    )

        # Округлим значения до 1 знака для potok_df
    data[['Вес', 'Объем ЭР']] = data[['Вес', 'Объем ЭР']].round(1)

    # Округлим значения до 1 знака для potok_client_transit
    columns_to_round = ['Вес_клиент', 'Объем ЭР_клиент', 'Вес_транзит', 'Объем ЭР_транзит', 'Вес итог', 'Объём итог']
    potok_client_transit[columns_to_round] = potok_client_transit[columns_to_round].round(1)

    # Округлим значения до 1 знака для result
    result[['Вес итог', 'Объём итог']] = result[['Вес итог', 'Объём итог']].round(1)


    data = data.rename(columns={'Наименование завода':'Наим завода где'})


    with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
        data.to_excel(writer, sheet_name="Сверка сроков доставки", index=False)
        potok_client_transit.to_excel(writer, sheet_name="Вес и Объем (Сумма)", index=False)
        result.to_excel(writer, sheet_name="Вес и Объём (Срзнач)", index=False)


        # Получите объекты 'workbook' и 'worksheet' из ExcelWriter для первого листа
        workbook = writer.book
        worksheet = writer.sheets['Сверка сроков доставки']

        # Получите объекты 'workbook' и 'worksheet' из ExcelWriter для первого листа
        workbook_2 = writer.book
        worksheet_2 = writer.sheets['Вес и Объем (Сумма)']

        # Создайте новые объекты 'workbook' и 'worksheet' для треьего листа
        workbook_3 = writer.book
        worksheet_3 = writer.sheets['Вес и Объём (Срзнач)']
        "------------------------------------------------------------------------------------------------"

        """АВТОШИРИНА СТОЛБЦОВ"""
        # Установите автоширину для каждого столбца, основываясь на его содержании для первого листа
        for col_num, col in enumerate(data.columns):
            max_len = max(data[col].astype(str).apply(len).max(), len(str(col)))
            worksheet.set_column(col_num, col_num, max_len)

        # Установите автоширину для каждого столбца, основываясь на его содержании для первого листа
        for col_num, col in enumerate(potok_client_transit.columns):
            max_len = max(potok_client_transit[col].astype(str).apply(len).max(), len(str(col)))
            worksheet_2.set_column(col_num, col_num, max_len)

        # Установите автоширину для каждого столбца, основываясь на его содержании для третьего листа
        for col_num, col in enumerate(result.columns):
            max_len = max(result[col].astype(str).apply(len).max(), len(str(col)))
            worksheet_3.set_column(col_num, col_num, max_len)
        """АВТОШИРИНА СТОЛБЦОВ"""

        "------------------------------------------------------------------------------------------------"

        """ГРАНИЦЫ СТОЛБЦОВ"""
        # Установите границы для всей таблицы
        worksheet.conditional_format(0, 0, len(data), len(data.columns) - 1,
                                    {'type': 'no_blanks',
                                    'format': workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})})


        # Установите границы для всей таблицы
        worksheet_2.conditional_format(0, 0, len(potok_client_transit), len(potok_client_transit.columns) - 1,
                                    {'type': 'no_blanks',
                                    'format': workbook_2.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})})
            
        # Установите границы для всей таблицы
        worksheet_3.conditional_format(0, 0, len(result), len(result.columns) - 1,
                                    {'type': 'no_blanks',
                                    'format': workbook_3.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})})
        """ГРАНИЦЫ СТОЛБЦОВ"""
        
        "------------------------------------------------------------------------------------------------"

        """ГРАНИЦЫ СТОЛБЦОВ ДЛЯ ПУСТЫХ И НЕ ПУСТЫХ ЯЧЕЕК"""

        # Установите границы для пустых и непустых ячеек
        for cell_type in ['no_blanks', 'blanks']:
            worksheet.conditional_format(0, 0, len(data), len(data.columns) - 1,
                                        {'type': cell_type,
                                        'format': workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})})
                
        # Установите границы для пустых и непустых ячеек
        for cell_type in ['no_blanks', 'blanks']:
                worksheet_2.conditional_format(0, 0, len(potok_client_transit), len(potok_client_transit.columns) - 1,
                                    {'type': cell_type,
                                        'format': workbook_2.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})})
                
        # Установите границы для пустых и непустых ячеек
        for cell_type in ['no_blanks', 'blanks']:
            worksheet_3.conditional_format(0, 0, len(result), len(result.columns) - 1,
                                        {'type': cell_type,
                                        'format': workbook_3.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})})    
        """ГРАНИЦЫ СТОЛБЦОВ ДЛЯ ПУСТЫХ И НЕ ПУСТЫХ ЯЧЕЕК"""
                
        "------------------------------------------------------------------------------------------------"   

        """ПЕРЕНОС ТЕКСТА ДЛЯ ЗАГОЛОВКОВ"""
        # Формат заголовков
        header_format_1 = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': 'gray', 'bold': True})
        header_format_2 = workbook_2.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': 'gray', 'bold': True})
        header_format_3 = workbook_3.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': 'gray', 'bold': True})
                
        # Установите перенос текста для заголовков
        for col_num, value in enumerate(data.columns.values):
            worksheet.write_string(0, col_num, value, header_format_1)

        # Установите перенос текста для заголовков
        for col_num, value in enumerate(potok_client_transit.columns.values):
            worksheet_2.write_string(0, col_num, value, header_format_2)

        # Установите перенос текста для заголовков
        for col_num, value in enumerate(result.columns.values):
            worksheet_3.write_string(0, col_num, value, header_format_3)
        """ПЕРЕНОС ТЕКСТА ДЛЯ ЗАГОЛОВКОВ"""

        "------------------------------------------------------------------------------------------------"


        """ДЕЛАЕМ ФИЛЬТРЫ"""
        # Добавьте фильтрацию к первому листу
        worksheet.autofilter(0, 0, len(data), len(data.columns) - 1)

        # Добавьте фильтрацию к первому листу
        worksheet_2.autofilter(0, 0, len(potok_client_transit), len(potok_client_transit.columns) - 1)

        # Добавьте фильтрацию ко второму листу
        worksheet_3.autofilter(0, 0, len(result), len(result.columns) - 1)
        """ДЕЛАЕМ ФИЛЬТРЫ"""

        "------------------------------------------------------------------------------------------------"


        """КРАСИМ СТРОКИ СО ЗНАЧЕНИЕМ В СТОЛБЦЕ 'ДЕНЬ НЕДЕЛИ' ПО ЗНАЧЕНИЮ 'ВС' ДЛЯ ЧИТАЕМОСТИ ДАННЫХ """
        # Определите диапазон таблицы
        start_row = 1  # начинается с первой строки (нумерация с 0)
        end_row = len(result)  # здесь предполагается, что result - это DataFrame или что-то подобное
        num_columns = len(result.columns)

        # Создайте формат для серого фона с выравниванием по центру
        gray_format = workbook_3.add_format({'bg_color': '#d9d9d9', 'align': 'center', 'valign': 'vcenter'})

        # Определите диапазон таблицы
        start_row = 1  # начинается с первой строки (нумерация с 0)
        end_row = len(result)  # здесь предполагается, что result - это DataFrame или что-то подобное
        num_columns = len(result.columns)

        # Проверьте условие в столбце 'День недели' и окрасьте только ячейки в пределах таблицы
        for idx, value in enumerate(result['День недели']):
            if value == 'вс' and start_row <= idx + 1 < end_row:
                # Используйте 'range' для ограничения форматирования только в пределах таблицы
                cell_range = 'A{}:{}'.format(idx + 2, chr(ord('A') + num_columns - 1) + str(idx + 2))
                # Первое условие для непустых ячеек
                worksheet_3.conditional_format(cell_range, {'type': 'no_blanks', 'format': gray_format})
                # Второе условие для пустых ячеек
                worksheet_3.conditional_format(cell_range, {'type': 'blanks', 'format': gray_format})
        """КРАСИМ СТРОКИ СО ЗНАЧЕНИЕМ В СТОЛБЦЕ 'ДЕНЬ НЕДЕЛИ' ПО ЗНАЧЕНИЮ 'ВС' ДЛЯ ЧИТАЕМОСТИ ДАННЫХ """


    result.to_html(output_file_html, index=False)
    print(True)



load_obj = utils.load_file_obj(sys.argv[1:])
output_file_excel = load_obj["output_file_excel"]
output_file_html = load_obj["output_file_html"]
files = load_obj["files"]

isExistData = False
current_data = None

for file in files:
    for sheet_name, sheet in file.items():
    # в списке files по очереди будут имена всех выгрузок загруженных через приложение
    # тут можно их все загрузить и как то обработать,
        current_data = sheet
        isExistData = True
    # текущее название одной из выгрузок (загруженных файлов через приложение) 

if isExistData:
    run_script(current_data, output_file_excel, output_file_html)
else:
    print(False)
