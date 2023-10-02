import sys
import os

# для работы import utils нужно подтянуть пути проекта
PROJECT_ROOT = os.path.abspath(os.path.join(
                  os.path.dirname(__file__), 
                  os.pardir)
)
sys.path.append(PROJECT_ROOT)
import utils

import datetime
import json
import openpyxl
import xlsxwriter
import pandas as pd
import numpy as np
from dotenv import load_dotenv
load_dotenv()
pd.options.mode.chained_assignment = None

def check_days_in_value(value):
    days = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']
    for day in days:
        if day in value:
            return False
    return True

def get_current_day_count_in_value(value, day):
    days = value.split(',')
    return days.count(day)

def check_plan_fact(value):
    if value < 0:
        return 'По факту было меньше машин'
    elif value > 0:
        return 'По факту было больше машин'
    else:
        return 'план-факт машин совпадает'

def day_of_week(date):
    day_number = date.isoweekday()
    if day_number == 1:
        return "пн"
    elif day_number == 2:
        return "вт"
    elif day_number == 3:
        return "ср"
    elif day_number == 4:
        return "чт"
    elif day_number == 5:
        return "пт"
    elif day_number == 6:
        return "сб"
    elif day_number == 7:
        return "вс"
    else:
        return "ошибка"

def run_script(plan_data, fact_data, output_file_excel, output_file_html):

    fact_data['день недели'] = fact_data['Эт:ФктДатО'].apply(day_of_week)

    plan_data = plan_data.dropna(subset=['НаправПВГ'])
    # Наименование завода и НаправПВГ
    plan_data = plan_data.drop_duplicates(subset=['Наименование завода', 'НаправПВГ'])

    # ЗвдОтпрвк, Назн.:Звд, Эт:ФктДатО
    fact_data = fact_data.drop_duplicates(subset=['ЗвдОтпрвк', 'Назн.:Звд', 'Эт:ФктДатО'])

    values_to_add_pref = utils.load_settings_table_column_values('дивизионы.json', 'Префикс')
    values_to_add_rp = utils.load_settings_table_column_values('дивизионы.json', 'РП')
    values_to_add_div = utils.load_settings_table_column_values('дивизионы.json', 'Дивизион')
    values_to_add_rp_num = utils.load_settings_table_column_values('дивизионы.json', 'Номер города')

    items_pref_rp = dict(zip(values_to_add_pref, values_to_add_rp))
    items_num_rp = dict(zip(values_to_add_rp_num, values_to_add_rp))
    items_rp_div = dict(zip(values_to_add_rp, values_to_add_div))

    plan_data['куда'] = plan_data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='НаправПВГ', items_list=items_pref_rp)
    plan_data['див'] = plan_data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наименование завода', items_list=items_rp_div)

    fact_data['з-д откуда'] = fact_data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='ЗвдОтпрвк', items_list=items_num_rp)
    fact_data['з-д куда'] = fact_data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Назн.:Звд', items_list=items_num_rp)
    fact_data['див'] = fact_data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='з-д откуда', items_list=items_rp_div)

     # удалить рп и див по списку в инструкции
    plan_data = plan_data[~plan_data['Наименование завода'].str.contains('Москва-Щербинка')]
    plan_data = plan_data.loc[plan_data['див'] != 'Ритейл']
    plan_data = plan_data.loc[plan_data['див'] != 'Аэропорт']
    fact_data = fact_data[~fact_data['з-д откуда'].str.contains('Шенкер')]
    fact_data = fact_data[~fact_data['з-д откуда'].str.contains('Москва-Щербинка')]
    fact_data = fact_data.loc[fact_data['див'] != 'Аэропорт']

    plan_data['откуда-куда план'] = plan_data['Наименование завода'].str.cat(plan_data['куда'], sep=';')
    plan_data = plan_data.sort_values(by='откуда-куда план')

    fact_data['откуда-куда факт'] = fact_data['з-д откуда'].str.cat(fact_data['з-д куда'], sep=';')

    res_dict = {}
    for index, row in fact_data.iterrows():
        key = row['откуда-куда факт']
        value = row['день недели']
        if key in res_dict:
            res_dict[key] += "," + value
        else:
            res_dict[key] = value

    Sheet2 = pd.DataFrame.from_dict(res_dict, orient='index', columns=['день недели'])
    Sheet2.index.name = 'откуда-куда факт'
    Sheet2 = Sheet2.sort_values(by='откуда-куда факт')
    Sheet2.reset_index(inplace=True)

    # Объединение таблиц по колонкам
    merged_df = plan_data.merge(Sheet2, left_on='откуда-куда план', right_on='откуда-куда факт')

    merged_df = merged_df.dropna(subset=['ГрафДвижТС'])

    mask = merged_df['ГрафДвижТС'].apply(check_days_in_value)

    merged_df = merged_df[~mask]

    merged_df['кол-во машин факт'] = merged_df['день недели'].str.split(',').str.len()
    merged_df['кол-во машин план'] = merged_df['ГрафДвижТС'].str.split(',').str.len()

    merged_df['сравнение план-факт'] = merged_df['кол-во машин факт'] - merged_df['кол-во машин план']
    merged_df['Выводы по количеству машин план-факт'] = merged_df['сравнение план-факт'].apply(check_plan_fact)

    result_table = pd.DataFrame()

    result_table['откуда-куда'] = merged_df['откуда-куда факт']

    result_table['пн(план)'] = merged_df['ГрафДвижТС'].apply(get_current_day_count_in_value, day='пн')
    result_table['вт(план)'] = merged_df['ГрафДвижТС'].apply(get_current_day_count_in_value, day='вт')
    result_table['ср(план)'] = merged_df['ГрафДвижТС'].apply(get_current_day_count_in_value, day='ср')
    result_table['чт(план)'] = merged_df['ГрафДвижТС'].apply(get_current_day_count_in_value, day='чт')
    result_table['пт(план)'] = merged_df['ГрафДвижТС'].apply(get_current_day_count_in_value, day='пт')
    result_table['сб(план)'] = merged_df['ГрафДвижТС'].apply(get_current_day_count_in_value, day='сб')
    result_table['вс(план)'] = merged_df['ГрафДвижТС'].apply(get_current_day_count_in_value, day='вс')

    result_table['пн(факт)'] = merged_df['день недели'].apply(get_current_day_count_in_value, day='пн')
    result_table['вт(факт)'] = merged_df['день недели'].apply(get_current_day_count_in_value, day='вт')
    result_table['ср(факт)'] = merged_df['день недели'].apply(get_current_day_count_in_value, day='ср')
    result_table['чт(факт)'] = merged_df['день недели'].apply(get_current_day_count_in_value, day='чт')
    result_table['пт(факт)'] = merged_df['день недели'].apply(get_current_day_count_in_value, day='пт')
    result_table['сб(факт)'] = merged_df['день недели'].apply(get_current_day_count_in_value, day='сб')
    result_table['вс(факт)'] = merged_df['день недели'].apply(get_current_day_count_in_value, day='вс')

    result_table['Итого пн'] = (result_table['пн(план)'] > 0) & (result_table['пн(факт)'] >= result_table['пн(план)'])
    result_table['Итого вт'] = (result_table['вт(план)'] > 0) & (result_table['вт(факт)'] >= result_table['вт(план)'])
    result_table['Итого ср'] = (result_table['ср(план)'] > 0) & (result_table['ср(факт)'] >= result_table['ср(план)'])
    result_table['Итого чт'] = (result_table['чт(план)'] > 0) & (result_table['чт(факт)'] >= result_table['чт(план)'])
    result_table['Итого пт'] = (result_table['пт(план)'] > 0) & (result_table['пт(факт)'] >= result_table['пт(план)'])
    result_table['Итого сб'] = (result_table['сб(план)'] > 0) & (result_table['сб(факт)'] >= result_table['сб(план)'])
    result_table['Итого вс'] = (result_table['вс(план)'] > 0) & (result_table['вс(факт)'] >= result_table['вс(план)'])

    result_table['пн(план/факт)'] = ''
    result_table['вт(план/факт)'] = ''
    result_table['ср(план/факт)'] = ''
    result_table['чт(план/факт)'] = ''
    result_table['пт(план/факт)'] = ''
    result_table['сб(план/факт)'] = ''
    result_table['вс(план/факт)'] = ''

    mask1 = result_table['пн(план)'] > 0
    mask2 = result_table['вт(план)'] > 0
    mask3 = result_table['ср(план)'] > 0
    mask4 = result_table['чт(план)'] > 0
    mask5 = result_table['пт(план)'] > 0
    mask6 = result_table['сб(план)'] > 0
    mask7 = result_table['вс(план)'] > 0

    result_table.loc[mask1, 'пн(план/факт)'] = result_table['пн(план)'].astype(str) + ' / ' + result_table['пн(факт)'].astype(str)
    result_table.loc[mask2, 'вт(план/факт)'] = result_table['вт(план)'].astype(str) + ' / ' + result_table['вт(факт)'].astype(str)
    result_table.loc[mask3, 'ср(план/факт)'] = result_table['ср(план)'].astype(str) + ' / ' + result_table['ср(факт)'].astype(str)
    result_table.loc[mask4, 'чт(план/факт)'] = result_table['чт(план)'].astype(str) + ' / ' + result_table['чт(факт)'].astype(str)
    result_table.loc[mask5, 'пт(план/факт)'] = result_table['пт(план)'].astype(str) + ' / ' + result_table['пт(факт)'].astype(str)
    result_table.loc[mask6, 'сб(план/факт)'] = result_table['сб(план)'].astype(str) + ' / ' + result_table['сб(факт)'].astype(str)
    result_table.loc[mask7, 'вс(план/факт)'] = result_table['вс(план)'].astype(str) + ' / ' + result_table['вс(факт)'].astype(str)

    result_table['Итого по факт'] = result_table['Итого пн'].astype(int) + result_table['Итого вт'].astype(int) + result_table['Итого ср'].astype(int) + result_table['Итого чт'].astype(int) + result_table['Итого пт'].astype(int) + result_table['Итого сб'].astype(int) + result_table['Итого вс'].astype(int)
    result_table['Итого по плану'] = merged_df['кол-во машин план']

    percentage_column = (result_table['Итого по факт'] / result_table['Итого по плану'])
    result_table['% качества'] = percentage_column.round(4)
    result_table['Выводы'] = np.where(result_table['% качества'] == 1, 'В соответствии с графиком', 'с нарушением графика')

    columns_to_remove = ['пн(план)', 'вт(план)', 'ср(план)', 'чт(план)', 'пт(план)',
                     'сб(план)', 'вс(план)', 'пн(факт)', 'вт(факт)', 'ср(факт)',
                     'чт(факт)', 'пт(факт)', 'сб(факт)', 'вс(факт)', 'Итого пн',
                     'Итого вт', 'Итого ср', 'Итого чт', 'Итого пт', 'Итого сб', 'Итого вс']

    result_table = result_table.drop(columns=columns_to_remove)

    with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
        book = writer.book
        percent_format = book.add_format({'num_format': '0.00%'})
        bold_format = book.add_format({'bold': True})

        Sheet2.to_excel(writer, sheet_name="Sheet3", index=False)
        merged_df.to_excel(writer, sheet_name="Sheet4", index=False)
        result_table.to_excel(writer, sheet_name="Итоги", index=False)
        worksheet = writer.sheets["Итоги"]

        worksheet.set_column('K:K', 16, percent_format)
        worksheet.set_column('A:A', 28)
        worksheet.set_column('B:J', 16)
        worksheet.set_column('L:L', 28)

    result_table.to_html(output_file_html, index=False)

    print(True)


load_obj = utils.load_file_obj(sys.argv[1:])
output_file_excel = load_obj["output_file_excel"]
output_file_html = load_obj["output_file_html"]
files = load_obj["files"]

plan_data = None
fact_data = None
isExistPlan = False
isExistFact = False

for file in files:
    for sheet_name, sheet in file.items():
        current_sheet = sheet
        if 'Наим операции' in current_sheet.columns:
            if current_sheet['Наим операции'].iloc[0] == 'Принятие транзитного груза' or current_sheet['Наим операции'].iloc[0] == 'Забор из города' or current_sheet['Наим операции'].iloc[0] == 'Получение от клиента':
                plan_data = current_sheet
                isExistPlan = True
        elif 'Узел отправки' in current_sheet.columns and 'Целевой узел' in current_sheet.columns:
            fact_data = current_sheet
            isExistFact = True
if isExistPlan and isExistFact:
    run_script(plan_data, fact_data, output_file_excel, output_file_html)
else:
    print(False)