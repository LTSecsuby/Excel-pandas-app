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

def get_fact_days(row):
    days_list = []
    result = ''
    if pd.notnull(row['пн']):
        days_list.append('пн')
    if pd.notnull(row['вт']):
        days_list.append('вт')
    if pd.notnull(row['ср']):
        days_list.append('ср')
    if pd.notnull(row['чт']):
        days_list.append('чт')
    if pd.notnull(row['пт']):
        days_list.append('пт')
    if pd.notnull(row['сб']):
        days_list.append('сб')
    if pd.notnull(row['вс']):
        days_list.append('вс')
    for day in days_list:
        if len(result) > 0:
            result += ','
        result += day
    return result

def get_merge_day(row, day):
    value = row[day]
    if pd.notnull(value):
        if day in row['График план'].split(','):
            value = 1
        else:
            value = 0
    else:
        if day in row['График план'].split(','):
            value = 0
    return value

def get_difference_count(row):
    count = 0
    if row['пн'] == 1:
        count += 1
    if row['вт'] == 1:
        count += 1
    if row['ср'] == 1:
        count += 1
    if row['чт'] == 1:
        count += 1
    if row['пт'] == 1:
        count += 1
    if row['сб'] == 1:
        count += 1
    if row['вс'] == 1:
        count += 1
    return count

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


# fact_data - МБ 51
def run_script(plan_data, fact_data, output_file_excel, output_file_html):

    fact_data = fact_data.dropna(subset=['Завод'])
    # Партия, Завод, Дата ввода
    fact_data = fact_data.drop_duplicates(subset=['Партия', 'Завод', 'Дата ввода'])

    fact_data['Партия'] = fact_data['Партия'].round().astype(int)
    fact_data['партия+завод'] = fact_data['Партия'].astype(str) + fact_data['Завод'].astype(str)

    plan_data = plan_data.dropna(subset=['НаправПВГ'])
    # Наименование завода и НаправПВГ и Партия
    plan_data = plan_data.drop_duplicates(subset=['Наименование завода', 'НаправПВГ', 'Партия'])

    plan_data['партия+завод'] = plan_data['Партия'].astype(str) + plan_data['Завод'].astype(str)

    values_to_add_pref = utils.load_settings_table_column_values('Див, номер завода.json', 'Префикс')
    values_to_add_rp = utils.load_settings_table_column_values('Див, номер завода.json', 'РП')
    values_to_add_div = utils.load_settings_table_column_values('Див, номер завода.json', 'Дивизион ТК')
    values_to_add_rp_num = utils.load_settings_table_column_values('Див, номер завода.json', 'Номер завода')

    items_pref_rp = dict(zip(values_to_add_pref, values_to_add_rp))
    items_num_rp = dict(zip(values_to_add_rp_num, values_to_add_rp))
    items_rp_div = dict(zip(values_to_add_rp, values_to_add_div))

    plan_data['див'] = plan_data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наименование завода', items_list=items_rp_div)
    plan_data['див откуда'] = plan_data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наим завод откуда', items_list=items_rp_div)
    plan_data['див куда'] = plan_data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наим завод куда', items_list=items_rp_div)

     # удалить рп и див по списку в инструкции
    plan_data = plan_data[~plan_data['Наименование завода'].str.contains('Москва-Щербинка')]
    plan_data = plan_data[~plan_data['Наим завод откуда'].str.contains('Москва-Щербинка')]
    plan_data = plan_data[~plan_data['Наим завод куда'].str.contains('Москва-Щербинка')]

    plan_data = plan_data.loc[plan_data['див'] != 'Ритейл']
    plan_data = plan_data.loc[plan_data['див'] != 'Аэропорт']
    plan_data = plan_data.loc[plan_data['див'] != 'удалить из выгрузки']

    plan_data = plan_data.loc[plan_data['див откуда'] != 'Ритейл']
    plan_data = plan_data.loc[plan_data['див откуда'] != 'Аэропорт']
    plan_data = plan_data.loc[plan_data['див откуда'] != 'удалить из выгрузки']

    plan_data = plan_data.loc[plan_data['див куда'] != 'Ритейл']
    plan_data = plan_data.loc[plan_data['див куда'] != 'Аэропорт']
    plan_data = plan_data.loc[plan_data['див куда'] != 'удалить из выгрузки']

    plan_data = plan_data.dropna(subset=['ГрафДвижТС'])
    mask = plan_data['ГрафДвижТС'].apply(check_days_in_value)
    plan_data = plan_data[~mask]

    merged_data = pd.DataFrame()
    merged_data = pd.merge(plan_data, fact_data[['партия+завод', 'Дата ввода']], on='партия+завод', how='left')
    merged_data.rename(columns={'Дата ввода': 'дата ввода из мб51'}, inplace=True)

    merged_data = merged_data.sort_values(by='дата ввода из мб51')
    merged_data['куда'] = merged_data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='НаправПВГ', items_list=items_pref_rp)

    merged_data = merged_data.dropna(subset=['дата ввода из мб51'])

    merged_data = merged_data.drop_duplicates(subset=['Наименование завода', 'куда', 'дата ввода из мб51'])

    merged_data['откуда-куда'] = merged_data['Наименование завода'].str.cat(merged_data['куда'], sep=';')

    merged_data['дата ввода из мб51 значение'] = pd.to_datetime(merged_data['дата ввода из мб51'], format='%Y-%m-%dT%H:%M:%S.%f')

    merged_data['График факт'] = merged_data['дата ввода из мб51 значение'].apply(day_of_week)

    temporary_table = pd.DataFrame()

    temporary_table = pd.pivot_table(merged_data,
            index='откуда-куда',
            columns='График факт',
            values='дата ввода из мб51 значение',
            aggfunc='count')

    temporary_table = temporary_table.reset_index()

    merged_df = pd.DataFrame()
    merged_df = pd.merge(temporary_table, merged_data[['откуда-куда', 'ГрафДвижТС']], on='откуда-куда', how='left')

    merged_df.rename(columns={'ГрафДвижТС': 'График план'}, inplace=True)

    merged_df = merged_df.drop_duplicates(subset=['График план', 'откуда-куда'])

    merged_df['График факт'] = merged_df.apply(get_fact_days, axis=1)

    merged_df['кол-во машин факт'] = merged_df['График факт'].str.split(',').str.len()
    merged_df['кол-во машин план'] = merged_df['График план'].str.split(',').str.len()

    merged_df['пн'] = merged_df.apply(get_merge_day, day='пн', axis=1)
    merged_df['вт'] = merged_df.apply(get_merge_day, day='вт', axis=1)
    merged_df['ср'] = merged_df.apply(get_merge_day, day='ср', axis=1)
    merged_df['чт'] = merged_df.apply(get_merge_day, day='чт', axis=1)
    merged_df['пт'] = merged_df.apply(get_merge_day, day='пт', axis=1)
    merged_df['сб'] = merged_df.apply(get_merge_day, day='сб', axis=1)
    merged_df['вс'] = merged_df.apply(get_merge_day, day='вс', axis=1)

    merged_df['итого/факт совпадений с планом'] = merged_df.apply(get_difference_count, axis=1)

    merged_df['% качества'] = (merged_df['итого/факт совпадений с планом'] / merged_df['кол-во машин план'] * 100).round(2)

    merged_df['Выводы'] = merged_df['% качества'].apply(lambda x: 'В соответствии с графиком' if x == 100 else 'С нарушением графика')

    with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
        book = writer.book

        merged_data.to_excel(writer, sheet_name="merged", index=False)
        temporary_table.to_excel(writer, sheet_name="сводная из сроков доставки", index=False)
        merged_df.to_excel(writer, sheet_name="итоги", index=False)
        worksheet = writer.sheets["итоги"]

        worksheet.set_column('A:A', 28)
        worksheet.set_column('I:J', 16)
        worksheet.set_column('O:O', 28)

    merged_df.to_html(output_file_html, index=False)

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
        elif 'Вид движения' in current_sheet.columns and 'Партия' in current_sheet.columns:
            if current_sheet['Вид движения'].iloc[0] == 641:
                fact_data = current_sheet
                isExistFact = True
if isExistPlan and isExistFact:
    run_script(plan_data, fact_data, output_file_excel, output_file_html)
else:
    print(False)