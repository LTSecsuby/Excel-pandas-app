# чтобы работать с настройками (не удалять)
import json
# библиотека pandas (не удалять)
import pandas as pd
# для работы с аргументами передаваемыми в скрипт (не удалять)
from datetime import datetime
import datetime as dt
import sys
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
    data = data.drop_duplicates()

    values_to_drop = utils.load_settings_table_column_values('удалить.json', 'Доп склад')
    data = data[~data['Наименование завода'].isin(values_to_drop)]

    values_to_add_rp = utils.load_settings_table_column_values('дивизионы.json', 'РП')
    values_to_add_div = utils.load_settings_table_column_values('дивизионы.json', 'Дивизион')
    items_div = dict(zip(values_to_add_rp, values_to_add_div))

    data['Дивизион'] = data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наименование завода', items_list=items_div, default_value='Пустой дивизион')

    error = utils.check_null_pointer_in_table_value(data, 'Дивизион', 'Наименование завода', 'Пустой дивизион')
    if error:
        print('unknowns_division')
    else:
        data_filter = data['План дата прибытия'].min()
        data['Груз к выдаче'] = data['Груз к выдаче'].astype(str)
        data['ДатаСписания'] = data['ДатаСписания'].astype(str)
        data['ДатаПогрузки'] = data['ДатаПогрузки'].astype(str)

        v_1 = data.loc[(data['План дата прибытия'] >= data_filter) & (data['Груз к выдаче'] == 'NaT') & (data['ДатаСписания'] == 'NaT') & (data['ДатаПогрузки'] == 'NaT')]
        v_1 = v_1.sort_values('ДатаПоступл', ascending=False)
        v_1['Партия'] = v_1['Партия'].astype(str)
        v_1['Завод'] = v_1['Завод'].astype(str)
        v_1 = v_1.sort_values(by='ДатаПоступл', ascending=False)
        v_1 = v_1.drop_duplicates(['Партия', 'Завод'])

        v_1 = v_1[v_1["Наименование завода"].str.contains("Ритейл")== False]
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

        data = data.sort_values('ДатаПоступл', ascending=False)
        data['Партия'] = data['Партия'].astype(str)
        data['Завод'] = data['Завод'].astype(str)
        data = data.drop_duplicates(['Партия', 'Завод'])

        data = data[data["Наименование завода"].str.contains("Ритейл")== False]
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

        data = data[data['Дивизион'] != 'Ритейл']

        ezhedn_otchet = pd.pivot_table(data, values='ЭР', index='Дивизион', aggfunc='count')

        data['Наименование завода'].agg({'ЭР' : 'count'})

        data[data['Наименование завода'] == 'Москва'].agg({'ЭР' : 'count'})

        x = data[data['Дивизион'] == 'ХМАО'].agg({'ЭР' : 'count'})
        y = data[data['Дивизион'] == 'ЯНАО'].agg({'ЭР' : 'count'})

        result_table = pd.DataFrame()
        result_table['Итог'] = x+y

        # !!!! ТУТ идет сохранение полученного результата, чтобы отправить обратно в приложение и скачать
        # тут можно накинуть стилей в уже сохраненные листы файла, сохранить нужные листы и тд (примеры ниже)
        with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
            # пересохраняем нужные листы
            v_1.to_excel(writer, sheet_name="Выгрузка_1", index=False)
            nd.to_excel(writer, sheet_name="Недостачи_1", index=False)
            v_drobl.to_excel(writer, sheet_name="Недостачи", index=False)
            ezhedn_otchet.to_excel(writer, sheet_name="Ежедневный_отчет_общая", index=False)
            result_table.to_excel(writer, sheet_name="Итог", index=False)

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
