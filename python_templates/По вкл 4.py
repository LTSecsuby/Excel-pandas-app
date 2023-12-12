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

    values_to_add_rp = utils.load_settings_table_column_values('дивизионы.json', 'РП')
    values_to_add_div = utils.load_settings_table_column_values('дивизионы.json', 'Дивизион')
    items_div = dict(zip(values_to_add_rp, values_to_add_div))

    data['Дивизион'] = data.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Наименование завода', items_list=items_div, default_value='Пустой дивизион')

    error = utils.check_null_pointer_in_table_value(data, 'Дивизион', 'Наименование завода', 'Пустой дивизион')
    if error:
        print('unknowns_division')
    else:
        data = data.drop_duplicates(subset='№ ЭР')
        data_1 = pd.merge(data, data_hour_p, on ='Наименование завода отправителя', how ='left')
        data_1 = data_1.drop(['Часовой пояс'], axis=1)
        data_1 = data_1.drop(['Значение'], axis=1)
        data_1 = data_1.drop(['формула дельта в часах ко времени ЕКБ'], axis=1)
        data_1['дельта в часах ко времени ЕКБ'] = data_1['дельта в часах ко времени ЕКБ'].astype(str)
        data_1['дельта в часах ко времени ЕКБ'] = data_1['дельта в часах ко времени ЕКБ'].str.replace('.0', '')
        data_1 = data_1[data_1['дельта в часах ко времени ЕКБ'] != 'nan']
        data_1['дельта в часах ко времени ЕКБ'] = data_1['дельта в часах ко времени ЕКБ'].astype(int).abs()
        data_1['дельта в часах ко времени ЕКБ']
        data_2 = pd.merge(data_1, data_hour_p, on ='Наименование завода отправителя', how ='left')
        data_2 = data_2.drop(['Часовой пояс'], axis=1)
        data_2 = data_2.drop(['формула дельта в часах ко времени ЕКБ'], axis=1)
        data_2 = data_2.drop(['дельта в часах ко времени ЕКБ_y'], axis=1)
        data_2 = data_2.rename(columns={'Значение' : 'Прибавить'})
        data_2['дельта в часах ко времени ЕКБ_x'] = data_2['дельта в часах ко времени ЕКБ_x'].astype(str)
        data_2['дельта в часах ко времени ЕКБ_x'] = data_2['дельта в часах ко времени ЕКБ_x'].str.replace('.0', '')
        data_2 = data_2[data_2['дельта в часах ко времени ЕКБ_x'] != 'nan']
        data_2['дельта в часах ко времени ЕКБ_x'] = data_2['дельта в часах ко времени ЕКБ_x'].astype(int)
        data_2['time'] = data_2['Дата создания'].astype(str) + " " + data_2['Время'].astype(str)
        data_2['time'] = pd.to_datetime(data_2['time'])
        def time_lp(data_2):
            if data_2['Прибавить'] == 'прибавить':
                return data_2['time'] + relativedelta(hours=data_2['дельта в часах ко времени ЕКБ_x'])
            else:
                return data_2['time'] - relativedelta(hours=data_2['дельта в часах ко времени ЕКБ_x'])
        data_2['дата и время создания с учетом час пояса'] = data_2.apply(time_lp, axis=1)
        data_2['Дата создания с час поясом'] = pd.to_datetime(data_2['дата и время создания с учетом час пояса']).dt.date
        data_2['Время создания с час поясом'] = pd.to_datetime(data_2['дата и время создания с учетом час пояса']).dt.time
        data_2['День недели'] = pd.to_datetime(data_2['Дата создания с час поясом']).dt.dayofweek
        data_2['Время создания с час поясом'] = pd.to_datetime(data_2['дата и время создания с учетом час пояса']).dt.hour
        def time_lp_1(data_2):
            if (data_2['День недели'] == 4 and data_2['Время создания с час поясом'] > 12):
                return data_2['дата и время создания с учетом час пояса'] + relativedelta(days=+3)
            elif data_2['День недели'] == 5:
                return data_2['дата и время создания с учетом час пояса'] + relativedelta(days=+2)
            elif data_2['День недели'] == 6:
                return data_2['дата и время создания с учетом час пояса'] + relativedelta(days=+1)
            else:
                return data_2['Дата создания с час поясом']
        data_2['Дата создания с учетом выходных'] = data_2.apply(time_lp_1, axis=1)
        data_2['Дата создания с учетом выходных'] = data_2['Дата создания с учетом выходных'].dt.date
        data_2['Плановая дата забора'] = data_2['Плановая дата забора'].astype(str)
        def time_lp_2(data_2):
            if data_2['Плановая дата забора'] == 'NaT':
                return 'убрать'
            else:
                return ''
        data_2['убрать нет план даты'] = data_2.apply(time_lp_2, axis=1)
        data_2['Дата статуса  "Груз принят"'] = data_2['Дата статуса  "Груз принят"'].astype(str)
        data_2['Плановая дата забора'] = pd.to_datetime(data_2['Плановая дата забора'])
        def time_lp_3(data_2):
            if (data_2['Дата статуса  "Груз принят"'] == 'NaT'):
                if (data_2['Плановая дата забора'] > datetime.now()):
                    return 'убрать'
            else:
                return ''
        data_2['убрать план дата больше даты создания отчета'] = data_2.apply(time_lp_3, axis=1)
        data_2['Дата статуса  "Груз принят"'] = pd.to_datetime(data_2['Дата статуса  "Груз принят"'])
        data_div = data_div.rename(columns={'РП' : 'Наименование завода отправителя'})
        data_3 = pd.merge(data_2, data_div, on ='Наименование завода отправителя', how ='left')
        data_3['Срок забора без выходных'] = (data_3['Дата статуса  "Груз принят"'] - data_3['Дата создания'])
        data_3['Срок забора без выходных'] = data_3['Срок забора без выходных'].dt.days
        data_3['Срок забора с выходными'] = (data_3['Дата статуса  "Груз принят"'] - pd.to_datetime(data_3['Дата создания с учетом выходных']))
        data_3['Срок забора с выходными'] = data_3['Срок забора с выходными'].dt.days
        data_3['Срок забора с выходными'] = data_3['Срок забора с выходными'].replace(-3, 0)
        data_3['Срок забора с выходными'] = data_3['Срок забора с выходными'].replace(-2, 1)
        data_3['Срок забора с выходными'] = data_3['Срок забора с выходными'].replace(-1, 2)
        data_3['меньшее'] = np.minimum(data_3['Срок забора без выходных'], data_3['Срок забора с выходными'])
        data_3['меньшее'] = data_3['меньшее'].fillna(-45000)
        data_3['убрать нет план даты'] = data_3['убрать нет план даты'].astype(str)
        data_3['убрать план дата больше даты создания отчета'] = data_3['убрать план дата больше даты создания отчета'].astype(str)
        def minimum(data_3):
            if data_3['меньшее'] == 0:
                return 'Забор день в день'
            elif data_3['меньшее'] == 1:
                return 'Забор на следующий день'
            elif data_3['меньшее'] == 2:
                return 'Забор через 2 дня'
            elif data_3['меньшее'] == 3:
                return 'Забор через 3 дня'
            elif data_3['меньшее'] == 4:
                return 'Забор через 4-5 дней'
            elif data_3['меньшее'] == 5:
                return 'Забор через 4-5 дней'
            elif data_3['меньшее'] >= 6:
                return 'Забор через 6 дней и более'
            else:
                return 'Забор груза не состоялся (возможно отсутствует услуга забора)'
        data_3['периоды'] = data_3.apply(minimum, axis=1)
        data_end = pd.merge(data_3, data_dop_scl, on = 'Наименование завода отправителя', how='left')
        data_end = data_end[data_end['Удалить из отчета ГЛ'] != 'удалить']
        data_end = data_end[data_end['убрать нет план даты'] != 'убрать']
        data_end = data_end[data_end['убрать план дата больше даты создания отчета'] != 'убрать']
        table = pd.pivot_table(data_end, values='№ ЭР', index='Дивизион', aggfunc='count', columns='периоды', margins_name='Итог количество', margins=True, dropna=True)
        table = table.reset_index()
        table['% от количества_1'] = (table['Забор груза не состоялся (возможно отсутствует услуга забора)'] / table['Итог количество']*100).round(2)
        table['% от количества_2'] = (table['Забор день в день'] / table['Итог количество']*100).round(2)
        table['% от количества_3'] = (table['Забор на следующий день'] / table['Итог количество']*100).round(2)
        table['% от количества_4'] = (table['Забор через 2 дня'] / table['Итог количество']*100).round(2)
        table['% от количества_5'] = (table['Забор через 3 дня'] / table['Итог количество']*100).round(2)
        table['% от количества_6'] = (table['Забор через 4-5 дней'] / table['Итог количество']*100).round(2)
        table['% от количества_7'] = (table['Забор через 6 дней и более'] / table['Итог количество']*100).round(2)
        table['% от количества_sum'] = (table['Итог количество'] / table['Итог количество']*100).round(2)
        table_2 = pd.pivot_table(data_end, values='№ ЭР', index=['Наименование завода отправителя'],
aggfunc='count', columns='периоды', margins_name='Итог количество', margins=True, dropna=True)
        table_2.reset_index()
        table_2['% от количества_1'] = (table_2['Забор груза не состоялся (возможно отсутствует услуга забора)'] / table_2['Итог количество']*100).round(2)
        table_2['% от количества_2'] = (table_2['Забор день в день'] / table_2['Итог количество']*100).round(2)
        table_2['% от количества_3'] = (table_2['Забор на следующий день'] / table_2['Итог количество']*100).round(2)
        table_2['% от количества_4'] = (table_2['Забор через 2 дня'] / table_2['Итог количество']*100).round(2)
        table_2['% от количества_5'] = (table_2['Забор через 3 дня'] / table_2['Итог количество']*100).round(2)
        table_2['% от количества_6'] = (table_2['Забор через 4-5 дней'] / table_2['Итог количество']*100).round(2)
        table_2['% от количества_7'] = (table_2['Забор через 6 дней и более'] / table_2['Итог количество']*100).round(2)
        table_2['% от количества_sum'] = (table_2['Итог количество'] / table_2['Итог количество']*100).round(2)
        table_2 = table_2[['Забор груза не состоялся (возможно отсутствует услуга забора)', '% от количества_1',
               'Забор день в день', '% от количества_2', 'Забор на следующий день', '% от количества_3',
              'Забор через 2 дня', '% от количества_4', 'Забор через 3 дня', '% от количества_5', 'Забор через 4-5 дней',
               '% от количества_6', 'Забор через 6 дней и более', '% от количества_7', 'Итог количество', '% от количества_sum']]
        table_2 = table_2.reset_index()
        table_itog = table_2.append(table)
        table_itog = table_itog[['Забор груза не состоялся (возможно отсутствует услуга забора)', '% от количества_1',
               'Забор день в день', '% от количества_2', 'Забор на следующий день', '% от количества_3',
              'Забор через 2 дня', '% от количества_4', 'Забор через 3 дня', '% от количества_5', 'Забор через 4-5 дней',
               '% от количества_6', 'Забор через 6 дней и более', '% от количества_7', 'Итог количество', '% от количества_sum']]
        table_itog = table_itog.reset_index()
        
        result_table = pd.DataFrame()        

        # !!!! ТУТ идет сохранение полученного результата, чтобы отправить обратно в приложение и скачать
        # тут можно накинуть стилей в уже сохраненные листы файла, сохранить нужные листы и тд (примеры ниже)
        with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
            # пересохраняем нужные листы
            v_1.to_excel(writer, sheet_name="Выгрузка_1", index=False)
            nd.to_excel(writer, sheet_name="Недостачи", index=False)
            v_drobl.to_excel(writer, sheet_name="Дробление", index=False)
            ezhedn_otchet.to_excel(writer, sheet_name="Ежедневный_отчет_общая", index=False)          

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
