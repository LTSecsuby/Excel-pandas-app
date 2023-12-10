# чтобы работать с настройками (не удалять)
import json
# библиотека pandas (не удалять)
import pandas as pd
# для работы с аргументами передаваемыми в скрипт (не удалять)
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


#  !!!! ТУТ начало скрипта, а сверху импорты нужных библиотек и тд
# в sys.argv хранится список агрументов (первый sys.argv[0] это служебный, а дальше это названия файлов загруженных через приложение)
# sys.argv[1] - загрузим первый файл, если их несколько то нужно загружать их в цикле for arg in sys.argv[1:]: 
# длина агрументов равна двум - значит загружен один файл, можно его обработать

load_obj = utils.load_file_obj(sys.argv[1:])
output_file_excel = load_obj["output_file_excel"]
output_file_html = load_obj["output_file_html"]
files = load_obj["files"]

current_data = None

# сюда попадаем если файлов больше одного, то есть их несколько, значит нужно считывать их в цикле for file in files:
# в file названия загруженных файлов (выгрузок)
for file in files:
    for sheet_name, sheet in file.items():
    # в списке files по очереди будут имена всех выгрузок загруженных через приложение
    # тут можно их все загрузить и как то обработать,
        current_data = sheet
    # текущее название одной из выгрузок (загруженных файлов через приложение) 

# тут мы взяли последнюю таблицу
Sheet1 = current_data

# !!!! ТУТ можно работать с загруженными выгрузками, они хранятся в словаре dict
# можно создать новую таблицу и сохранять туда нужные данные, а потом сохранить в приложение и скачать (например Sheet2 = pd.DataFrame())
Sheet2 = pd.DataFrame()




# как загрузить настройки созданные в приложении
# настройки хранятся в папке 'SAVED_SETTINGS_FILES_PATH' utils.createEnvPath создает путь до нужной настройки (например дивизионы.json)
# чтобы загрузить настройку по названию можно воспользоваться методом load_settings_table_column_values из utils:
# в метод в качестве параметров передается название настройки это например дивизионы.json или любое другое, а также название колонки в таблице
# это например Склад
# и в переменных будут лежать списки данных из соответствующей колонки настройки
values_to_add_stock = utils.load_settings_table_column_values('Див, номер завода.json', 'Категория')
values_to_add_rp_num = utils.load_settings_table_column_values('Див, номер завода.json', 'Номер завода')
values_to_add_rp = utils.load_settings_table_column_values('Див, номер завода.json', 'РП')
values_to_add_div = utils.load_settings_table_column_values('Див, номер завода.json', 'Дивизион ТК')

# далее списки можно объединить в словари ключ-значение, для удобных соответствий
# items_num_rp = dict(zip(values_to_add_rp_num, values_to_add_rp))
# items_div = dict(zip(values_to_add_rp, values_to_add_div))
# items_stock = dict(zip(values_to_add_rp, values_to_add_stock))
# или получить из них колонки со значениями например так:
# divisions = pd.DataFrame()
# divisions['РП'] = values_to_add_rp
# divisions['Дивизион'] = values_to_add_div


# теперь объединив данные из настроек в словарь можно проверить есть ли дивизион в items_div по колонке 'Завод пользователя', 
# check_value_in_list_and_set_value устанавливает 'Пустой дивизион' в 'Дивизион' если не нашел и соответствующий дивизион если нашел
# так можно проверить любую колонку на значения загруженные из настроек (новые настройки можно создавать)
# row - исходная таблица
# row_name - название столбца
# items_list - список столбцов для которых нужно сделать новые значения
# default_value - значение если в списке не найдей ключ
# Sheet1['Дивизион'] = Sheet1.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Завод пользователя', items_list=items_div, default_value='Пустой дивизион')
# далее можно проверить что в Sheet1['Дивизион'] установились все значения, то есть они все были найдены в настройках
# выше мы устанавили что если значений нет то устанавливаем значение 'Пустой дивизион'
# теперь мы сможем проверить если пустой дивизион в значениях с помощью check_null_pointer_in_table_value из utils

# в error из метода вернется значение правда или ложь и соответственно мы можем это обработать
# Sheet1 - это таблица для поиска
# 'Дивизион' - колонка в которой проверяем есть ли значение для поиска
# 'Завод пользователя' - название колонки откуда взять значение если не будет найдено значение для поиска,
# это нужно для того чтобы если произойдет исключение 'unknowns_division', то можно было бы узнать на каком значении оно произошло
# 'Пустой дивизион' - значение для поиска
# error = utils.check_null_pointer_in_table_value(Sheet1, 'Дивизион', 'Завод пользователя', 'Пустой дивизион')
# if error:
    # print('unknowns_division') - unknowns_division тут это особое исключение которое может обработать приложение и вернуть ошибку
    # при необходимости список исключений можно расширить и так приложение будет выдавать больше сообщений о проблемах пользователю
# else:
    # а тут продолжаем работу скрипта если не было пустых дивизионов, пишем тут код






# !!!! ТУТ идет сохранение полученного результата, чтобы отправить обратно в приложение и скачать
# тут можно накинуть стилей в уже сохраненные листы файла, сохранить нужные листы и тд (примеры ниже)
with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
    book = writer.book
    num_format = book.add_format({'num_format': '0'})
    wrap_format = book.add_format({'bold': True})
    wrap_format.set_text_wrap()

    # пересохраняем нужные листы
    Sheet1.to_excel(writer, sheet_name="Новое название", index=False)
    Sheet2.to_excel(writer, sheet_name="zimg", index=False)
    worksheet = writer.sheets["zimg"]

    # можно нужно колонке применить формат значений и стили
    worksheet.set_column('A:A', 30, num_format)
    worksheet.set_column('B:B', 30, num_format)
    # worksheet.set_column('C:C', 18, num_format)
    # worksheet.set_column('D:D', 30, num_format)
    # worksheet.set_column('E:E', 30, num_format)
    # worksheet.write(0, 0, "Дата/время размещения фотографий/документов", wrap_format)
    # worksheet.write(0, 1, "Номер документа-основания= Ключ объекта", wrap_format)
    # worksheet.write(0, 2, "Завод пользователя", wrap_format)
    # worksheet.write(0, 3, "Наименован завода польз", wrap_format)
    # worksheet.write(0, 4, "ввв", wrap_format)

# сохраним html для вывода таблицы в приложении в браузере
Sheet1.to_html(output_file_html, index=False)
# print(True) - заканчивает выполнение скрипта и выходит в приложение
print(True)

