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
if len(sys.argv) < 2:
    # длина агрументов меньше двух - значит нет файлов, кидаем исключение и выходим из скрипта, вызов print(False) закончит выполнение срипта,
    # приложение обработает исключение и покажет ошибку
    print(False)
# sys.argv[1] - загрузим первый файл, если их несколько то нужно загружать их в цикле for arg in sys.argv[1:]: 
# длина агрументов равна двум - значит загружен один файл, можно его обработать
elif len(sys.argv) == 2:
    # !!!! ТУТ загрузка выгрузки в скрипте, чтобы с ней можно было работать
    # sys.argv[0] трогать не нужно это служебный агрумент, начнем с sys.argv[1], там имя выгрузки
    file_name = sys.argv[1]
    # 'SAVED_FILES_PATH' папка для сохранения выгрузок (загруженных файлов в систему)
    # функция utils.createEnvPath создает правильный путь для сохранения выгрузок в приложение (до места где хранится загруженный файл + второй аргумент это название файла)
    excel_file = utils.createEnvPath('SAVED_FILES_PATH', file_name)
    # read_excel() загрузит первую страницу файла который находится по пути excel_file и поместит таблицу в data
    data = pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')




    # !!!! ТУТ выгрузка уже загружена, тут можно работать с ней, как то менять данные и добавлять/удалять колонки
    # пишите тут свой код



    # !!!! ТУТ идет сохранение полученного результата, чтобы отправить обратно в приложение и скачать
    # 'PYTHON_SAVED_FILES_PATH' папка для сохранения обработанных скриптом и сохраненных файлов с возможностью выгрузить их из приложения
    # функция utils.createEnvPath создает правильный путь для сохранения (до места где хранится обработанные скриптом файлы + второй аргумент это название файла)
    output_file_excel = utils.createEnvPath('PYTHON_SAVED_FILES_PATH', file_name)
    # у пути к файлу excel_file и output_file_excel обязательно должно быть одно имя file_name, так как при загрузки в приложение выгрузки
    # ей генерируется имя и нужно сопоставить загруженный и выгруженный файл

    # output_file_html путь для сохрания таблицы в формате html, это вернется в приложение в браузер в качестве ответа и покажется в виде таблицы
    output_file_html = os.path.splitext(output_file_excel)[0] + '.html'

    # !!!! ТУТ сохраняется измененая таблица data в виде excel и html
    # если в процессе выполнения вы решили создать еще таблиц (например Sheet2 = pd.DataFrame()) и сохранить/показать уже ее
    # то нужно вызвать to_excel() и to_html() уже у новой таблицы (например Sheet2.to_excel(output_file_excel, index=False)/Sheet2.to_html(output_file_html, index=False))
    data.to_excel(output_file_excel, index=False)
    data.to_html(output_file_html, index=False)
    # print(True) - заканчивает выполнение скрипта и выходит в приложение
    print(True)
else:
    # сюда попадаем если файлов больше одного, то есть их несколько, значит нужно считывать их в цикле for arg in sys.argv[1:]:
    # в arg названия загруженных файлов (выгрузок)
    dict = {}
    # ниже загрузим все файлы, но сохраним название первого, по этому названию потом будем сохранять результат работы скрипта 
    save_file_name = sys.argv[1]
    for arg in sys.argv[1:]:
        # sys.argv[0] трогать не нужно это служебный агрумент, названия файлов начинаются с sys.argv[1], пропустим служебный
        if sys.argv[0] != arg:
            # в arg имя одной из выгрузок, они в списке и тут по очереди будут имена всех выгрузок загруженных через приложение
            # тут можно их все загрузить и как то обработать,
            # например поместить их в словарь(dict), где ключ это название выгрузки, а значение это таблица
            
            # текущее название одной из выгрузок (загруженных файлов через приложение) 
            file_name = arg
            
            excel_file = utils.createEnvPath('SAVED_FILES_PATH', file_name)
            # прочитаем самую первую страницу или если знаем название, то можно указать его pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')
            current_sheet = pd.read_excel(excel_file, engine='openpyxl')

            # сохраняем текущую в dict теперь ключ/значение - это название/таблица
            dict[file_name] = current_sheet


    # !!!! ТУТ можно работать с загруженными выгрузками, они хранятся в словаре dict
    # можно создать новую таблицу и сохранять туда нужные данные, а потом сохранить в приложение и скачать (например Sheet2 = pd.DataFrame())
    Sheet2 = pd.DataFrame()
    # тут мы взяли первую таблицу (сохранять ниже обязательно по созданному пути с названием save_file_name - см. output_file)
    Sheet1 = dict[save_file_name]


    # как загрузить настройки созданные в приложении
    # настройки хранятся в папке 'SAVED_SETTINGS_FILES_PATH' utils.createEnvPath создает путь до нужной настройки (например дивизионы.json)
    json_file = utils.createEnvPath('SAVED_SETTINGS_FILES_PATH', 'дивизионы.json')
    with open(json_file, encoding="utf-8") as f:
        load_json = json.load(f)
        # в переменных лежат списки данных из соответствующей колонки настройки
        # values_to_add_stock = load_json['table'][0]['values']
        # values_to_add_rp_num = load_json['table'][1]['values']
        # values_to_add_rp = load_json['table'][2]['values']
        # values_to_add_div = load_json['table'][3]['values']

        # списки можно объединить в словари ключ-значение, для удобных соответствий
        # items_num_rp = dict(zip(values_to_add_rp_num, values_to_add_rp))
        # items_div = dict(zip(values_to_add_rp, values_to_add_div))
        # items_stock = dict(zip(values_to_add_rp, values_to_add_stock))


        # проверка есть ли дивизион в items_div по колонке 'Завод пользователя', 
        # check_value_in_list_and_set_value устанавливает 'Пустой дивизион' в 'Дивизион' если не нашел и дивизион если нашел
        # так можно проверить любую колонку на значения загруженные из настроек (новые настройки можно создавать)
        # row - исходная таблица
        # row_name - название столбца
        # items_list - список столбцов для которых нужно сделать новые значения
        # default_value - значение если в списке не найдей ключ
        # Sheet1['Дивизион'] = Sheet1.apply(utils.check_value_in_list_and_set_value, axis=1, row_name='Завод пользователя', items_list=items_div, default_value='Пустой дивизион')







    # !!!! ТУТ идет сохранение полученного результата, чтобы отправить обратно в приложение и скачать
    # 'PYTHON_SAVED_FILES_PATH' папка для сохранения обработанных скриптом и сохраненных файлов с возможностью выгрузить их из приложения
    # функция utils.createEnvPath создает правильный путь для сохранения (до места где хранится обработанные скриптом файлы + второй аргумент это название файла)
    # если в процессе выполнения вы решили создать еще таблиц (например Sheet2 = pd.DataFrame()) и сохранить/показать уже ее
    # то нужно вызвать to_excel() и to_html() уже у новой таблицы (например Sheet2.to_excel(output_file, index=False)/Sheet2.to_html(output_file_html, index=False))
    output_file = utils.createEnvPath('PYTHON_SAVED_FILES_PATH', save_file_name)
    output_file_html = os.path.splitext(output_file)[0] + '.html'
    Sheet1.to_excel(output_file, index=False)

    # тут можно накинуть стилей в уже сохраненные листы файла, сохранить нужные листы и тд (примеры ниже)
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
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

