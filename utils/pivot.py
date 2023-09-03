import pandas as pd

# utils.create_pivot_table_and_get_div_list - создать сводную типа нарушение да нет процент итого и возвращает ее и список дивизионов в итоговой таблице
# data_table - исходная таблица с данными
# violation_col - столбц с данными по нарушениями
# div_col_name - название столбца с дивизионами
# rp_col_name - название столбца с заводами
# values_col_name - название столбца значений
def create_pivot_table_and_get_div_list(data_table, violation_col, div_col_name, rp_col_name, values_col_name):
    temporary_table = pd.DataFrame()
    temporary_table = pd.pivot_table(data_table,
            index=[div_col_name, rp_col_name],
            columns=violation_col,
            values=values_col_name,
            aggfunc='count')

    temporary_table['да'] = temporary_table['да'].fillna(0)
    temporary_table['нет'] = temporary_table['нет'].fillna(0)

    temporary_table = temporary_table.reset_index()

    result_table = pd.DataFrame(columns=temporary_table.columns)
    new_rows = []

    previous_division = temporary_table.iloc[0][div_col_name]

    division_list = []
    division_list.append(previous_division)

    last = None
    previous_sum1 = 0
    previous_sum2 = 0

    for index, row in temporary_table.iterrows():
        current_division = row[div_col_name]
        last = row[div_col_name]

        if current_division != previous_division:
            new_row = {div_col_name: None, 
                    rp_col_name: previous_division, 
                    'да': previous_sum1,
                    'нет': previous_sum2 }
            new_rows.append(new_row)
            division_list.append(current_division)
            
            previous_division = current_division
            previous_sum1 = 0
            previous_sum2 = 0
            previous_sum1 += int(row['да'])
            previous_sum2 += int(row['нет'])
        else:
            previous_sum1 += int(row['да'])
            previous_sum2 += int(row['нет'])

    last_row_dict = {div_col_name: None, 
                rp_col_name: last, 
                'да': previous_sum1,
                'нет': previous_sum2 }
    new_rows.append(last_row_dict)
    division_list.append(last)

    new_rows_index = 0

    result_table.loc[len(result_table)] = new_rows[new_rows_index]
    new_rows_index = new_rows_index + 1
    previous_division = temporary_table.iloc[0][div_col_name]

    for index, row in temporary_table.iterrows():
        current_division = row[div_col_name]

        if current_division != previous_division:
            result_table.loc[len(result_table)] = new_rows[new_rows_index]
            previous_division = current_division
            new_rows_index = new_rows_index + 1
            result_table.loc[len(result_table)] = row
        else:
            result_table.loc[len(result_table)] = row

    total_fact = temporary_table['да'].sum()
    total_cost = temporary_table['нет'].sum()
    total_count = total_fact + total_cost
    total_row = pd.DataFrame({rp_col_name: ['Общий итог'],
                            div_col_name: [''],
                            'да': [total_fact],
                            'нет': [total_cost],
                            'Общий итог': [total_count]})
    result_table = pd.concat([result_table, total_row])

    sum_column = result_table['да'] + result_table['нет']
    result_table['Общий итог'] = sum_column

    percentage_column = (result_table['нет'] / result_table['Общий итог'])
    result_table['Процент %'] = percentage_column.round(4)

    result_table['да'] = result_table['да'].apply(lambda x: round(x)).astype(int)
    result_table['нет'] = result_table['нет'].apply(lambda x: round(x)).astype(int)
    result_table['Общий итог'] = result_table['Общий итог'].apply(lambda x: round(x)).astype(int)

    division_list.append('Общий итог')

    result_table = result_table.drop([div_col_name], axis=1)

    res_dict = {'table': result_table, 'div': division_list}
    return res_dict


# utils.create_pivot_table_and_get_div_list - создать сводную типа нарушение да нет процент итого и возвращает ее и список дивизионов в итоговой таблице
# data_table - исходная таблица с данными
# violation_col - столбц с данными по нарушениями
# div_col_name - название столбца с дивизионами
# rp_col_name - название столбца с заводами
# values_col_name - название столбца значений
def create_pivot_table_and_get_div_list_2(data_table, violation_col, div_col_name, rp_col_name, values_col_name):
    temporary_table = pd.DataFrame()
    temporary_table = pd.pivot_table(data_table,
            index=[div_col_name, rp_col_name],
            columns=violation_col,
            values=values_col_name,
            aggfunc='count')

    temporary_table['1.0-4 часа'] = temporary_table['1.0-4 часа'].fillna(0)
    temporary_table['2.4-12 часов'] = temporary_table['2.4-12 часов'].fillna(0)
    temporary_table['3.Более 12 часов'] = temporary_table['3.Более 12 часов'].fillna(0)
    temporary_table['4.Не завершена'] = temporary_table['4.Не завершена'].fillna(0)

    temporary_table = temporary_table.reset_index()


    result_table = pd.DataFrame(columns=temporary_table.columns)
    new_rows = []

    previous_division = temporary_table.iloc[0][div_col_name]

    division_list = []
    division_list.append(previous_division)

    last = None
    previous_sum1 = 0
    previous_sum2 = 0
    previous_sum3 = 0
    previous_sum4 = 0

    for index, row in temporary_table.iterrows():
        current_division = row[div_col_name]
        last = row[div_col_name]

        if current_division != previous_division:
            new_row = {div_col_name: None, 
                    rp_col_name: previous_division, 
                    '1.0-4 часа': previous_sum1,
                    '2.4-12 часов': previous_sum2,
                    '3.Более 12 часов': previous_sum3,
                    '4.Не завершена': previous_sum4,
                    }
            new_rows.append(new_row)
            division_list.append(current_division)
            
            previous_division = current_division
            previous_sum1 = 0
            previous_sum2 = 0
            previous_sum3 = 0
            previous_sum4 = 0
            previous_sum1 += int(row['1.0-4 часа'])
            previous_sum2 += int(row['2.4-12 часов'])
            previous_sum3 += int(row['3.Более 12 часов'])
            previous_sum4 += int(row['4.Не завершена'])
        else:
            previous_sum1 += int(row['1.0-4 часа'])
            previous_sum2 += int(row['2.4-12 часов'])
            previous_sum3 += int(row['3.Более 12 часов'])
            previous_sum4 += int(row['4.Не завершена'])

    last_row_dict = {div_col_name: None, 
                rp_col_name: last,
                '1.0-4 часа': previous_sum1,
                '2.4-12 часов': previous_sum2,
                '3.Более 12 часов': previous_sum3,
                '4.Не завершена': previous_sum4,
                }
    new_rows.append(last_row_dict)
    division_list.append(last)

    new_rows_index = 0

    result_table.loc[len(result_table)] = new_rows[new_rows_index]
    new_rows_index = new_rows_index + 1
    previous_division = temporary_table.iloc[0][div_col_name]

    for index, row in temporary_table.iterrows():
        current_division = row[div_col_name]

        if current_division != previous_division:
            result_table.loc[len(result_table)] = new_rows[new_rows_index]
            previous_division = current_division
            new_rows_index = new_rows_index + 1
            result_table.loc[len(result_table)] = row
        else:
            result_table.loc[len(result_table)] = row

    total_1 = temporary_table['1.0-4 часа'].sum()
    total_2 = temporary_table['2.4-12 часов'].sum()
    total_3 = temporary_table['3.Более 12 часов'].sum()
    total_4 = temporary_table['4.Не завершена'].sum()
    total_count = total_1 + total_2 + total_3 + total_4
    total_row = pd.DataFrame({rp_col_name: ['Общий итог'],
                            div_col_name: [''],
                            '1.0-4 часа': [total_1],
                            '2.4-12 часов': [total_2],
                            '3.Более 12 часов': [total_3],
                            '4.Не завершена': [total_4],
                            'Общий итог': [total_count]})
    result_table = pd.concat([result_table, total_row])

    sum_column = result_table['1.0-4 часа'] + result_table['2.4-12 часов'] + result_table['3.Более 12 часов'] + result_table['4.Не завершена']
    result_table['Общий итог'] = sum_column

    percentage_column1 = (result_table['1.0-4 часа'] / result_table['Общий итог'])
    result_table.insert(3, '1.0-4 часа %', percentage_column1.round(4))

    percentage_column2 = (result_table['2.4-12 часов'] / result_table['Общий итог'])
    result_table.insert(5, '2.4-12 часов %', percentage_column2.round(4))

    percentage_column3 = (result_table['3.Более 12 часов'] / result_table['Общий итог'])
    result_table.insert(7, '3.Более 12 часов %', percentage_column3.round(4))

    percentage_column4 = (result_table['4.Не завершена'] / result_table['Общий итог'])
    result_table.insert(9, '4.Не завершена %', percentage_column4.round(4))

    result_table['1.0-4 часа'] = result_table['1.0-4 часа'].apply(lambda x: round(x)).astype(int)
    result_table['2.4-12 часов'] = result_table['2.4-12 часов'].apply(lambda x: round(x)).astype(int)
    result_table['3.Более 12 часов'] = result_table['3.Более 12 часов'].apply(lambda x: round(x)).astype(int)
    result_table['4.Не завершена'] = result_table['4.Не завершена'].apply(lambda x: round(x)).astype(int)
    result_table['Общий итог'] = result_table['Общий итог'].apply(lambda x: round(x)).astype(int)

    division_list.append('Общий итог')

    result_table = result_table.drop([div_col_name], axis=1)

    res_dict = {'table': result_table, 'div': division_list}
    return res_dict
