import pandas as pd

# Создать строчки с итоговыми значениями под каждой группой
# table - исходная таблица
# group_name - название столбца с группами
# total_columns - список столбцов для которых нужно сделать итог
def add_total_by_field(table, group_name, total_columns):
    result_table = pd.DataFrame(columns=table.columns)

    prev_group = table.iloc[0][group_name]
    for i in range(len(table)):
        current_group = table.iloc[i][group_name]
        if current_group != prev_group:
            row_dict = {}
            row_dict[group_name] = [prev_group]
            for column in total_columns:
                column_sum = table.loc[table[group_name] == prev_group][column].astype(float).sum()
                row_dict[column] = [column_sum]

            row = pd.DataFrame(row_dict)
            result_table = pd.concat([result_table, row])
            result_table = pd.concat([result_table, table.iloc[i:i+1]])
            prev_group = current_group
        else:
            result_table = pd.concat([result_table, table.iloc[i:i+1]])

    last_row_dict = {}
    last_row_dict[group_name] = [prev_group]
    for column in total_columns:
        column_sum = table.loc[table[group_name] == prev_group][column].astype(float).sum()
        last_row_dict[column] = [column_sum]

    last_row = pd.DataFrame(last_row_dict)
    result_table = pd.concat([result_table, last_row])

    return result_table