import sys
import os
import json
import pandas as pd
import numpy as np
from dotenv import load_dotenv
load_dotenv()

def createEnvPath(env_path, last = None):
    if os.getenv('MODE') == 'production':
        if last:
            return os.path.join(os.getcwd(), 'dist', os.getenv(env_path), last)
        return os.path.join(os.getcwd(), 'dist', os.getenv(env_path))
    else:
        if last:
            return os.path.join(os.getcwd(), os.getenv(env_path), last)
    return os.path.join(os.getcwd(), os.getenv(env_path))

def sum_all(row):
    if row['name']:
        return 'city'
    else:
        return row['name']
        
if len(sys.argv) < 2:
    print('No file given')
else:
    excel_file = createEnvPath('SAVED_FILES_PATH', sys.argv[1])
    data = pd.read_excel(excel_file, sheet_name='Sheet1')
    # pivot = pd.pivot_table(
    #     data,
    #     values=['summ'],
    #     index=['division', 'name'], 
    #     aggfunc='sum',
    #     columns='periods'
    # )
    # pivot.loc[('Total', 'Total'), :] = pivot.sum(axis=0)
    # pivot.loc[('Total', 'Total'), 'summ'] = 'Total'
    # df = pivot.apply(sum_all. axis=1)

    output_file_excel = createEnvPath('PYTHON_SAVED_FILES_PATH', sys.argv[1])
    output_file_html = os.path.splitext(output_file_excel)[0] + '.html'

    data.to_excel(output_file_excel, index=False)
    data.to_html(output_file_html, index=False)
    print(True)

