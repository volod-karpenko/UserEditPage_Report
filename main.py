import os
import pyodbc
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime, timedelta
import sys
import json

load_dotenv(override=True)
PATH_TO_REPORTS_FOLDER = os.path.join(os.getcwd(), 'reports')

def get_date_params():
    current_date = datetime.today()
    one_day_delta = timedelta(days=1)
    seconds_delta = timedelta(seconds=24 * 3600 - 1)
    date_from, date_to, *rest = sys.argv[1:]
    try:
        date_from = datetime.strptime(date_from, '%Y-%m-%d')
    except:
        date_from = datetime(year=current_date.year, month=current_date.month - 1, day=1)
    try:
        date_to = datetime.strptime(date_to, '%Y-%m-%d')
    except:
        date_to = datetime(year=current_date.year, month=current_date.month, day=1) - one_day_delta

    return date_from, date_to + seconds_delta

def get_database_url():
    DATABASE_URL = os.getenv('DATABASE_URL')
    driver = pyodbc.drivers()[-1]
    return f"{DATABASE_URL};DRIVER={{{driver}}};"

def get_columns(columns):
    COLUMNS = {
        'Road': 'Дорога',
        'MainTable': 'Таблиця',
        'Entity': 'Окрема облікова картка',
        'Operation': 'Дія',
        'EditAt': 'Дата та час',
        'CntObjects': 'К-сть записів',
        'User': 'Користувач'
    }
    return [COLUMNS.get(column, column) for column in columns]

def get_data(date_from, date_to):
    query = 'EXEC [dbo].[sp_UserEditPage_Report] ?, ?'
    driver = get_database_url()
    cnxn = pyodbc.connect(driver)
    cursor = cnxn.cursor()
    cursor.execute(query, [date_from.isoformat(sep=' '), date_to.isoformat(sep=' ')])
    rows = cursor.fetchall()
    columns = get_columns([column[0] for column in cursor.description])
    cursor.close()
    cnxn.close()
    return pd.DataFrame.from_records(data=rows, columns=columns)

def get_file_name(date_from, date_to):
    date_from_text = date_from.isoformat().replace('-', '_')[:10]
    date_to_text = date_to.isoformat().replace('-', '_')[:10]
    timestamp = datetime.now().timestamp() * 1000
    return f'ReportFrom{date_from_text}To{date_to_text}__{timestamp:.0f}.xlsx'

if __name__ == '__main__':
    file_name = None
    error = None
    try:
        date_from, date_to = get_date_params()
        data = get_data(date_from=date_from, date_to=date_to)
        file_name = get_file_name(date_from=date_from, date_to=date_to)
        file_name = os.path.join(PATH_TO_REPORTS_FOLDER, file_name)
        data.to_excel(file_name, sheet_name='Дані', index=False)
    except Exception as e:
        file_name = None
        error = e
    print(json.dumps({'file_name': file_name, 'error': error}))