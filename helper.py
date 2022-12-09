import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv()

XLSXFORMS = os.getenv('XLSX_FORMS_FOLDER')

style = """
    <style>
    td > div > table > tbody > tr > td {
        color: #264985;
        font-size: 0.9rem;
    }
    </style>
    """

NON_SURVEY_WORKSHEETS = ['choices', 'calculates', 'queries', 'settings', 'model', 'prompt_types', 'initial', 'End']
TYPE_2_ESCAPE = ['nan', 'user_branch', 'note', 'linked_table', 'finalize']

def get_xlsx_files():
    forms_location = os.path.join(os.getcwd(), XLSXFORMS)
    os.chdir(forms_location)
    xlsx = []
    for file in os.listdir():
        if file.endswith(".xlsx") or file.endswith(".xls"):
            xlsx.append(os.path.join(forms_location, file))
    xlsx.sort()
    return xlsx

def format_relevant(value):
    if str(value) == 'nan':
        return '-'
    return value

def format_question(value):
    if str(value) == 'nan':
        return 'Hidden from user'
    return value

def get_value(type, choice_df='', choice_name = ''):
    if type in ['select_one', 'select_multiple', 'select_one_integer']:
        filtered_choice_df = choice_df[choice_df["choice_list_name"]==choice_name]
        values = []
        for i, r in filtered_choice_df.iterrows():
            values.append([r["data_value"], '-', r["display.text"]])
        return values
    if type in ['integer', 'decimal']:
        return 'User entered number'
    if type == 'text':
        return 'User entered text'
    if type in ['string', 'select_one_integer', 'async_assign_string', 'assign', 'async_assign_count', 'async_assign_position']:
        return 'Calculate field'
    if type in ['date', 'eth_date']:
        return 'User entered date'
    if type == 'time':
        return 'User entered time'
    if type == 'dateTime':
        return 'User entered date time'
    if type == 'image':
        return 'Captured image'
    if type == 'audio':
        return 'Recorded audio'
    if type == 'video':
        return 'Recorded video'
    if type == 'barcode':
        return 'User scanned barcode'
    if type == 'geopoint':
        return 'Geographic coordinate'

    return "-"