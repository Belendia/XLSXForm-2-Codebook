# import pandas lib as pd
import os
import pandas as pd
import re
from tabulate import tabulate
import html
import minify_html
# from helper import survey_df, get_value, format_relevant, format_question, style
from helper import get_xlsx_files, NON_SURVEY_WORKSHEETS, TYPE_2_ESCAPE, get_value, format_relevant, format_question, style

CODEBOOK_FILE_NAME = os.getenv('CODEBOOK_FILE_NAME')
table = []


xlsx_files = get_xlsx_files()

for xlsx_file in xlsx_files:
    fields_included = []
    xls = pd.ExcelFile(xlsx_file)
    choice_df = pd.read_excel(xls, sheet_name="choices")
    table.append(['<h2>{}</h2>'.format(xlsx_file)])
    for book in xls.book:
        if book.title not in NON_SURVEY_WORKSHEETS:
            survey_df = pd.read_excel(xls, sheet_name=book.title)
            for index, row in survey_df.iterrows():
                question_type = str(row['type'])
                is_session_variable = True if 'model.isSessionVariable'in row and row['model.isSessionVariable']==1 else False
                if question_type not in TYPE_2_ESCAPE and is_session_variable == False and row['name'] not in fields_included:
                    fields_included.append(row['name'])
                    table.append(["<b style='font-size: 1rem;'>{}</b>".format(format_question(row["display.text"]))])
                    table.append(['<div style="padding-left:20px; font-size: 0.9rem; color: #595959;">Name: <b>{}</b></div>'.format(row["name"])])
                    table.append(['<div style="padding-left:20px; font-size: 0.9rem; color: #595959;">Question type: <b>{}</b></div>'.format(question_type)])
                    table.append(['<div style="padding-left:20px; font-size: 0.9rem; color: #595959;">Relevant: {}</div>'.format(format_relevant(row["required"]))])

                    if question_type in ['select_one', 'select_multiple', 'select_one_integer']:
                        values = get_value(question_type, choice_df, str(row['values_list']))
                        table.append(['<div style="padding-left:20px; font-size: 0.9rem; color: #595959;">Values: {}</div>'.format(tabulate(values, tablefmt="html"))])
                    else:
                        table.append(['<div style="padding-left:20px; font-size: 0.9rem; color: #595959;">Values: {}</div>'.format(get_value(question_type))])

                    table.append([""])


with open(CODEBOOK_FILE_NAME, "w") as f:
    final_html = "{}{}".format(style, tabulate(table, tablefmt='html'))
    f.write(minify_html.minify(html.unescape(final_html)))