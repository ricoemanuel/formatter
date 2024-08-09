import pandas as pd
from io import BytesIO
import base64
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, numbers
import numpy as np
import pandas as pd
from datetime import datetime
from io import StringIO

def decode_content(contentBytes):
    decoded = base64.b64decode(contentBytes)
    excel = BytesIO(decoded)
    df = pd.read_excel(excel, skiprows=5, skipfooter=4, engine='openpyxl')
    return df

def filter_and_group_data(df):
    df = df[df['Live Check Amount'] > 0]
    grouped_data = df.groupby('Client').agg(
        number_of_live_checks=pd.NamedAgg(column='Live Check Amount', aggfunc='count'),
        check_totals=pd.NamedAgg(column='Live Check Amount', aggfunc='sum')
    ).reset_index()
    return grouped_data

def add_totals_row(grouped_data):
    grouped_data.loc[len(grouped_data)] = {
        'Client': 'Totals',
        'number_of_live_checks': grouped_data['number_of_live_checks'].sum(),
        'check_totals': grouped_data['check_totals'].sum()
    }
    return grouped_data

def format_worksheet(grouped_data):
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(grouped_data, index=False, header=True):
        ws.append(r)
    blue_fill = PatternFill(start_color="94DCF8", end_color="94DCF8", fill_type="solid")
    for cell in ws[1]:
        cell.fill = blue_fill
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width
    for cell in ws['C']:
        cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    last_row = ws[len(ws['A'])]
    for cell in last_row:
        cell.fill = blue_fill
    return wb

def save_workbook(wb):
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def formatExcel(contentBytes):
    df = decode_content(contentBytes)
    grouped_data = filter_and_group_data(df)
    grouped_data = add_totals_row(grouped_data)
    wb = format_worksheet(grouped_data)
    return save_workbook(wb)

def formatFromJson(content):
    df= pd.DataFrame(content)
    df['Live Check Amount'] = df['Live Check Amount'].replace('', np.nan)

    df['Live Check Amount'] = df['Live Check Amount'].fillna(0)

    df['Live Check Amount'] = df['Live Check Amount'].astype(float)
    grouped_data = filter_and_group_data(df)
    grouped_data = add_totals_row(grouped_data)
    return grouped_data.to_dict(orient='records')

def createFile():

    date = datetime.now()
    quarter = (date.month-1)//3 + 1

    output_directory = fr"{date.year}\Q{quarter}"
    output_file = fr"SNIC Monthly Payroll_{date.strftime('%B')} {date.year}.xlsx"


    return {"file_name":output_file,"path":output_directory}
