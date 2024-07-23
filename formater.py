import pandas as pd
from io import BytesIO
import base64
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill
from openpyxl.styles import numbers

def formatExcel(contentBytes):
    decoded = base64.b64decode(contentBytes)
    excel = BytesIO(decoded)
    
    df = pd.read_excel(excel, skiprows=5, skipfooter=4, engine='openpyxl')
    

    df = df[df['Live Check Amount'] > 0]
    grouped_data = df.groupby('Client').agg(
        number_of_live_checks=pd.NamedAgg(column='Live Check Amount', aggfunc='count'),
        check_totals=pd.NamedAgg(column='Live Check Amount', aggfunc='sum')
    ).reset_index()

    grouped_data.loc[len(grouped_data)]={
        'Client': 'Totals',
        'number_of_live_checks': grouped_data['number_of_live_checks'].sum(),
        'check_totals': grouped_data['check_totals'].sum()
    }
    

    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    
    for r in dataframe_to_rows(grouped_data, index=False, header=True):
        ws.append(r)

    blue_fill = PatternFill(start_color="94DCF8",
                            end_color="94DCF8",
                            fill_type="solid")
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

    wb.save(output)
    output.seek(0)

    return output
