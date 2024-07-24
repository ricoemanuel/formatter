import pandas as pd
from io import BytesIO
import base64
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, numbers

class ExcelFormatter:
    #initialize class attributes
    def __init__(self, contentBytes):
        self.contentBytes = contentBytes
        self.df = None
        self.grouped_data = None
        self.wb = Workbook()
        self.ws = self.wb.active

    #read excel file and convert it into a pandas data frame
    def decode_content(self):
        decoded = base64.b64decode(self.contentBytes)
        excel = BytesIO(decoded)
        self.df = pd.read_excel(excel, skiprows=5, skipfooter=4, engine='openpyxl')

    #apply conditions into the new data frame
    def filter_and_group_data(self):
        self.df = self.df[self.df['Live Check Amount'] > 0]
        self.grouped_data = self.df.groupby('Client').agg(
            number_of_live_checks=pd.NamedAgg(column='Live Check Amount', aggfunc='count'),
            check_totals=pd.NamedAgg(column='Live Check Amount', aggfunc='sum')
        ).reset_index()

    #add grand totals row
    def add_totals_row(self):
        self.grouped_data.loc[len(self.grouped_data)] = {
            'Client': 'Totals',
            'number_of_live_checks': self.grouped_data['number_of_live_checks'].sum(),
            'check_totals': self.grouped_data['check_totals'].sum()
        }

    def format_worksheet(self):
        #convert dataframe to a worksheet
        for r in dataframe_to_rows(self.grouped_data, index=False, header=True):
            self.ws.append(r)
        #initialize main color for headers and footers
        blue_fill = PatternFill(start_color="94DCF8", end_color="94DCF8", fill_type="solid")
        #apply color
        for cell in self.ws[1]:
            cell.fill = blue_fill
        #ajust cell size according to the lenght of the content
        for column in self.ws.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            self.ws.column_dimensions[column[0].column_letter].width = adjusted_width
        #give 'c' column USD format
        for cell in self.ws['C']:
            cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        # turn footer into main color
        last_row = self.ws[len(self.ws['A'])]
        for cell in last_row:
            cell.fill = blue_fill
    #export binary content of the file
    def save_workbook(self):
        output = BytesIO()
        self.wb.save(output)
        output.seek(0)
        return output
#main controller of the endpoint
def formatExcel(contentBytes):
    formatter = ExcelFormatter(contentBytes)
    formatter.decode_content()
    formatter.filter_and_group_data()
    formatter.add_totals_row()
    formatter.format_worksheet()
    return formatter.save_workbook()
