import math
import pandas as pd
from io import BytesIO
import base64
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, numbers
import numpy as np
import warnings

def custom_warning_handler(message, category, filename, lineno, file=None, line=None):
    if "Cannot parse header or footer" in str(message):
        return
    warnings.showwarning(message, category, filename, lineno, file, line)

warnings.showwarning = custom_warning_handler

class ExcelDecoder:
    @staticmethod
    def decode_content(contentBytes, skiprows=0, skipfooter=0):
        decoded = base64.b64decode(contentBytes)
        excel = BytesIO(decoded)
        df = pd.read_excel(excel, skiprows=skiprows, skipfooter=skipfooter, engine='openpyxl')
        return df

class DataProcessor:
    @staticmethod
    def filter_and_group_data(df):
        df = df[df['Live Check Amount'] > 0]
        grouped_data = df.groupby('Client').agg(
            number_of_live_checks=pd.NamedAgg(column='Live Check Amount', aggfunc='count'),
            check_totals=pd.NamedAgg(column='Live Check Amount', aggfunc='sum')
        ).reset_index()
        return grouped_data

    @staticmethod
    def add_totals_row(grouped_data):
        grouped_data.loc[len(grouped_data)] = {
            'Client': 'Totals',
            'number_of_live_checks': grouped_data['number_of_live_checks'].sum(),
            'check_totals': grouped_data['check_totals'].sum()
        }
        return grouped_data

class ExcelFormatter:
    @staticmethod
    def format_worksheet(grouped_data):
        wb = Workbook()
        ws = wb.active
        for r in dataframe_to_rows(grouped_data, index=False, header=True):
            ws.append(r)
        ExcelFormatter._apply_styles(ws)
        return wb

    @staticmethod
    def _apply_styles(ws):
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

class ExcelSaver:
    @staticmethod
    def save_workbook(wb):
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

def formatExcel(contentBytes):
    df = ExcelDecoder.decode_content(contentBytes, skiprows=5, skipfooter=4)
    grouped_data = DataProcessor.filter_and_group_data(df)
    grouped_data = DataProcessor.add_totals_row(grouped_data)
    wb = ExcelFormatter.format_worksheet(grouped_data)
    return ExcelSaver.save_workbook(wb)

def formatFromJson(content):
    df = pd.DataFrame(content)
    df['Live Check Amount'] = df['Live Check Amount'].replace('', np.nan).fillna(0).astype(float)
    grouped_data = DataProcessor.filter_and_group_data(df)
    grouped_data = DataProcessor.add_totals_row(grouped_data)
    return grouped_data.to_dict(orient='records')

def discrepancies_report_ssn(contentBytes, path):
    df = ExcelDecoder.decode_content(contentBytes)
    if "aetna" in path.lower():
        dfs = find_tables_in_excel(df)
       
        for df in dfs:
            ssn_columns = [col for col in df.columns if isinstance(col, str) and pd.notna(col) and 'ssn' in col.lower()]
            for ssn_column in ssn_columns:
                df[ssn_column] = df[ssn_column].apply(remove_leading_zero)
        columns_to_keep = {
        'EE SSN': 'SSN',
        'SSN': 'SSN',
        'Comments': 'Comments',
        'comments': 'Comments',
        'Notes': 'Notes',
        'notes': 'Notes',
        'Dep SSN':'Dep SSN'
        }
        combined_df = pd.DataFrame()
        
        for df in dfs:
            filtered_df = df[[col for col in df.columns if col in columns_to_keep]].rename(columns=columns_to_keep)
            if 'EE SSN' in filtered_df.columns and 'SSN' in filtered_df.columns:
                filtered_df['SSN'] = filtered_df['EE SSN'].combine_first(filtered_df['SSN'])
                filtered_df = filtered_df.drop(columns=['EE SSN'])
            filtered_df['Carrier'] = 'aetna'
            filtered_df['PEO_ID'] = ''
            combined_df = pd.concat([combined_df, filtered_df], ignore_index=True)
        ssn_col = next((col for col in combined_df.columns if col.lower() in ['ssn', 'full ssn', 'ee ssn']), None)
        if ssn_col:
            ssn_records = combined_df[ssn_col].tolist()
            ssn_records = [int(ssn) if isinstance(ssn, str) and ssn.isdigit() else ssn for ssn in ssn_records]
        return {"ssn":ssn_records}

    elif "legal shield" in path.lower():
        ssn_col = next((col for col in df.columns if col.lower() in ['ssn', 'full ssn', 'ee ssn']), None)
        if ssn_col:
            ssn_records = df[ssn_col].tolist()
            ssn_records = [int(ssn) if isinstance(ssn, str) and ssn.isdigit() else ssn for ssn in ssn_records]
        return {"ssn":ssn_records}
    elif "empire" in path.lower():
        df.columns = [col.strip() for col in df.columns]
        ssn_col = next((col for col in df.columns if col.lower() in ['ssn', 'full ssn', 'ee ssn']), None)
        
        if ssn_col:
            print(ssn_col)
            ssn_records = df[ssn_col].tolist()
            ssn_records = [int(ssn) if isinstance(ssn, str) and ssn.isdigit() else ssn for ssn in ssn_records]
        return {"ssn":ssn_records}

    
def discrepancies_report(contentBytes, path, planTermDetails, termDates):
    df = ExcelDecoder.decode_content(contentBytes)
    if "aetna" in path.lower():
        dfs = find_tables_in_excel(df)
       
        for df in dfs:
            ssn_columns = [col for col in df.columns if isinstance(col, str) and pd.notna(col) and 'ssn' in col.lower()]
            for ssn_column in ssn_columns:
                df[ssn_column] = df[ssn_column].apply(remove_leading_zero)
        columns_to_keep = {
        'EE SSN': 'SSN',
        'SSN': 'SSN',
        'Comments': 'Comments',
        'comments': 'Comments',
        'Notes': 'Notes',
        'notes': 'Notes',
        'Dep SSN':'Dep SSN'
        }
        combined_df = pd.DataFrame()
        
        for df in dfs:
            filtered_df = df[[col for col in df.columns if col in columns_to_keep]].rename(columns=columns_to_keep)
            if 'EE SSN' in filtered_df.columns and 'SSN' in filtered_df.columns:
                filtered_df['SSN'] = filtered_df['EE SSN'].combine_first(filtered_df['SSN'])
                filtered_df = filtered_df.drop(columns=['EE SSN'])
            filtered_df['Carrier'] = 'aetna'
            filtered_df['PEO_ID'] = ''
            combined_df = pd.concat([combined_df, filtered_df], ignore_index=True)
        combined_df=find_requirement_aetna(combined_df,planTermDetails,termDates)
        return save_tables_to_excel([combined_df])

    elif "legal shield" in path.lower():
        df=find_requirement_legalShield(df,planTermDetails,termDates)
        return save_tables_to_excel([df])
    elif "empire" in path.lower():
        df=find_requirement_empire(df,planTermDetails,termDates)
        return save_tables_to_excel([df])

def remove_leading_zero(ssn):
    if pd.notna(ssn):
        if int(ssn) > 9:
            return ssn.lstrip('0')
    return ssn

def find_tables_in_excel(df):
    tables = []
    table_start = None
    for i, row in df.iterrows():
        row_lower = row.str.lower()
        if any(row_lower.str.contains('csa', na=False)) and (any(row_lower.str.contains('name', na=False)) or any(row_lower.str.contains('ee name', na=False))):
            table_start = i
        if table_start is not None and row.isnull().all():
            table = df.iloc[table_start:i].reset_index(drop=True)
            table.columns = table.iloc[0]
            table = table.drop(0).reset_index(drop=True)
            tables.append(table)
            table_start = None
    return tables

def save_tables_to_excel(tables):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for idx, table in enumerate(tables):
            sheet_name = f"Table_{idx + 1}"
            table.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output.getvalue()

def find_requirement_aetna(df,carrierPlanDetails,carrierTermDates):
    discrepancies = pd.read_excel("DISCREPANCIES.xlsx")
    
    df['Found Data'] = ''
    
    for index, item in df.iterrows():
        
        comment = item['Comments']
        found_keywords = []
        for _, keyword_row in discrepancies.iterrows():
            keyword = str(keyword_row['Key word'])
            if keyword.lower() in comment.lower():
                found_keywords.append(keyword_row.to_dict())
        found_keywords = {item['Data Base']: item for item in found_keywords}.values()
        found_keywords = list(found_keywords)
        
        if len(found_keywords) > 0:
            key_word = found_keywords[0]
            
            item_ssn = str(item["SSN"])
            if pd.notna(key_word["Data Base"]):
                df.at[index, 'key word'] = key_word["Data Base"]
                if key_word["Data Base"] != "TERMDATE":
                    carrierPlanDetails['SSN'] = carrierPlanDetails['SSN'].astype(str)
                    resultado = carrierPlanDetails[carrierPlanDetails['SSN'] == item_ssn]
                    if not resultado.empty:
                        dep_ssn = str(item["Dep SSN"])
                        resultadodep = resultado[resultado['DEP_SSN'] == dep_ssn]
                        if not resultadodep.empty:
                            datos = resultadodep[key_word["Data Base"]].values
                            datos = list(set(resultadodep[key_word["Data Base"]].values))
                        else:
                            datos = resultado[key_word["Data Base"]].values 
                        datos_filtrados = list({dato for dato in datos if '/' not in str(dato)} | {dato for dato in datos if '/' in str(dato)})
                        datos_joined = ';'.join(map(str, datos_filtrados))
                        df.at[index, 'Found Data'] = datos_joined
                    else:
                        df.at[index, 'Found Data'] = ''
                else:
                    carrierTermDates['SSN'] = carrierTermDates['SSN'].astype(str)
                    resultado = carrierTermDates[carrierTermDates['SSN'] == item_ssn]
                    if not resultado.empty:
                        datos = resultado[key_word["Data Base"]].values 
                        datos = [f"{date.astype('datetime64[D]').item().month}/{date.astype('datetime64[D]').item().day}/{date.astype('datetime64[D]').item().year}" for date in datos]
                        datos_filtrados = list({dato for dato in datos if '/' not in str(dato)} | {dato for dato in datos if '/' in str(dato)})
                        datos_joined = ';'.join(map(str, datos_filtrados))
                        df.at[index, 'Found Data'] = datos_joined
                    else:
                        df.at[index, 'Found Data'] = ''
            else:
                df.at[index, 'key word'] = 'Invalid field'
                df.at[index, 'Found Data'] = ''
        else:
            df.at[index, 'Found Data'] = ''
    
    return df

def find_requirement_legalShield(df,carrierPlanDetails,carrierTermDates):

    for index, item in df.iterrows():
        item_ssn = str(item["FULL SSN"])
        carrierPlanDetails['SSN'] = carrierPlanDetails['SSN'].astype(str)
        carrierTermDates['SSN'] = carrierTermDates['SSN'].astype(str)
        resultado = carrierPlanDetails[carrierPlanDetails['SSN'] == item_ssn]
        field='COVERAGE_END_DATE'
        datos=resultado[field].values
        
        datos_filtrados = datos_filtrados = list(
            {dato for dato in datos if '/' not in str(dato) and not (isinstance(dato, float) and math.isnan(dato))} |
            {dato for dato in datos if '/' in str(dato) and not (isinstance(dato, float) and math.isnan(dato))}
        )
        
        if len(datos_filtrados)==0:
            field='TERMDATE'
            resultado=carrierTermDates[carrierTermDates['SSN'] == item_ssn]
            datos=resultado[field].values
            datos = [
                f"{date.astype('datetime64[D]').item().month}/{date.astype('datetime64[D]').item().day}/{date.astype('datetime64[D]').item().year}"
                if isinstance(date, np.datetime64) else date
                for date in datos if not pd.isna(date)
            ]
            datos_filtrados = list({dato for dato in datos if '/' not in str(dato)} | {dato for dato in datos if '/' in str(dato)})
        datos_joined = ';'.join(map(str, datos_filtrados))
        df.at[index, field] = datos_joined
    
    return df


def find_requirement_empire(df,carrierPlanDetails,carrierTermDates):
    discrepancies = pd.read_excel("DISCREPANCIES.xlsx")
    
    df['Found Data'] = ''
    
    for index, item in df.iterrows():
        
        comment = item['HOW TO RESOLVE  (ERROR DESCRIPTION)']
        found_keywords = []
        for _, keyword_row in discrepancies.iterrows():
            keyword = str(keyword_row['Key word'])
            if keyword.lower() in comment.lower():
                found_keywords.append(keyword_row.to_dict())
        found_keywords = {item['Data Base']: item for item in found_keywords}.values()
        found_keywords = list(found_keywords)
        
        if len(found_keywords) > 0:
            key_word = found_keywords[0]
            
            item_ssn = str(item["SSN"])
            if pd.notna(key_word["Data Base"]):
                df.at[index, 'key word'] = key_word["Data Base"]
                if key_word["Data Base"] != "TERMDATE":
                    carrierPlanDetails['SSN'] = carrierPlanDetails['SSN'].astype(str)
                    resultado = carrierPlanDetails[carrierPlanDetails['SSN'] == item_ssn]
                    if not resultado.empty:
                        datos = resultado[key_word["Data Base"]].values 
                        datos_filtrados = list({dato for dato in datos if '/' not in str(dato)} | {dato for dato in datos if '/' in str(dato)})
                        datos_joined = ';'.join(map(str, datos_filtrados))
                        df.at[index, 'Found Data'] = datos_joined
                    else:
                        df.at[index, 'Found Data'] = ''
                else:
                    carrierTermDates['SSN'] = carrierTermDates['SSN'].astype(str)
                    resultado = carrierTermDates[carrierTermDates['SSN'] == item_ssn]
                    if not resultado.empty:
                        datos = resultado[key_word["Data Base"]].values 
                        datos = [f"{date.astype('datetime64[D]').item().month}/{date.astype('datetime64[D]').item().day}/{date.astype('datetime64[D]').item().year}" for date in datos]
                        datos_filtrados = list({dato for dato in datos if '/' not in str(dato)} | {dato for dato in datos if '/' in str(dato)})
                        datos_joined = ';'.join(map(str, datos_filtrados))
                        df.at[index, 'Found Data'] = datos_joined
                    else:
                        df.at[index, 'Found Data'] = ''
            else:
                df.at[index, 'key word'] = 'Invalid field'
                df.at[index, 'Found Data'] = ''
        else:
            df.at[index, 'Found Data'] = ''
    
    return df
  

