import pandas as pd
from io import BytesIO
import base64

def formatExcel(contentBytes):
    decoded = base64.b64decode(contentBytes)
    excel = BytesIO(decoded)
    
    df = pd.read_excel(excel, skiprows=5, skipfooter=4, engine='openpyxl')
    
    df_zero = pd.DataFrame(columns=df.columns)
  
    grouped_data = df.groupby('Client').agg(
        number_of_live_checks=pd.NamedAgg(column='Live Check Amount', aggfunc=count_non_zero),
        check_totals=pd.NamedAgg(column='Live Check Amount', aggfunc='sum')
    ).reset_index()

    grouped_data.loc[len(grouped_data)]={
        'Client': 'Totals',
        'number_of_live_checks': grouped_data['number_of_live_checks'].sum(),
        'check_totals': grouped_data['check_totals'].sum()
    }
    

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        grouped_data.to_excel(writer, index=False)
    
    output.seek(0)

    return output, df_zero

def count_non_zero(series):
    return (series != 0).sum()
