import pandas as pd
from io import BytesIO
import base64

def formatExcel(contentBytes):
    decoded = base64.b64decode(contentBytes)
    excel = BytesIO(decoded)
    
    df = pd.read_excel(excel, skiprows=5, skipfooter=4, engine='openpyxl')
    
    df_zero=df[df['Live Check Amount'] == 0]
    df_zero=df_zero.groupby('Client').agg(
        number_of_live_checks=pd.NamedAgg(column='Live Check Amount', aggfunc='count'),
        check_totals=pd.NamedAgg(column='Live Check Amount', aggfunc='sum')
    ).reset_index()

    df = df[df['Live Check Amount'] > 0]
    grouped_data = df.groupby('Client').agg(
        number_of_live_checks=pd.NamedAgg(column='Live Check Amount', aggfunc='count'),
        check_totals=pd.NamedAgg(column='Live Check Amount', aggfunc='sum')
    ).reset_index()

    df_concat = pd.concat([df_zero, grouped_data])

    grouped_data=df_concat.drop_duplicates(subset='Client', keep='last', inplace=True)

    grouped_data.loc[len(grouped_data)]={
        'Client': 'Totals',
        'number_of_live_checks': grouped_data['number_of_live_checks'].sum(),
        'check_totals': grouped_data['check_totals'].sum()
    }
    

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        grouped_data.to_excel(writer, index=False)
    
    output.seek(0)

    return output
