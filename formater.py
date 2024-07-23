import pandas as pd
from io import BytesIO
import base64

def formatExcel(contentBytes):
    decoded = base64.b64decode(contentBytes)
    excel = BytesIO(decoded)
    
    df = pd.read_excel(excel, skiprows=5, skipfooter=4, engine='openpyxl')

    grouped_data = df.groupby('Client').agg(
        number_of_live_checks=pd.NamedAgg(column='Live Check Amount', aggfunc='count'),
        check_totals=pd.NamedAgg(column='Live Check Amount', aggfunc='sum')
    ).reset_index()
    sum_row = pd.DataFrame({
        'Client': ['suma'],
        'number_of_live_checks': [grouped_data['number_of_live_checks'].sum()],
        'check_totals': [grouped_data['check_totals'].sum()]
    })
    grouped_data = grouped_data.add(sum_row)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        grouped_data.to_excel(writer, index=False)
    
    output.seek(0)

    return output
