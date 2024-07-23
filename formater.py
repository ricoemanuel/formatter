import pandas as pd
from io import BytesIO
import base64

def formatExcel(contentBytes):
    decoded = base64.b64decode(contentBytes)
    excel = BytesIO(decoded)
    
    df = pd.read_excel(excel, skiprows=5, skipfooter=4, engine='openpyxl')
    
    # Create a new DataFrame for rows where 'Live Check Amount' is 0
    df_zero = pd.DataFrame(columns=df.columns)
    
    for index, row in df.iterrows():
        if row['Live Check Amount'] == 0:
            # Check if the client is already in the df_zero DataFrame
            if row['Client'] not in df_zero['Client'].values:
                # Append the row to df_zero DataFrame
                df_zero = df_zero.append(row)
    
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
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        grouped_data.to_excel(writer, index=False)
    
    output.seek(0)

    return output, df_zero
