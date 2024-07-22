import pandas as pd
from io import BytesIO
import base64

def formatExcel(contentBytes):
    decoded=base64.b64decode(contentBytes)
    excel=BytesIO(decoded)
    df = pd.read_excel(excel, skiprows=5, skipfooter=4,engine='openpyxl')

    # Agrupar los datos y calcular el número de cheques en vivo y los totales de cheques
    grouped_data = df.groupby('Client').agg(
        number_of_live_checks=pd.NamedAgg(column='Live Check Amount', aggfunc='count'),
        check_totals=pd.NamedAgg(column='Live Check Amount', aggfunc='sum')
    ).reset_index()

    # Escribir los datos agrupados en un archivo Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        grouped_data.to_excel(writer, index=False)
    output.seek(0)

    # Return the contentBytes
    return base64.b64encode(output.read()).decode('utf-8')
