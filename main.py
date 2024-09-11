import base64
import datetime
from io import BytesIO
from flask import Flask, Response, jsonify,request, send_file
import pandas as pd
from formater import formatExcel, formatFromJson, discrepancies_report, discrepancies_report_ssn
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route('/format', methods=['POST'])
def format():
    contentBytes = request.get_data()

    content=formatExcel(contentBytes)
    
    return send_file(content,download_name='file.xlsx', as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheet.sheet")

@app.route('/format-json', methods=['POST'])
def format_json():
    data = request.get_json()
    response = formatFromJson(data)
    return response


@app.route('/discrepancies', methods=['POST'])
def discrepancies():
    # Obtener el contenido de la solicitud POST
    data = request.get_json()
    
    # Extraer el contenido de 'content'
    content = data[0].get('content')
    path = data[0].get('path')
    columns = data[0]["carrierplandetails"][0]
    rows = data[0]["carrierplandetails"][1:]

    dfcarrierplandetails = pd.DataFrame(rows, columns=columns)

    columns = data[0]["carrierplandetailsByDep"][0]
    rows = data[0]["carrierplandetailsByDep"][1:]
    for row in rows:
        excel_serial_date = int(row[-1])
        date = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=excel_serial_date)
        row[-1] = date
    dfcarrierplandetailsByDep = pd.DataFrame(rows, columns=columns)

    columns = data[0]["termdates"][0]
    rows = data[0]["termdates"][1:]

    df_concatenado = pd.concat([dfcarrierplandetails, dfcarrierplandetailsByDep], ignore_index=True)


    for row in rows:
        excel_serial_date = int(row[-1])
        date = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=excel_serial_date)
        row[-1] = date
    dftermdates = pd.DataFrame(rows, columns=columns)

  
    processed_content = discrepancies_report(content, path,df_concatenado,dftermdates)
    
    # Wrap the bytes in a BytesIO object
    processed_file = BytesIO(processed_content)
    
    return send_file(processed_file, download_name='file.xlsx', as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route('/list_ssn', methods=['POST'])
def discrepancies_ssns():
    # Obtener el contenido de la solicitud POST
    data = request.get_json()
    
    # Extraer el contenido de 'content'
    content = data[0].get('content')
    path = data[0].get('path')
    return discrepancies_report_ssn(content, path)
    
if __name__ == '__main__':
    app.run(debug=True, port=5000)