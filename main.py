import base64
import datetime
from io import BytesIO
import io
from flask import Flask, Response, jsonify,request, send_file
import pandas as pd
from formater import formatExcel, formatFromJson, discrepancies_report, discrepancies_report_ssn
from flask_cors import CORS
from io import StringIO

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
    data = request.get_json()
    
    content = data[0].get('content')
    path = data[0].get('path')
    db = data[0]["carrierplandetails"]
    
    csv= "CARRIER,PEO_ID,CLIENT_ID,EMPLOYEE_ID,EMPLOYEE_NAME,EE_SSN,EMPLOYEE_BENEFIT_STATUS,PLAN_ID,INSURANCE_CLASS,OFFER_TYPE,ENROLLMENT_DATE,COVERAGE_END_DATE,TERMDATE,DEPENDENT_NAME,DEP_SSN,DEP_BIRTH_DATE,DEP_EFFECTIVE_DATE,RELATION_TYPE_CODE,RELATION_TYPE,RELATION_CODE,DEP_STATUS_CODE,DEP_STATUS,COBRA_ENROLLED,SALARY,EE_GENDER,DEP_GENDER\n"
    for rec in db:
        csv+=rec+"\n"

    csv_file = io.StringIO(csv)

    df = pd.read_csv(csv_file,dtype=str)
    processed_content = discrepancies_report(content, path,df)
    
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