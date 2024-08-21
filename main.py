from flask import Flask, Response, jsonify,request, send_file
from formater import formatExcel, formatFromJson, discrepancies_report
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
    
    # Extraer el contenido de $content
    content = data.get('contentBytes')
    
    # Procesar el contenido
    processed_content = discrepancies_report(content)
    
    return send_file(processed_content,download_name='file.xlsx', as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheet.sheet")



if __name__ == '__main__':
    app.run(debug=True, port=5000)