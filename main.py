from flask import Flask, Response, jsonify,request, send_file
from formater import formatExcel, formatFromJson, createFile
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

@app.route('/create_file', methods=['GET'])
def create_file():
    response = createFile()
    return response

if __name__ == '__main__':
    app.run(debug=True, port=5000)