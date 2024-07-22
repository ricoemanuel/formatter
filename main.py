from flask import Flask, Response, jsonify,request
from formater import formatExcel
from flask_cors import CORS
# Crear una nueva aplicación flask
app = Flask(__name__)
CORS(app)

@app.route('/format', methods=['POST'])
def format():
    contentBytes = request.get_data()

    content=formatExcel(contentBytes)
    headers = {
        "Content-Disposition": "attachment; filename='myfile.xlsx'",
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "*",
        "Access-Control-Allow-Methods": "GET",
    }
    return Response(
        response=content,
        headers=headers,
    )

if __name__ == '__main__':
    app.run(debug=True, port=5000)