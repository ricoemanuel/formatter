from flask import Flask, jsonify,request
from formater import formatExcel
from flask_cors import CORS
# Crear una nueva aplicación flask
app = Flask(__name__)
CORS(app)

@app.route('/format', methods=['POST'])
def format():
    contentBytes = request.get_data()

    content=formatExcel(contentBytes)

    return jsonify({'data': content})

if __name__ == '__main__':
    app.run(debug=True, port=5000)