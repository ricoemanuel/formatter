from flask import Flask, jsonify,request
from formater import formatExcel
# Crear una nueva aplicación flask
app = Flask(__name__)


@app.route('/format', methods=['POST'])
def format():
    contentBytes = request.get_data()

    content=formatExcel(contentBytes)

    return jsonify({'data': content})

if __name__ == '__main__':
    app.run(debug=True)