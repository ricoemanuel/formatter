from flask import Flask, Response, jsonify, request, send_file
from formater import formatExcel, formatFromJson, discrepancies_report
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route('/format', methods=['POST'])
def format():
    if request.content_type != 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        return jsonify({"error": "Unsupported Media Type"}), 415

    try:
        contentBytes = request.get_data()
        content = formatExcel(contentBytes)
        return send_file(content, download_name='file.xlsx', as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/format-json', methods=['POST'])
def format_json():
    if request.content_type != 'application/json':
        return jsonify({"error": "Unsupported Media Type"}), 415

    data = request.get_json()
    response = formatFromJson(data)
    return response

@app.route('/discrepancies', methods=['POST'])
def discrepancies():
    if request.content_type != 'application/json':
        return jsonify({"error": "Unsupported Media Type"}), 415

    data = request.get_json()
    content = data.get('$content')
    processed_content = discrepancies_report(content)
    return send_file(processed_content, download_name='file.xlsx', as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    app.run(debug=True, port=5000)
