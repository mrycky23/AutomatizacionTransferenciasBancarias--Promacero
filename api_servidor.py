# api_servidor.py  (corre en la PC de oficina)
from flask import Flask, request, jsonify
from openpyxl import load_workbook
import base64, os
app = Flask(__name__)

@app.route('/transferencia', methods=['POST'])
def registrar():
    data = request.json
    # Guarda imagen, actualiza Excel, genera PDF
    ...
    return jsonify({"ok": True})

app.run(host="0.0.0.0", port=5000)