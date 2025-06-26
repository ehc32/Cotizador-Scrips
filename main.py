from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docx import Document
import os
import uuid
import subprocess
import gspread
from google.oauth2.service_account import Credentials
import json 
from docxtpl import DocxTemplate

app = Flask(__name__)
CORS(app)

def numero_a_texto(numero):
    try:
        from num2words import num2words
        return num2words(numero, lang='es').capitalize() + " pesos colombianos"
    except:
        return str(numero)

def guardar_en_google_sheets(data):
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    creds_dict = json.loads(os.environ.get("GOOGLE_CREDENTIALS"))
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_key("1qWXMepGrgxjZK9QLPxcLcCDlHtsLBIH9Fo0GJDgr_Go").sheet1  

    fila = [
        data.get("nombre", ""),
        data.get("correo", ""),
        data.get("Subtotal_1", ""),
        data.get("Subtotal_2", ""),
        data.get("Total", ""),
        data.get("texto", "")
    ]
    sheet.append_row(fila)

@app.route('/generar-word', methods=['POST'])
def generar_word():
    data = request.get_json()

    data["Acompañamie"] = "$ 1.516.141"
    data["Diseño_Calculo"] = "$ 23.918.292"
    data["Diseño_Sanitario"] = "$ 20.501.393"

    def extraer_numero(texto):
        if not texto:
            return 0
        return int(''.join(filter(str.isdigit, str(texto))) or 0)

    def formatear_moneda(numero):
        return "${:,.0f}".format(numero).replace(",", ".")

    subtotal1 = extraer_numero(data.get("Subtotal_1")) or (
        extraer_numero(data.get("Diseño_Ar")) +
        extraer_numero(data.get("Diseño_Calculo")) +
        extraer_numero(data.get("Acompañamie"))
    )
    subtotal2 = extraer_numero(data.get("Subtotal_2")) or (
        extraer_numero(data.get("Diseño_Calculo")) +
        extraer_numero(data.get("Diseño_Sanitario")) +
        extraer_numero(data.get("Presupuesta"))
    )
    total = subtotal1 + subtotal2
    data["Subtotal_1"] = formatear_moneda(subtotal1)
    data["Subtotal_2"] = formatear_moneda(subtotal2)
    data["Total"] = formatear_moneda(total)
    data["texto"] = numero_a_texto(total)

    guardar_en_google_sheets(data)

    plantilla_path = os.path.join(os.path.dirname(__file__), "plantilla.docx")
    if not os.path.exists(plantilla_path):
        return jsonify({"error": "No se encontró la plantilla Word"}), 500

    doc = DocxTemplate(plantilla_path)
    doc.render(data)

    unique_id = str(uuid.uuid4())
    docx_path = f"cotizacion_{unique_id}.docx"
    doc.save(docx_path)

    pdf_path = f"cotizacion_{unique_id}.pdf"
    pdf_full_path = os.path.join(os.path.dirname(docx_path), pdf_path)
    try:
        subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf", docx_path, 
            "--outdir", os.path.dirname(docx_path)
        ], check=True, timeout=30)
        if not os.path.exists(pdf_full_path):
            raise Exception("La conversión a PDF falló.")
    except Exception as e:
        os.remove(docx_path)  
        return jsonify({"error": f"Error al convertir a PDF: {str(e)}"}), 500

    response = send_file(pdf_full_path, as_attachment=True, download_name=pdf_path)

    @response.call_on_close
    def cleanup():
        try:
            os.remove(docx_path)
            os.remove(pdf_full_path)
        except Exception:
            pass

    return response

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)