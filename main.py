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
    try:
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
    except Exception as e:
        print(f"Error al guardar en Google Sheets: {e}")

def convertir_a_pdf(docx_path, pdf_path):
    """Convierte un archivo DOCX a PDF usando diferentes métodos"""
    
    print(f"Intentando convertir {docx_path} a PDF...")
    
    # Método 1: LibreOffice (método principal)
    try:
        print("Intentando con LibreOffice...")
        result = subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf", docx_path, 
            "--outdir", os.path.dirname(docx_path)
        ], check=True, timeout=60, capture_output=True, text=True)
        
        print(f"LibreOffice output: {result.stdout}")
        if os.path.exists(pdf_path):
            print(f"PDF creado exitosamente: {pdf_path}")
            return True
        else:
            print("LibreOffice no creó el archivo PDF")
    except (subprocess.TimeoutExpired, subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"LibreOffice falló: {e}")
    
    # Método 2: Unoconv (alternativa)
    try:
        print("Intentando con unoconv...")
        result = subprocess.run([
            "unoconv", "-f", "pdf", "-o", os.path.dirname(docx_path), docx_path
        ], check=True, timeout=60, capture_output=True, text=True)
        
        print(f"Unoconv output: {result.stdout}")
        if os.path.exists(pdf_path):
            print(f"PDF creado exitosamente con unoconv: {pdf_path}")
            return True
        else:
            print("Unoconv no creó el archivo PDF")
    except (subprocess.TimeoutExpired, subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"Unoconv falló: {e}")
    
    # Método 3: Pandoc (si está disponible)
    try:
        print("Intentando con pandoc...")
        result = subprocess.run([
            "pandoc", docx_path, "-o", pdf_path
        ], check=True, timeout=60, capture_output=True, text=True)
        
        print(f"Pandoc output: {result.stdout}")
        if os.path.exists(pdf_path):
            print(f"PDF creado exitosamente con pandoc: {pdf_path}")
            return True
        else:
            print("Pandoc no creó el archivo PDF")
    except (subprocess.TimeoutExpired, subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"Pandoc falló: {e}")
    
    # Método 4: Intentar con ruta absoluta de LibreOffice
    try:
        print("Intentando con ruta absoluta de LibreOffice...")
        result = subprocess.run([
            "/usr/bin/soffice", "--headless", "--convert-to", "pdf", docx_path, 
            "--outdir", os.path.dirname(docx_path)
        ], check=True, timeout=60, capture_output=True, text=True)
        
        print(f"LibreOffice (ruta absoluta) output: {result.stdout}")
        if os.path.exists(pdf_path):
            print(f"PDF creado exitosamente: {pdf_path}")
            return True
    except (subprocess.TimeoutExpired, subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"LibreOffice (ruta absoluta) falló: {e}")
    
    print("Todos los métodos de conversión fallaron")
    return False

@app.route('/generar-word', methods=['POST'])
def generar_word():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No se recibieron datos"}), 400

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
        pdf_path = f"cotizacion_{unique_id}.pdf"
        
        # Guardar el documento Word
        doc.save(docx_path)
        print(f"Documento Word guardado: {docx_path}")
        
        # Intentar convertir a PDF
        print("Iniciando conversión a PDF...")
        pdf_converted = convertir_a_pdf(docx_path, pdf_path)
        
        if pdf_converted and os.path.exists(pdf_path):
            # Enviar PDF
            print(f"Enviando PDF: {pdf_path}")
            response = send_file(pdf_path, as_attachment=True, download_name=pdf_path)
            
            @response.call_on_close
            def cleanup():
                try:
                    if os.path.exists(docx_path):
                        os.remove(docx_path)
                        print(f"Archivo DOCX eliminado: {docx_path}")
                    if os.path.exists(pdf_path):
                        os.remove(pdf_path)
                        print(f"Archivo PDF eliminado: {pdf_path}")
                except Exception as e:
                    print(f"Error en cleanup: {e}")
            
            return response
        else:
            # Si no se puede convertir a PDF, enviar el DOCX
            print(f"No se pudo convertir a PDF. Enviando documento Word: {docx_path}")
            response = send_file(docx_path, as_attachment=True, download_name=docx_path)
            
            @response.call_on_close
            def cleanup():
                try:
                    if os.path.exists(docx_path):
                        os.remove(docx_path)
                        print(f"Archivo DOCX eliminado: {docx_path}")
                except Exception as e:
                    print(f"Error en cleanup: {e}")
            
            return response

    except Exception as e:
        # Limpiar archivos en caso de error
        try:
            if 'docx_path' in locals() and os.path.exists(docx_path):
                os.remove(docx_path)
            if 'pdf_path' in locals() and os.path.exists(pdf_path):
                os.remove(pdf_path)
        except:
            pass
        
        print(f"Error en generar_word: {e}")
        return jsonify({"error": f"Error interno del servidor: {str(e)}"}), 500

if __name__ == '__main__':
    port = 5000
    app.run(debug=False, host='0.0.0.0', port=port)