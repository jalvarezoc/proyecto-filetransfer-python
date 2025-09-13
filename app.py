from flask import Flask, render_template, request, send_file, redirect, url_for, jsonify
import os
import datetime
import json
from docxtpl import DocxTemplate

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
SAP_DATA_FILE = "data/sap_projects.json"

# Creamos subcarpetas para cada categoría
CATEGORIAS = ["internas", "h2h", "noh2h"]
for cat in CATEGORIAS:
    os.makedirs(os.path.join(UPLOAD_FOLDER, cat), exist_ok=True)

# Crear directorio para datos SAP si no existe
os.makedirs(os.path.dirname(SAP_DATA_FILE), exist_ok=True)

# Función que lista documentos de una categoría
def get_documentos(categoria):
    folder = os.path.join(UPLOAD_FOLDER, categoria)
    docs = []
    for filename in os.listdir(folder):
        filepath = os.path.join(folder, filename)
        if os.path.isfile(filepath):
            stats = os.stat(filepath)
            docs.append({
                "id": filename,
                "nombre": filename,
                "fecha_modificacion": datetime.datetime.fromtimestamp(stats.st_mtime).strftime("%d/%m/%Y %H:%M"),
                "tamano": f"{round(stats.st_size/1024,1)} KB",
                "categoria": categoria
            })
    return docs

# Función para guardar datos SAP - SOLO GUARDA EL ÚLTIMO REGISTRO
def guardar_datos_sap(datos):
    try:
        # Siempre crear una nueva lista con solo el último registro
        datos_existentes = [datos]
        
        # Guardar con formato
        with open(SAP_DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(datos_existentes, f, indent=4, ensure_ascii=False)
            
        return True
    except Exception as e:
        print(f"Error guardando datos SAP: {e}")
        return False


@app.route("/")
def index():
    return render_template("index.html")

@app.route("/<categoria>")
def listar_categoria(categoria):
    print(categoria)
    if categoria not in CATEGORIAS:
        return "Categoría no encontrada", 404
    
    documentos = get_documentos(categoria)
    
    # Renderizar plantilla específica para cada categoría
    if categoria == "internas":
        return render_template("interfaz_internas.html", documentos=documentos, tipo=categoria.upper())
    elif categoria == "h2h":
        return render_template("interfaz_h2h.html", documentos=documentos, tipo=categoria.upper())
    elif categoria == "noh2h":
        return render_template("lista.html", documentos=documentos, tipo=categoria.upper())

@app.route("/procesar_sap", methods=["POST"])
def procesar_sap():
    try:
        print("Datos recibidos:", request.form.to_dict())
        
        # Obtener datos del formulario
        datos = {
            "modalidad": request.form.get("tipoDocumento"),
            "cantidad_usuarios": request.form.get("cantidadUsuarios"),
            "nombre_mvp": request.form.get("nombreMVP"),
            "nombre_empresa": request.form.get("nombreEmpresa"),
            "formato_directorio": request.form.get("formatoDirectorio"),
            "formato_reglas": request.form.get("formatoReglas"),
            "ambientes": [],
            "fecha_creacion": datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        }
        
        # Procesar ambientes
        ambientes = request.form.getlist("ambiente[]")
        sender_codes = request.form.getlist("senderCode[]")
        receiver_codes = request.form.getlist("receiverCode[]")
        descriptions = request.form.getlist("description[]")
        text1_list = request.form.getlist("text1[]")
        text2_list = request.form.getlist("text2[]")
        text6_list = request.form.getlist("text6[]")
        text7_list = request.form.getlist("text7[]")
        
        for i in range(len(ambientes)):
            ambiente_data = {
                "ambiente": ambientes[i],
                "sender_code": sender_codes[i],
                "receiver_code": receiver_codes[i],
                "description": descriptions[i],
                "text1": text1_list[i],
                "text2": text2_list[i],
                "text6": text6_list[i],
                "text7": text7_list[i]
            }
            datos["ambientes"].append(ambiente_data)
        
        # Guardar datos (solo el último registro)
        if guardar_datos_sap(datos):
            # Generar el documento Word inmediatamente después de guardar
            output_path = generar_documento_word()
            
            if output_path and os.path.exists(output_path):
                # Devolver tanto el mensaje de éxito como la información del documento
                return jsonify({
                    "success": True, 
                    "message": "Proyecto SAP guardado correctamente",
                    "document_path": output_path,
                    "document_name": f"PROYECTO_SAP_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                })
            else:
                return jsonify({
                    "success": True, 
                    "message": "Proyecto SAP guardado pero error al generar documento",
                    "document_path": None
                })
        else:
            return jsonify({"success": False, "message": "Error al guardar el proyecto"})
            
    except Exception as e:
        print(f"Error en procesar_sap: {str(e)}")
        return jsonify({"success": False, "message": f"Error: {str(e)}"})
    
@app.route("/documento/<categoria>/<path:doc_id>")
def documento(categoria, doc_id):
    filepath = os.path.join(UPLOAD_FOLDER, categoria, doc_id)
    if not os.path.exists(filepath):
        return "Documento no encontrado", 404
    stats = os.stat(filepath)
    doc = {
        "id": doc_id,
        "nombre": doc_id,
        "fecha_modificacion": datetime.datetime.fromtimestamp(stats.st_mtime).strftime("%d/%m/%Y %H:%M"),
        "tamano": f"{round(stats.st_size/1024,1)} KB",
        "categoria": categoria
    }
    return render_template("documento.html", documento=doc)

@app.route("/upload/<categoria>", methods=["POST"])
def upload(categoria):
    if categoria not in CATEGORIAS:
        return "Categoría no encontrada", 404
    if "file" not in request.files:
        return "No file part"
    file = request.files["file"]
    if file.filename == "":
        return "No selected file"
    filepath = os.path.join(UPLOAD_FOLDER, categoria, file.filename)
    file.save(filepath)
    return redirect(url_for("listar_categoria", categoria=categoria))

@app.route("/descargar_documento")
def descargar_documento():
    # Obtener nombre personalizado si se proporciona
    custom_name = request.args.get('name', None)
    
    # Usar el último documento generado
    output_path = "output/documento_generado.docx"
    
    if os.path.exists(output_path):
        download_name = custom_name or f"PROYECTO_SAP_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        return "Error al generar el documento", 404



def generar_documento_word():
    try:
        # Leer los datos del JSON
        if not os.path.exists(SAP_DATA_FILE) or os.path.getsize(SAP_DATA_FILE) == 0:
            return None
        
        with open(SAP_DATA_FILE, 'r', encoding='utf-8') as f:
            datos = json.load(f)
        
        # Si los datos están en una lista, tomar el primero
        if isinstance(datos, list) and len(datos) > 0:
            datos = datos[0]
        
        # Cargar la plantilla Word
        template_path = "templates/plantilla_sap.docx"
        # if not os.path.exists(template_path):
        #     # Crear una plantilla básica si no existe
        #     crear_plantilla_basica(template_path)
        
        doc = DocxTemplate(template_path)
        
        # Preparar el contexto para la plantilla
        context = {
            'fecha_actual': datetime.datetime.now().strftime("%d/%m/%Y"),
            'hora_actual': datetime.datetime.now().strftime("%H:%M:%S"),
            'modalidad': datos.get('modalidad', 'N/A'),
            'cantidad_usuarios': datos.get('cantidad_usuarios', 'N/A'),
            'nombre_mvp': datos.get('nombre_mvp', 'N/A'),
            'nombre_empresa': datos.get('nombre_empresa', 'N/A'),
            'formato_directorio': datos.get('formato_directorio', 'N/A'),
            'formato_reglas': datos.get('formato_reglas', 'N/A'),
            'fecha_creacion': datos.get('fecha_creacion', 'N/A'),
            'ambientes': datos.get('ambientes', [])
        }
        
        # Renderizar el documento
        doc.render(context)
        
        # Guardar el documento generado
        output_path = "output/documento_generado.docx"
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)
        
        return output_path
        
    except Exception as e:
        print(f"Error generando documento Word: {e}")
        return None


@app.route("/download/<categoria>/<path:filename>")
def download(categoria, filename):
    return send_file(os.path.join(UPLOAD_FOLDER, categoria, filename), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)