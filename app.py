import os
import csv
import json
import gc
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify
import unicodedata
from difflib import SequenceMatcher

app = Flask(__name__)
app.secret_key = "clave_secreta_cavelier"

# Configuraciones de carpetas
UPLOAD_FOLDER = "uploads"
PDF_FOLDER = "static/pdfs"
for f in [UPLOAD_FOLDER, PDF_FOLDER]:
    os.makedirs(f, exist_ok=True)

UMBRAL_CORTE = 70.0

def limpiar(texto):
    if not texto: return ""
    texto = str(texto).upper().strip()
    # Eliminar tildes y caracteres especiales
    nfkd_form = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def extraer_clases(campo):
    if not campo: return set()
    import re
    return {int(n) for n in re.findall(r'\d+', str(campo))}

def calcular_similitud(a, b):
    return SequenceMatcher(None, a, b).ratio() * 100

@app.route("/")
def inicio():
    return render_template("index.html") # Asegúrate de tener tus plantillas

@app.route("/procesar", methods=["POST"])
def procesar():
    file_c = request.files.get("archivo_clientes")
    file_g = request.files.get("archivo_gaceta")
    
    if not file_c or not file_g:
        flash("Sube ambos archivos en formato CSV")
        return redirect(url_for("inicio"))

    path_c = os.path.join(UPLOAD_FOLDER, "clientes.csv")
    path_g = os.path.join(UPLOAD_FOLDER, "gaceta.csv")
    file_c.save(path_c)
    file_g.save(path_g)

    def iterar_csv(path):
        # Intentamos latin-1 que es el estándar de Excel en Windows
        with open(path, mode='r', encoding='latin-1') as f:
            # Autodetectar si usa coma o punto y coma
            linea = f.readline()
            separador = ';' if ';' in linea else ','
            f.seek(0)
            reader = csv.reader(f, delimiter=separador)
            next(reader, None) # Saltar cabecera
            for row in reader:
                try:
                    if len(row) < 2 or not row[1]: continue
                    yield {
                        "ID": row[0],
                        "Marca_Orig": row[1],
                        "Marca_Limpia": limpiar(row[1]),
                        "Clases": extraer_clases(row[8]) if len(row) > 8 else set()
                    }
                except: continue

    resultados = []
    # PROCESAMIENTO ULTRA-LIGERO
    for c in iterar_csv(path_c):
        for g in iterar_csv(path_g):
            # Filtro rápido: si comparten al menos una clase
            if c["Clases"] & g["Clases"]:
                score = calcular_similitud(c["Marca_Limpia"], g["Marca_Limpia"])
                if score >= UMBRAL_CORTE:
                    resultados.append({
                        "pdf_id": f"{c['ID']}_{g['ID']}",
                        "marca_cliente": c["Marca_Orig"],
                        "marca_gaceta": g["Marca_Orig"],
                        "score": round(score, 2)
                    })
        gc.collect() # Forzar limpieza de RAM constante

    if not resultados:
        flash("No se encontraron conflictos.")
        return redirect(url_for("inicio"))

    session["resultados"] = resultados
    return redirect(url_for("resultados_view"))

@app.route("/resultados")
def resultados_view():
    res = session.get("resultados", [])
    return render_template("resultados.html", resultados=res)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
