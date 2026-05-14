"""
CAVELIER ABOGADOS — Sistema de Vigilancia de Marcas
Versión ligera: sin pandas/numpy para correr en Render free tier
"""
import os, re, unicodedata, json, secrets
from io import BytesIO
from difflib import SequenceMatcher
from datetime import datetime

import openpyxl
from PIL import Image
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from flask import (Flask, render_template, request, redirect,
                   url_for, session, send_file, jsonify, flash)

app = Flask(__name__)
app.secret_key = secrets.token_hex(32)

UMBRAL_CORTE  = 60
PDF_FOLDER    = "pdfs_generados"
CARPETA_IMG   = "temp_logos"
TEXTO_PIE     = "Cra. 4 N° 72A - 35 Bogotá D.C. | Tel. (+57) 601 3473611 | cavelier@cavelier.com"

for folder in [PDF_FOLDER, CARPETA_IMG]:
    os.makedirs(folder, exist_ok=True)

USUARIOS = {
    "cavelier":  "marcas2024",
    "abogado1":  "clave123",
    "abogado2":  "clave456",
}

CLASES_VINCULADAS = [
    { 1,  2}, { 1,  5}, { 1,  6}, { 1, 16}, { 1, 17}, { 1, 19}, { 1, 31}, { 1, 40},
    { 2,  4}, { 2, 16}, { 2, 19},
    { 3,  5}, { 3, 10}, { 3, 14}, { 3, 18}, { 3, 21}, { 3, 25}, { 3, 44},
    { 4, 12}, { 4, 13}, { 4, 37}, { 4, 39}, { 4, 40},
    { 5, 10}, { 5, 29}, { 5, 30}, { 5, 31}, { 5, 44},
    { 6,  8}, { 6, 19},
    { 7,  8}, { 7, 12}, { 7, 17},
    { 8, 21},
    { 9, 10}, { 9, 15}, { 9, 16}, { 9, 28}, { 9, 35}, { 9, 38}, { 9, 41}, { 9, 42}, { 9, 45},
    {10, 44},
    {11, 19}, {11, 21}, {11, 37},
    {12, 37}, {12, 39}, {12, 40},
    {14, 18}, {14, 25},
    {15, 41},
    {16, 17}, {16, 35}, {16, 38}, {16, 41},
    {18, 25},
    {19, 27}, {19, 37},
    {20, 21}, {20, 22}, {20, 24}, {20, 27},
    {21, 24},
    {22, 24}, {22, 27}, {22, 28},
    {23, 24}, {23, 26},
    {24, 25}, {24, 26},
    {26, 31},
    {28, 41},
    {29, 30}, {29, 31}, {29, 32}, {29, 43},
    {30, 31}, {30, 32}, {30, 43},
    {31, 32}, {31, 43}, {31, 44},
    {32, 33}, {32, 43},
    {33, 34}, {33, 43},
    {35, 36}, {35, 38}, {35, 41}, {35, 42},
    {36, 37},
    {37, 40},
    {38, 41}, {38, 42},
    {39, 43},
]

def safe_text(text):
    if text is None or str(text).strip() in ("", "nan", "None"):
        return "N/A"
    text = str(text).replace('\u2018',"'").replace('\u2019',"'") \
                    .replace('\u201c','"').replace('\u201d','"').replace('\u2013','-')
    return text.encode('latin-1','replace').decode('latin-1')

def limpiar_titular(texto):
    if not texto: return "Sin Titular"
    t = str(texto).upper()
    partes = re.split(r';| Y ', t)
    patrones = [r", CL",r", KR",r", CR",r", CRA",r", CALLE",r", CARRERA",
                r", AVENIDA",r", AV\.",r", EDIFICIO",r", APTO",r", BOGOTA"]
    limpios = []
    for p in partes:
        tmp = p.strip()
        for pat in patrones:
            tmp = re.split(pat, tmp)[0]
        limpios.append(tmp.strip())
    res = " / ".join(filter(None, limpios))
    return res if res else t.strip()

def limpiar(texto):
    if not texto or str(texto).strip() == "": return ""
    texto = str(texto).upper().strip()
    if len(texto) <= 1: return ""
    texto = unicodedata.normalize('NFD', texto).encode('ascii','ignore').decode('utf-8')
    return re.sub(r'[^A-Z0-9 ]','', texto)

def limpiar_id(texto):
    return re.sub(r'[^A-Z0-9]','_', str(texto).upper())

def calcular_similitud(a, b):
    if not a or not b: return 0
    ratio = SequenceMatcher(None, a, b).ratio()
    if a in b or b in a: ratio = max(ratio, 0.85)
    return round(ratio * 100, 2)

def extraer_clases(texto):
    if not texto or str(texto).strip() == "": return set()
    return set(int(x) for x in re.findall(r'\d+', str(texto)))

def clases_en_conflicto(cc, cg):
    for a in cc:
        for b in cg:
            if a == b: return True
            for par in CLASES_VINCULADAS:
                if a in par and b in par: return True
    return False

def calcular_clases_conflicto(cc, cg):
    resultado = set()
    for a in cc:
        for b in cg:
            if a == b:
                resultado.add(a); resultado.add(b)
            for par in CLASES_VINCULADAS:
                if a in par and b in par:
                    resultado.add(a); resultado.add(b)
    return resultado

def filtrar_productos(texto, clases):
    if not texto or str(texto).strip() == "": return "N/A"
    bloques = re.split(r'(?=\b\d{1,2}\.\s)', str(texto).strip())
    resultado = []
    for bloque in bloques:
        bloque = bloque.strip()
        if not bloque: continue
        m = re.match(r'^(\d{1,2})\.\s', bloque)
        if m and int(m.group(1)) in clases:
            resultado.append(bloque)
        elif not m:
            resultado.append(bloque)
    return " ".join(resultado) if resultado else str(texto)

def formatear_fecha(valor):
    if valor is None: return "N/A"
    if isinstance(valor, datetime):
        return valor.strftime('%d/%m/%Y')
    return str(valor)

def leer_excel_bytes(file_bytes):
    bio = BytesIO(file_bytes)
    wb  = openpyxl.load_workbook(bio, data_only=True)
    sheet = wb.active
    if hasattr(sheet, '_images'):
        for img in sheet._images:
            try:
                row = img.anchor._from.row + 1
                exp_id = sheet.cell(row=row, column=1).value
                if exp_id:
                    ruta = os.path.join(CARPETA_IMG, f"{limpiar_id(str(exp_id))}.png")
                    with open(ruta, "wb") as f: f.write(img._data())
            except: continue
    rows = list(sheet.iter_rows(values_only=True))
    if len(rows) < 2: return []
    registros = []
    for row in rows[1:]:
        if len(row) < 9: continue
        exp_id  = row[0]
        fecha   = row[3]
        titular = row[7]
        marca   = row[1]
        prod    = row[10] if len(row) > 10 else row[-1]
        clases  = row[8]
        marca_str = str(marca).strip() if marca else ""
        if not marca_str or marca_str.lower() == "none": continue
        registros.append({
            "Expediente_ID":   str(exp_id) if exp_id else "N/A",
            "Fecha_Rad":       formatear_fecha(fecha),
            "Titular":         limpiar_titular(titular),
            "Marca_Original":  marca_str,
            "Marca_Limpia":    limpiar(marca_str),
            "Productos_Texto": str(prod) if prod else "",
            "Clases":          extraer_clases(clases),
        })
    return registros

class PDFCavelier(FPDF):
    def header(self):
        if os.path.exists("encabezado.png"):
            self.image("encabezado.png", x=60, y=8, w=90)
        self.ln(25)
    def footer(self):
        self.set_y(-15)
        self.set_text_color(44,62,80)
        self.set_font("Helvetica","B",7)
        self.set_draw_color(44,62,80)
        self.line(10,282,200,282)
        self.cell(0,10,f"{TEXTO_PIE}  |  Página {self.page_no()}",align="C")

def generar_pdf(c, g, score, clases_conf, concepto="", ruta_pdf=None):
    pdf = PDFCavelier()
    pdf.add_page()
    pdf.set_font("Helvetica","B",16); pdf.set_text_color(44,62,80)
    pdf.cell(190,10,"INFORME TÉCNICO DE OPOSICIÓN",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.ln(2)
    pdf.set_fill_color(240,240,240)
    pdf.rect(55, pdf.get_y(), 80, 4, style="F")
    color = (231,76,60) if score>=80 else (241,196,15) if score>=70 else (46,204,113)
    pdf.set_fill_color(*color)
    pdf.rect(55, pdf.get_y(), score*0.8, 4, style="F")
    pdf.set_font("Helvetica","B",10); pdf.set_x(140)
    pdf.cell(30,4,f"{score}% Similitud", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.ln(10)
    y_img = pdf.get_y()
    def colocar_logo(ruta, x_marco):
        ancho_marco, alto_marco, padding = 50, 40, 4
        ancho_max = ancho_marco - padding*2
        alto_max  = alto_marco  - padding*2
        if os.path.exists(ruta):
            pdf.set_draw_color(220,220,220)
            pdf.rect(x_marco, y_img, ancho_marco, alto_marco)
            DPI = 96
            max_px_w = int(round(ancho_max/25.4*DPI))
            max_px_h = int(round(alto_max /25.4*DPI))
            with Image.open(ruta) as im:
                img_rgba = im.convert("RGBA")
                wo, ho = img_rgba.size
            escala  = min(max_px_w/wo, max_px_h/ho)
            nw = max(1,int(wo*escala)); nh = max(1,int(ho*escala))
            resized = img_rgba.resize((nw,nh), Image.LANCZOS)
            tmp = ruta.replace(".png","_t96.png")
            resized.save(tmp, dpi=(DPI,DPI))
            pdf.image(tmp, x=x_marco+(ancho_marco-nw/DPI*25.4)/2,
                      y=y_img+(alto_marco-nh/DPI*25.4)/2,
                      w=nw/DPI*25.4, h=nh/DPI*25.4)
            try: os.remove(tmp)
            except: pass
    img_c = os.path.join(CARPETA_IMG, f"{limpiar_id(c['Expediente_ID'])}.png")
    img_g = os.path.join(CARPETA_IMG, f"{limpiar_id(g['Expediente_ID'])}.png")
    colocar_logo(img_c, 25)
    colocar_logo(img_g, 115)
    pdf.set_y(y_img + 45)
    pdf.set_font("Helvetica","B",11); pdf.set_text_color(0,0,0)
    pdf.cell(95,7,f"CLIENTE: {safe_text(c['Marca_Original'])}", align="L")
    pdf.cell(95,7,f"GACETA: {safe_text(g['Marca_Original'])}",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica","",9); pdf.set_text_color(80,80,80)
    pdf.cell(95,5,f"Exp: {safe_text(c['Expediente_ID'])}", align="L")
    pdf.cell(95,5,f"Exp: {safe_text(g['Expediente_ID'])}",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.cell(95,5,f"Rad: {safe_text(c['Fecha_Rad'])}", align="L")
    pdf.cell(95,5,f"Rad: {safe_text(g['Fecha_Rad'])}",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica","I",8)
    y_a = pdf.get_y()
    pdf.multi_cell(90,4,f"Titular: {safe_text(c['Titular'])}", align="L")
    yc = pdf.get_y()
    pdf.set_y(y_a); pdf.set_x(105)
    pdf.multi_cell(90,4,f"Titular: {safe_text(g['Titular'])}", align="L")
    yg = pdf.get_y()
    pdf.set_y(max(yc,yg)+5)
    ALTO_L = 4; MARGEN_INF = 20; ALTO_T = 7
    espacio = 297 - MARGEN_INF - pdf.get_y() - ALTO_T
    if concepto.strip(): espacio -= 50
    max_lin = max(4, int(espacio / ALTO_L))
    CPL = 62
    def truncar(t, ml):
        t = safe_text(t)
        lim = ml * CPL
        if len(t) > lim:
            c2 = t[:lim].rfind(' ')
            t = t[:c2 if c2>0 else lim] + "..."
        return t
    tc = truncar(filtrar_productos(c['Productos_Texto'], clases_conf), max_lin)
    tg = truncar(filtrar_productos(g['Productos_Texto'], clases_conf), max_lin)
    pdf.set_text_color(44,62,80); pdf.set_font("Helvetica","B",10)
    pdf.cell(95,ALTO_T," PRODUCTOS (CLIENTE):", border="B", align="L")
    pdf.cell(95,ALTO_T," PRODUCTOS (GACETA):",  border="B",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_text_color(0,0,0); pdf.set_font("Helvetica","",8)
    yp = pdf.get_y()
    pdf.multi_cell(90, ALTO_L, tc, align="J")
    pdf.set_y(yp); pdf.set_x(105)
    pdf.multi_cell(90, ALTO_L, tg, align="J")
    if concepto.strip():
        pdf.ln(6)
        pdf.set_font("Helvetica","B",10); pdf.set_text_color(44,62,80)
        pdf.cell(190,7," CONCEPTO JURÍDICO:", border="B",
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
        pdf.set_font("Helvetica","",9); pdf.set_text_color(0,0,0)
        pdf.multi_cell(190, 5,
                       concepto.encode('latin-1','replace').decode('latin-1'), align="J")
    buf = BytesIO()
    pdf.output(buf)
    buf.seek(0)
    if ruta_pdf:
        with open(ruta_pdf,"wb") as f: f.write(buf.getvalue())
        buf.seek(0)
    return buf

def login_requerido(f):
    from functools import wraps
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("usuario"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper

@app.route("/", methods=["GET","POST"])
def login():
    if session.get("usuario"):
        return redirect(url_for("inicio"))
    error = None
    if request.method == "POST":
        u = request.form.get("usuario","").strip()
        p = request.form.get("password","").strip()
        if USUARIOS.get(u) == p:
            session["usuario"] = u
            return redirect(url_for("inicio"))
        error = "Usuario o contraseña incorrectos."
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/inicio")
@login_requerido
def inicio():
    return render_template("inicio.html", usuario=session["usuario"])

@app.route("/procesar", methods=["POST"])
@login_requerido
def procesar():
    archivo_c = request.files.get("archivo_clientes")
    archivo_g = request.files.get("archivo_gaceta")
    if not archivo_c or not archivo_g:
        flash("Debes subir los dos archivos Excel.")
        return redirect(url_for("inicio"))
    for f in os.listdir(CARPETA_IMG):
        try: os.remove(os.path.join(CARPETA_IMG, f))
        except: pass
    df_c = leer_excel_bytes(archivo_c.read())
    df_g = leer_excel_bytes(archivo_g.read())
    if not df_c or not df_g:
        flash("No se pudieron leer los archivos. Verifica el formato.")
        return redirect(url_for("inicio"))
    resultados = []
    datos_completos = []
    for c in df_c:
        if not c["Marca_Limpia"]: continue
        for g in df_g:
            if not g["Marca_Limpia"]: continue
            score = calcular_similitud(c["Marca_Limpia"], g["Marca_Limpia"])
            if score < UMBRAL_CORTE: continue
            clases_c = c.get("Clases", set()) or set()
            clases_g = g.get("Clases", set()) or set()
            if not clases_en_conflicto(clases_c, clases_g): continue
            clases_conf = calcular_clases_conflicto(clases_c, clases_g)
            pdf_id = f"{limpiar_id(c['Marca_Original'])}__vs__{limpiar_id(g['Marca_Original'])}"
            ruta_pdf = os.path.join(PDF_FOLDER, pdf_id + ".pdf")
            c_ser = {k: (list(v) if isinstance(v, set) else v) for k,v in c.items()}
            g_ser = {k: (list(v) if isinstance(v, set) else v) for k,v in g.items()}
            generar_pdf(c, g, score, clases_conf, concepto="", ruta_pdf=ruta_pdf)
            resultados.append({
                "pdf_id":          pdf_id,
                "exp_cliente":     c["Expediente_ID"],
                "marca_cliente":   c["Marca_Original"],
                "titular_cliente": c["Titular"],
                "clases_c":        str(sorted(clases_c)),
                "exp_gaceta":      g["Expediente_ID"],
                "marca_gaceta":    g["Marca_Original"],
                "titular_gaceta":  g["Titular"],
                "clases_g":        str(sorted(clases_g)),
                "score":           score,
                "concepto":        "",
            })
            datos_completos.append({
                "pdf_id":      pdf_id,
                "score":       score,
                "clases_conf": list(clases_conf),
                "c": c_ser,
                "g": g_ser,
            })
    if not resultados:
        flash("No se encontraron coincidencias con las clases vinculadas.")
        return redirect(url_for("inicio"))
    session["resultados_meta"] = resultados
    with open("datos_sesion.json","w",encoding="utf-8") as f:
        json.dump(datos_completos, f, ensure_ascii=False, default=str)
    return redirect(url_for("resultados"))

@app.route("/resultados")
@login_requerido
def resultados():
    metas = session.get("resultados_meta", [])
    if not metas:
        flash("No hay resultados. Sube los archivos primero.")
        return redirect(url_for("inicio"))
    return render_template("resultados.html", resultados=metas, usuario=session["usuario"])

@app.route("/guardar_concepto", methods=["POST"])
@login_requerido
def guardar_concepto():
    data     = request.get_json()
    pdf_id   = data.get("pdf_id","")
    concepto = data.get("concepto","").strip()
    if not pdf_id:
        return jsonify({"ok": False, "msg": "pdf_id faltante"})
    try:
        with open("datos_sesion.json","r",encoding="utf-8") as f:
            todos = json.load(f)
    except:
        return jsonify({"ok": False, "msg": "Sesión expirada, vuelve a procesar."})
    entrada = next((x for x in todos if x["pdf_id"] == pdf_id), None)
    if not entrada:
        return jsonify({"ok": False, "msg": "Registro no encontrado."})
    c = entrada["c"]; g = entrada["g"]
    c["Clases"] = set(); g["Clases"] = set()
    score = entrada["score"]
    clases_conf = set(entrada["clases_conf"])
    ruta_pdf = os.path.join(PDF_FOLDER, pdf_id + ".pdf")
    generar_pdf(c, g, score, clases_conf, concepto=concepto, ruta_pdf=ruta_pdf)
    metas = session.get("resultados_meta", [])
    for m in metas:
        if m["pdf_id"] == pdf_id:
            m["concepto"] = concepto[:120] + "..." if len(concepto)>120 else concepto
    session["resultados_meta"] = metas
    session.modified = True
    return jsonify({"ok": True, "msg": "PDF actualizado con el concepto jurídico."})

@app.route("/descargar_pdf/<pdf_id>")
@login_requerido
def descargar_pdf(pdf_id):
    ruta = os.path.join(PDF_FOLDER, pdf_id + ".pdf")
    if not os.path.exists(ruta):
        flash("PDF no encontrado.")
        return redirect(url_for("resultados"))
    return send_file(ruta, as_attachment=True,
                     download_name=f"Oposicion_{pdf_id[:50]}.pdf")

@app.route("/descargar_excel")
@login_requerido
def descargar_excel():
    metas = session.get("resultados_meta", [])
    if not metas:
        flash("No hay resultados.")
        return redirect(url_for("resultados"))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vigilancia"
    ws.append(["Exp_Cliente","Marca_Cliente","Titular_Cliente","Clases_Cliente",
               "Similitud_%","Exp_Gaceta","Marca_Gaceta","Titular_Gaceta",
               "Clases_Gaceta","Concepto_Juridico"])
    for r in metas:
        ws.append([r["exp_cliente"], r["marca_cliente"], r["titular_cliente"],
                   r["clases_c"], r["score"], r["exp_gaceta"], r["marca_gaceta"],
                   r["titular_gaceta"], r["clases_g"], r["concepto"]])
    buf = BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name="Resultado_Vigilancia.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5000)
