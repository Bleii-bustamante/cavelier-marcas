╔══════════════════════════════════════════════════════════════════╗
║       CAVELIER ABOGADOS — Sistema de Vigilancia de Marcas        ║
║       Guía de instalación y uso                                   ║
╚══════════════════════════════════════════════════════════════════╝

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 ESTRUCTURA DE ARCHIVOS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
cavelier_app/
├── app.py                  ← Servidor principal
├── encabezado.png          ← Logo para los PDFs (cópialo aquí)
├── templates/
│   ├── login.html
│   ├── inicio.html
│   └── resultados.html
├── uploads/                ← Se crea automáticamente
├── pdfs_generados/         ← Se crea automáticamente
└── temp_logos/             ← Se crea automáticamente

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 REQUISITOS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Python 3.8 o superior
• Las siguientes librerías (instalar una sola vez):

  pip install flask fpdf2 pandas openpyxl pillow

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 CÓMO ARRANCAR EL SERVIDOR
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. Abre una terminal (cmd o PowerShell en Windows).
2. Navega a la carpeta del proyecto:
     cd ruta/a/cavelier_app
3. Ejecuta:
     python app.py
4. Abre el navegador en:
     http://localhost:5000          ← desde el mismo PC
     http://IP_DEL_SERVIDOR:5000   ← desde otros PCs en la misma red

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 USUARIOS Y CONTRASEÑAS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Los usuarios están definidos en app.py, sección USUARIOS:

  USUARIOS = {
      "cavelier":  "marcas2024",
      "abogado1":  "clave123",
      "abogado2":  "clave456",
  }

Para agregar o cambiar usuarios, edita ese diccionario y reinicia el servidor.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 FLUJO DE USO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. Inicia sesión con usuario y contraseña.
2. Sube el Excel de CLIENTES y el Excel de GACETA.
3. Haz clic en "Analizar y generar reportes".
4. Espera unos segundos (dependiendo del tamaño de los archivos).
5. Verás la tabla con todos los pares en conflicto:
   • Rojo  ≥ 80% similitud
   • Amarillo 70–79%
   • Verde  60–69%
6. Para cada par puedes:
   • Escribir el concepto jurídico → "Guardar y actualizar PDF"
   • Descargar el PDF individual
7. Al terminar, descarga el Excel resumen con todos los pares
   y sus conceptos jurídicos.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 PARA SUBIR AL SERVIDOR DE LA FIRMA
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Opciones recomendadas (de más fácil a más completo):

A) Red interna (más sencillo):
   - Corre python app.py en un PC de la firma.
   - Los demás acceden por http://IP_DEL_PC:5000
   - Solo funciona dentro de la red de la firma.

B) Servidor con dominio propio:
   - Instala en un servidor Linux (Ubuntu).
   - Usa gunicorn + nginx:
       pip install gunicorn
       gunicorn -w 4 -b 0.0.0.0:5000 app:app
   - Configura nginx como proxy inverso.
   - Agrega SSL (certbot) para https.

C) Nube (más accesible desde cualquier lugar):
   - Render.com o Railway.app permiten desplegar
     apps Flask gratis o con bajo costo.
   - Sube el código a GitHub y conecta el repositorio.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 NOTAS IMPORTANTES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Copia tu archivo encabezado.png a la raíz de cavelier_app/
  para que aparezca en los PDFs.
• El umbral de similitud (60%) se puede cambiar en app.py:
    UMBRAL_CORTE = 60
• Los PDFs generados se guardan en pdfs_generados/
  y persisten hasta que reinicies el servidor.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 SOPORTE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Desarrollado con Python + Flask + fpdf2
Cualquier ajuste al diseño, usuarios o lógica se hace
directamente en app.py y los archivos de templates/.
