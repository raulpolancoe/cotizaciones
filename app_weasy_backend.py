from flask import Flask, request, jsonify, send_file, render_template_string
import pandas as pd
from datetime import datetime
import locale
import os
import unicodedata
import re
from weasyprint import HTML, CSS
from jinja2 import Template
from pathlib import Path

app = Flask(__name__)

# Rutas
EXCEL_PATH = "Lista_Precios.xlsx"
TEMPLATE_HTML = "plantilla_cotizacion.html"

def limpiar_nombre(nombre):
    nombre = unicodedata.normalize('NFKD', nombre).encode('ascii', 'ignore').decode('ascii')
    nombre = re.sub(r'[^a-zA-Z0-9_]', '_', nombre)
    return nombre.strip().replace("__", "_").replace(" ", "_")

@app.route("/")
def index():
    return open("cotizaciones.html", encoding="utf-8").read()

@app.route("/productos")
def productos():
    try:
        df = pd.read_excel(EXCEL_PATH)
        productos = df.to_dict(orient="records")
        return jsonify(productos)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/generar", methods=["POST"])
def generar():
    pdf_path = None  # ← inicializar la variable antes de usarla en finally
    try:
        data = request.get_json()
        nombre = data.get("nombre", "Cliente")
        items = data.get("items", [])

        total = sum(item["valor"] * item["cantidad"] for item in items)

        # Establecer fecha en español
        # try:
        #     locale.setlocale(locale.LC_TIME, "es_CO.utf8")
        # except locale.Error:
        #     locale.setlocale(locale.LC_TIME, "es_CO")
        # fecha = datetime.today().strftime("%d de %B de %Y").replace(" 0", " ")
        import babel.dates
        fecha = babel.dates.format_date(datetime.today(), locale="es", format="d 'de' MMMM 'de' y")


        # Cargar plantilla y renderizar
        with open(TEMPLATE_HTML, encoding="utf-8") as f:
            html_template = f.read()

        rendered_html = render_template_string(
            html_template,
            nombre=nombre,
            fecha=fecha,
            items=items,
            total=total
        )

        # Generar PDF temporal
        safe_name = limpiar_nombre(nombre)
        pdf_path = f"Cotizacion_{safe_name}.pdf"
        # HTML(string=rendered_html).write_pdf(pdf_path)
        base_url = Path(__file__).parent.resolve()
        HTML(string=rendered_html, base_url=base_url).write_pdf(pdf_path)

        return send_file(pdf_path, as_attachment=True, download_name=pdf_path, mimetype="application/pdf")

    finally:
        # Limpieza del PDF generado
        if pdf_path and os.path.exists(pdf_path):
        #if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except:
                pass

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
