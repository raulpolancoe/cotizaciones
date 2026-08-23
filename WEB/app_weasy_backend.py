from datetime import date, datetime
from pathlib import Path
import os
import re
import tempfile
import unicodedata

import babel.dates
import pandas as pd
from flask import Flask, jsonify, request, send_file
from jinja2 import Environment, FileSystemLoader, select_autoescape
from num2words import num2words
from weasyprint import HTML

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "Lista_Precios.xlsx"
TEMPLATES = {"cotizacion": "plantilla_cotizacion.html", "cuenta_cobro": "plantilla_cuentacobro.html"}
app = Flask(__name__)
jinja = Environment(loader=FileSystemLoader(BASE_DIR), autoescape=select_autoescape(["html", "xml"]))

def limpiar_nombre(nombre):
    nombre = unicodedata.normalize("NFKD", nombre).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-zA-Z0-9_-]+", "_", nombre).strip("_") or "Cliente"

def numero(valor, campo, minimo=0):
    try: resultado = float(valor)
    except (TypeError, ValueError): raise ValueError(f"El campo {campo} debe ser un número.")
    if resultado < minimo: raise ValueError(f"El campo {campo} no puede ser menor que {minimo}.")
    return resultado

def moneda(valor):
    valor = round(float(valor)); signo = "-" if valor < 0 else ""
    return f"{signo}${abs(valor):,.0f}".replace(",", ".")

def fecha_larga(valor=None):
    if valor:
        try: fecha = datetime.strptime(valor, "%Y-%m-%d").date()
        except ValueError: raise ValueError("La fecha del servicio no es válida.")
    else: fecha = date.today()
    return babel.dates.format_date(fecha, locale="es_CO", format="d 'de' MMMM 'de' y")

jinja.filters["moneda"] = moneda

@app.route("/")
def index(): return (BASE_DIR / "cotizaciones.html").read_text(encoding="utf-8")

@app.route("/productos")
def productos():
    try: return jsonify(pd.read_excel(EXCEL_PATH).fillna("").to_dict(orient="records"))
    except Exception as exc: return jsonify({"error": str(exc)}), 500

@app.route("/generar", methods=["POST"])
def generar():
    pdf_path = None
    try:
        data = request.get_json(silent=True) or {}; tipo = data.get("tipo", "cotizacion")
        if tipo not in TEMPLATES: raise ValueError("El tipo de documento no es válido.")
        nombre = str(data.get("nombre", "")).strip()
        if not nombre: raise ValueError("Ingresa el nombre del cliente.")
        items = []
        for item in data.get("items", []):
            descripcion = str(item.get("descripcion", "")).strip()
            cantidad = numero(item.get("cantidad"), "cantidad", 0.01); valor = numero(item.get("valor"), "valor", float("-inf"))
            if not descripcion: raise ValueError("Todos los artículos deben tener descripción.")
            items.append({"descripcion": descripcion, "cantidad": cantidad, "valor": valor})
        if not items: raise ValueError("Agrega al menos un artículo al documento.")
        total = sum(item["valor"] * item["cantidad"] for item in items); abonos = []
        if tipo == "cuenta_cobro":
            for indice, abono in enumerate(data.get("abonos", []), start=1):
                abonos.append({"etiqueta": str(abono.get("etiqueta", "")).strip() or f"Abono {indice}", "valor": numero(abono.get("valor"), f"abono {indice}", 0.01)})
        total_abonos = sum(a["valor"] for a in abonos)
        contexto = {"nombre": nombre, "fecha": fecha_larga(), "items": items, "total": total, "abonos": abonos, "total_abonos": total_abonos, "saldo": total-total_abonos, "numero_cuenta": str(data.get("numero_cuenta", "")).strip(), "nit_cliente": str(data.get("nit_cliente", "")).strip(), "concepto": str(data.get("concepto", "")).strip(), "fecha_servicio": fecha_larga(data.get("fecha_servicio")) if tipo == "cuenta_cobro" else "", "total_en_letras": num2words(round(total), lang="es").upper()+" PESOS"}
        if tipo == "cuenta_cobro" and any(not contexto[x] for x in ("numero_cuenta", "nit_cliente", "concepto")): raise ValueError("Completa todos los datos de la cuenta de cobro.")
        html = jinja.get_template(TEMPLATES[tipo]).render(**contexto)
        temporal = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf"); pdf_path = temporal.name; temporal.close()
        HTML(string=html, base_url=str(BASE_DIR)).write_pdf(pdf_path)
        prefijo = "Cotizacion" if tipo == "cotizacion" else "Cuenta_de_cobro"
        return send_file(pdf_path, as_attachment=True, download_name=f"{prefijo}_{limpiar_nombre(nombre)}.pdf", mimetype="application/pdf")
    except ValueError as exc: return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        app.logger.exception("No se pudo generar el documento"); return jsonify({"error": f"No se pudo generar el PDF: {exc}"}), 500
    finally:
        if pdf_path and os.path.exists(pdf_path):
            try: os.remove(pdf_path)
            except OSError: pass

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
