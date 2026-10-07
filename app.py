from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from io import BytesIO
from datetime import date
import traceback

from derivaciones import procesar_gano, generar_excel_derivaciones, generar_txt_gornitz

app = Flask(__name__)
app.secret_key = "domilab-gornitz-2026"

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/paso1", methods=["POST"])
def paso1():
    """Recibe Excel Gano → devuelve Excel de derivaciones."""
    if "archivo" not in request.files or request.files["archivo"].filename == "":
        flash("Seleccioná el archivo Excel de Gano.", "error")
        return redirect(url_for("index"))

    archivo = request.files["archivo"]
    try:
        filas = procesar_gano(archivo.read())
        if not filas:
            flash("No se encontraron derivaciones en ese archivo.", "warning")
            return redirect(url_for("index"))

        excel_bytes = generar_excel_derivaciones(filas)
        hoy = date.today().strftime("%d%m%Y")
        nombre = f"DERIVACIONES_GORNITZ_{hoy}.xlsx"

        return send_file(
            BytesIO(excel_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=nombre,
        )
    except Exception as e:
        flash(f"Error al procesar el archivo: {e}", "error")
        traceback.print_exc()
        return redirect(url_for("index"))


@app.route("/paso2", methods=["POST"])
def paso2():
    """Recibe Excel con barcodes → devuelve TXT para Gornitz."""
    if "archivo" not in request.files or request.files["archivo"].filename == "":
        flash("Seleccioná el archivo Excel con los barcodes.", "error")
        return redirect(url_for("index"))

    archivo = request.files["archivo"]
    try:
        txt_contenido, n_lineas = generar_txt_gornitz(archivo.read())
        if n_lineas == 0:
            flash("No se encontraron filas con barcode en ese archivo.", "warning")
            return redirect(url_for("index"))

        hoy = date.today().strftime("%d%m%Y")
        nombre = f"788809_{hoy}.txt"

        return send_file(
            BytesIO(txt_contenido.encode("utf-8")),
            mimetype="text/plain",
            as_attachment=True,
            download_name=nombre,
        )
    except Exception as e:
        flash(f"Error al generar el TXT: {e}", "error")
        traceback.print_exc()
        return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5000)
