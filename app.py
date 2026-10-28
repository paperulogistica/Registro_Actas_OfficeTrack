"""
Registro de Actas de Entrega/Devolución de Equipos — app web.

- Subes uno o varios .msg / .pdf -> se procesan en memoria (no se guardan
  los archivos originales en el servidor).
- Cada fila extraída se AGREGA a una base de datos acumulativa (SQLite):
  nada se borra ni se sobrescribe entre subidas, y cada fila queda con
  su "Fecha de procesamiento" (puesta automáticamente por el servidor).
- Descargas en cualquier momento el Excel con TODO lo acumulado hasta
  ahora, en el mismo diseño ya validado.
- Sin SharePoint/OneDrive, sin suscripciones de pago.
"""
import os

from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import io

from registro_actas import db
from registro_actas.parser import process_uploaded_file, REVIEW_COL
from registro_actas.excel_export import build_excel_bytes

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-cambia-esto")

MAX_CONTENT_LENGTH = 60 * 1024 * 1024  # 60 MB por lote subido
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


@app.before_request
def _ensure_db():
    db.init_db()


@app.route("/", methods=["GET"])
def index():
    rows = db.get_all_rows()
    total = len(rows)
    pendientes = sum(1 for r in rows if r.get(REVIEW_COL))
    # muestra solo las últimas 50 filas en la vista previa, para no cargar
    # una tabla gigante en el navegador
    preview = rows[-50:]
    return render_template("index.html", total=total, pendientes=pendientes, preview=list(reversed(preview)))


@app.route("/subir", methods=["POST"])
def subir():
    files = request.files.getlist("archivos")
    if not files or all(f.filename == "" for f in files):
        flash("No se seleccionó ningún archivo.", "error")
        return redirect(url_for("index"))

    total_nuevas, total_omitidas, avisos = 0, 0, []

    for f in files:
        if not f.filename:
            continue
        filename = f.filename
        content = f.read()

        file_hash, ya_procesado = db.file_already_processed(content)
        if ya_procesado:
            total_omitidas += 1
            avisos.append(
                f"{filename}: ya había sido procesado el {ya_procesado['processed_at']} "
                f"(como '{ya_procesado['filename']}') — no se duplicó."
            )
            continue

        rows, warnings = process_uploaded_file(filename, content)
        avisos.extend(warnings)
        if rows:
            db.insert_rows(rows, file_hash, filename)
            total_nuevas += len(rows)
        else:
            # igual se registra el hash para no reintentarlo infinitamente
            # si el archivo no trae ningún Acta reconocible
            db.insert_rows([], file_hash, filename)

    if total_nuevas:
        flash(f"Se agregaron {total_nuevas} fila(s) nueva(s) a la base acumulada.", "success")
    if total_omitidas:
        flash(f"{total_omitidas} archivo(s) ya habían sido procesados antes y se omitieron (sin duplicar).", "info")
    for a in avisos:
        flash(a, "warning")

    return redirect(url_for("index"))


@app.route("/limpiar", methods=["POST"])
def limpiar():
    """
    Borra TODO lo acumulado. Es irreversible, así que exige que el
    usuario escriba la palabra BORRAR en el formulario (además de la
    confirmación de JavaScript en el navegador) — un clic accidental en
    el botón no alcanza para borrar nada.
    """
    confirmacion = (request.form.get("confirmacion") or "").strip().upper()
    if confirmacion != "BORRAR":
        flash('Para limpiar la base tienes que escribir exactamente "BORRAR" en el campo de confirmación. No se borró nada.', "error")
        return redirect(url_for("index"))

    n = db.clear_all()
    flash(f"Se borraron {n} fila(s) acumulada(s). La base quedó vacía.", "success")
    return redirect(url_for("index"))


@app.route("/descargar")
def descargar():
    rows = db.get_all_rows()
    excel_bytes = build_excel_bytes(rows)
    return send_file(
        io.BytesIO(excel_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Registro_de_Actas.xlsx",
    )


if __name__ == "__main__":
    db.init_db()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
