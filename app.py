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
from functools import wraps

from flask import Flask, render_template, request, send_file, redirect, url_for, flash, Response
import io

from registro_actas import db
from registro_actas.parser import process_uploaded_file, REVIEW_COL, ALL_COLUMNS
from registro_actas.excel_export import build_excel_bytes

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-cambia-esto")

MAX_CONTENT_LENGTH = 60 * 1024 * 1024  # 60 MB por lote subido
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

# Contraseña de acceso (opcional). Si no se define la variable de entorno
# APP_PASSWORD, la app queda abierta sin login (igual que antes).
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")


def _check_auth(password):
    return password == APP_PASSWORD


@app.before_request
def _ensure_db():
    db.init_db()


@app.before_request
def _require_login():
    if not APP_PASSWORD:
        return None  # sin contraseña configurada -> app abierta, como antes
    auth = request.authorization
    if not auth or not _check_auth(auth.password):
        return Response(
            "Acceso restringido. Ingresa la contraseña.", 401,
            {"WWW-Authenticate": 'Basic realm="Registro de Actas"'},
        )
    return None


@app.route("/", methods=["GET"])
def index():
    total = db.count_rows()
    pendientes = len(db.search_rows(solo_pendientes=True, limit=100000))

    q = (request.args.get("q") or "").strip()
    solo_pend = request.args.get("pendientes") == "1"
    if q or solo_pend:
        preview = db.search_rows(query=q, solo_pendientes=solo_pend, limit=200)
    else:
        preview = list(reversed(db.get_all_rows()[-50:]))

    return render_template("index.html", total=total, pendientes=pendientes,
                            preview=preview, q=q, solo_pend=solo_pend)


@app.route("/editar/<int:row_id>", methods=["GET", "POST"])
def editar(row_id):
    row = db.get_row(row_id)
    if not row:
        flash(f"No se encontró la fila #{row_id}.", "error")
        return redirect(url_for("index"))

    if request.method == "POST":
        nuevos = {}
        for col in ALL_COLUMNS:
            if col == "Fecha de procesamiento":
                continue  # la pone el sistema, no se edita a mano
            if col in request.form:
                nuevos[col] = request.form.get(col, "")
        db.update_row(row_id, nuevos)
        flash(f"Fila #{row_id} actualizada.", "success")
        return redirect(url_for("index"))

    return render_template("editar.html", row=row, columns=ALL_COLUMNS)


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
