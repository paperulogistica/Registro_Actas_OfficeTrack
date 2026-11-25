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
import zipfile
from functools import wraps

from flask import Flask, render_template, request, send_file, redirect, url_for, flash, Response
import io

from registro_actas import db
from registro_actas.parser import process_uploaded_file, REVIEW_COL, ALL_COLUMNS, build_renamed_filename
from registro_actas.excel_export import build_excel_bytes

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-cambia-esto")

MAX_CONTENT_LENGTH = 150 * 1024 * 1024  # 150 MB por lote subido
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

    pdfs_vigentes = db.get_valid_pdf_hashes()

    return render_template("index.html", total=total, pendientes=pendientes,
                            preview=preview, q=q, solo_pend=solo_pend,
                            pdfs_vigentes=pdfs_vigentes, pdf_ttl_min=db.PDF_TTL_SECONDS // 60)


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

        rows, warnings, pdf_blobs = process_uploaded_file(filename, content)
        avisos.extend(warnings)
        if rows:
            db.insert_rows(rows, file_hash, filename)
            total_nuevas += len(rows)
            for pdf_hash, (pdf_bytes, pdf_name) in pdf_blobs.items():
                db.store_pdf_temp(pdf_hash, pdf_bytes, pdf_name)
        else:
            # igual se registra el hash para no reintentarlo infinitamente
            # si el archivo no trae ningún Acta reconocible
            db.insert_rows([], file_hash, filename)

    if total_nuevas:
        flash(f"Se agregaron {total_nuevas} fila(s) nueva(s) a la base acumulada. "
              f"Los PDF de este lote se pueden descargar renombrados durante la próxima hora.", "success")
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


@app.route("/pdf/<int:row_id>")
def descargar_pdf(row_id):
    """
    Descarga el PDF del Acta de esa fila, renombrado como
    "Equipo - Serie - Nombre.pdf". Solo funciona dentro de la hora
    siguiente a haberse procesado (PDF_TTL_SECONDS) — pasado ese tiempo
    el archivo ya se borró del servidor; los datos de la fila siguen
    intactos, solo el documento deja de estar disponible.
    """
    row = db.get_row(row_id)
    if not row:
        flash(f"No se encontró la fila #{row_id}.", "error")
        return redirect(url_for("index"))

    pdf = db.get_pdf_if_valid(row.get("pdf_hash"))
    if not pdf:
        flash(f"El PDF de la fila #{row_id} ya no está disponible — se conserva solo "
              f"{db.PDF_TTL_SECONDS // 60} minutos después de procesarse. "
              f"Los datos extraídos siguen intactos; si necesitas el documento, vuelve a subir el correo.", "error")
        return redirect(url_for("index"))

    pdf_bytes, _original_name = pdf
    nombre_descarga = build_renamed_filename(row)
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=nombre_descarga,
    )


@app.route("/pdfs_zip")
def descargar_pdfs_zip():
    """
    Descarga en un .zip todos los PDF todavía vigentes (dentro de su hora)
    que coincidan con el filtro de búsqueda actual (q / pendientes) -- el
    mismo filtro que se esté viendo en la tabla. Sin filtro, intenta con
    todas las filas con PDF vigente.
    """
    q = (request.args.get("q") or "").strip()
    solo_pend = request.args.get("pendientes") == "1"

    if q or solo_pend:
        rows = db.search_rows(query=q, solo_pendientes=solo_pend, limit=100000)
    else:
        rows = db.get_all_rows()

    pdfs_vigentes = db.get_valid_pdf_hashes()
    rows_con_pdf = [r for r in rows if r.get("pdf_hash") and r["pdf_hash"] in pdfs_vigentes]

    if not rows_con_pdf:
        flash(f"No hay PDFs vigentes para descargar ahora mismo (se conservan solo "
              f"{db.PDF_TTL_SECONDS // 60} minutos tras procesarse). Los datos de las filas siguen intactos.", "error")
        return redirect(url_for("index", q=q, pendientes="1" if solo_pend else None))

    buf = io.BytesIO()
    nombres_usados = {}
    pdf_cache = {}  # evita leer de la base el mismo PDF varias veces si varias filas lo comparten
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for row in rows_con_pdf:
            pdf_hash = row["pdf_hash"]
            if pdf_hash not in pdf_cache:
                pdf_cache[pdf_hash] = db.get_pdf_if_valid(pdf_hash)
            pdf = pdf_cache[pdf_hash]
            if not pdf:
                continue  # expiró justo entre la consulta y ahora
            pdf_bytes, _ = pdf
            nombre = build_renamed_filename(row)
            if nombre in nombres_usados:
                nombres_usados[nombre] += 1
                base, ext = nombre.rsplit(".", 1)
                nombre = f"{base} ({nombres_usados[nombre]}).{ext}"
            else:
                nombres_usados[nombre] = 0
            zf.writestr(nombre, pdf_bytes)

    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name="Actas_PDF.zip",
    )


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
