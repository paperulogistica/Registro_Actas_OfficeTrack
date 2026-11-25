"""
Base de datos acumulativa (SQLite) del Registro de Actas.

Principios:
 - Nunca se borra nada automáticamente: cada lote subido AGREGA filas.
 - Cada fila guarda, además de los datos del acta, la fecha/hora exacta
   en que el sistema la procesó (PROCESSED_AT_COL), puesta por el propio
   servidor al momento de insertar — el usuario no la escribe ni la edita.
 - Un archivo de origen ya procesado no se vuelve a duplicar si se sube
   dos veces por error (se detecta por hash del contenido del archivo).
"""
import hashlib
import os
import sqlite3
from datetime import datetime, timedelta

from .parser import ALL_COLUMNS, PROCESSED_AT_COL

DB_PATH = os.environ.get("ACTAS_DB_PATH", os.path.join(os.path.dirname(__file__), "data", "actas.db"))

# Cuánto tiempo se conserva el PDF original para poder descargarlo
# renombrado. Pasado este tiempo se borra solo; la fila con los datos
# extraídos NO se borra nunca, solo el archivo.
PDF_TTL_SECONDS = 3600


def _connect():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = _connect()
    cols_sql = ", ".join(f'"{c}" TEXT' for c in ALL_COLUMNS)
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS actas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            {cols_sql},
            file_hash TEXT,
            pdf_hash TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS archivos_procesados (
            file_hash TEXT PRIMARY KEY,
            filename TEXT,
            processed_at TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS pdf_temporal (
            pdf_hash TEXT PRIMARY KEY,
            pdf_bytes BLOB,
            pdf_filename TEXT,
            expira_en TEXT
        )
    """)
    # migración suave: si la tabla "actas" viene de antes de agregar
    # pdf_hash, se la agregamos sin romper lo ya guardado.
    try:
        conn.execute('ALTER TABLE actas ADD COLUMN pdf_hash TEXT')
    except sqlite3.OperationalError:
        pass  # la columna ya existe
    conn.commit()
    conn.close()


def file_already_processed(file_bytes):
    h = hashlib.sha256(file_bytes).hexdigest()
    conn = _connect()
    row = conn.execute("SELECT filename, processed_at FROM archivos_procesados WHERE file_hash = ?", (h,)).fetchone()
    conn.close()
    if row:
        return h, dict(row)
    return h, None


def insert_rows(rows, file_hash, filename):
    """Inserta filas nuevas con la fecha de procesamiento puesta AHORA."""
    if not rows:
        return
    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    conn = _connect()
    cols = ALL_COLUMNS
    placeholders = ", ".join("?" for _ in cols) + ", ?, ?"
    col_names = ", ".join(f'"{c}"' for c in cols) + ", file_hash, pdf_hash"
    for row in rows:
        row = dict(row)
        row[PROCESSED_AT_COL] = now
        pdf_hash = row.get("_pdf_hash", "")
        values = [row.get(c, "") for c in cols] + [file_hash, pdf_hash]
        conn.execute(f'INSERT INTO actas ({col_names}) VALUES ({placeholders})', values)
    conn.execute(
        "INSERT OR REPLACE INTO archivos_procesados (file_hash, filename, processed_at) VALUES (?, ?, ?)",
        (file_hash, filename, now),
    )
    conn.commit()
    conn.close()


def store_pdf_temp(pdf_hash, pdf_bytes, pdf_filename):
    """Guarda (o renueva) el PDF original por PDF_TTL_SECONDS."""
    expira_en = (datetime.now() + timedelta(seconds=PDF_TTL_SECONDS)).isoformat()
    conn = _connect()
    conn.execute(
        "INSERT OR REPLACE INTO pdf_temporal (pdf_hash, pdf_bytes, pdf_filename, expira_en) VALUES (?, ?, ?, ?)",
        (pdf_hash, pdf_bytes, pdf_filename, expira_en),
    )
    conn.commit()
    conn.close()


def _prune_expired_pdfs(conn):
    conn.execute("DELETE FROM pdf_temporal WHERE expira_en < ?", (datetime.now().isoformat(),))


def get_pdf_if_valid(pdf_hash):
    """Devuelve (pdf_bytes, pdf_filename) si el PDF sigue vigente, o None."""
    if not pdf_hash:
        return None
    conn = _connect()
    _prune_expired_pdfs(conn)
    conn.commit()
    row = conn.execute(
        "SELECT pdf_bytes, pdf_filename FROM pdf_temporal WHERE pdf_hash = ?", (pdf_hash,)
    ).fetchone()
    conn.close()
    if not row:
        return None
    return row["pdf_bytes"], row["pdf_filename"]


def get_valid_pdf_hashes():
    """Set de pdf_hash todavía vigentes (para marcar qué filas sí se pueden descargar)."""
    conn = _connect()
    _prune_expired_pdfs(conn)
    conn.commit()
    rows = conn.execute("SELECT pdf_hash FROM pdf_temporal").fetchall()
    conn.close()
    return {r["pdf_hash"] for r in rows}


def get_all_rows():
    conn = _connect()
    rows = conn.execute("SELECT * FROM actas ORDER BY id ASC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def search_rows(query="", solo_pendientes=False, limit=200):
    """
    Busca filas por coincidencia parcial en Nombre, DNI, Cod. Inventario,
    Asunto del correo o Tipo de acta. query vacío = últimas `limit` filas.
    """
    conn = _connect()
    sql = "SELECT * FROM actas WHERE 1=1"
    params = []
    if query:
        like = f"%{query}%"
        sql += (
            ' AND ("Nombre" LIKE ? OR "DNI / CE" LIKE ? OR "Cod. Inventario" LIKE ?'
            ' OR "Asunto del correo" LIKE ? OR "Tipo de acta (formato)" LIKE ?'
            ' OR "Archivo origen" LIKE ?)'
        )
        params += [like, like, like, like, like, like]
    if solo_pendientes:
        sql += ' AND "Revisar" IS NOT NULL AND "Revisar" != \'\''
    sql += " ORDER BY id DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_row(row_id):
    conn = _connect()
    row = conn.execute("SELECT * FROM actas WHERE id = ?", (row_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def update_row(row_id, values):
    """values: dict columna -> nuevo valor (solo de ALL_COLUMNS)."""
    if not values:
        return
    conn = _connect()
    set_sql = ", ".join(f'"{c}" = ?' for c in values)
    conn.execute(f'UPDATE actas SET {set_sql} WHERE id = ?', list(values.values()) + [row_id])
    conn.commit()
    conn.close()


def count_rows():
    conn = _connect()
    n = conn.execute("SELECT COUNT(*) FROM actas").fetchone()[0]
    conn.close()
    return n


def clear_all():
    """
    Borra TODO lo acumulado (filas y el registro de archivos ya
    procesados). Es destructivo e irreversible desde la propia web, así
    que la ruta que llama a esto exige una confirmación explícita antes
    de ejecutarlo. Devuelve cuántas filas se borraron.
    """
    n = count_rows()
    conn = _connect()
    conn.execute("DELETE FROM actas")
    conn.execute("DELETE FROM archivos_procesados")
    conn.execute("DELETE FROM pdf_temporal")
    conn.commit()
    conn.close()
    return n
