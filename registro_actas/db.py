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
from datetime import datetime

from .parser import ALL_COLUMNS, PROCESSED_AT_COL

DB_PATH = os.environ.get("ACTAS_DB_PATH", os.path.join(os.path.dirname(__file__), "data", "actas.db"))


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
            file_hash TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS archivos_procesados (
            file_hash TEXT PRIMARY KEY,
            filename TEXT,
            processed_at TEXT
        )
    """)
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
    placeholders = ", ".join("?" for _ in cols) + ", ?"
    col_names = ", ".join(f'"{c}"' for c in cols) + ", file_hash"
    for row in rows:
        row = dict(row)
        row[PROCESSED_AT_COL] = now
        values = [row.get(c, "") for c in cols] + [file_hash]
        conn.execute(f'INSERT INTO actas ({col_names}) VALUES ({placeholders})', values)
    conn.execute(
        "INSERT OR REPLACE INTO archivos_procesados (file_hash, filename, processed_at) VALUES (?, ?, ?)",
        (file_hash, filename, now),
    )
    conn.commit()
    conn.close()


def get_all_rows():
    conn = _connect()
    rows = conn.execute("SELECT * FROM actas ORDER BY id ASC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


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
    conn.commit()
    conn.close()
    return n
