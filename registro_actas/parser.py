"""
Parser de Actas de Entrega/Devolución (.msg / .pdf) -> filas de datos.

Toda la lógica es la misma validada en el prototipo de escritorio:
 - Reconoce formularios TI-F-001 (celular) y TI-F-002 (equipo de cómputo).
 - Código no reconocido -> fila marcada para revisión manual, nunca se
   descarta ni se inventa estructura.
 - Una fila por equipo (una acta puede traer varios activos).
 - Entrega/Devolución detectado por POSICIÓN de la "X" (coordenadas).
 - Tabla de activos extraída por posición de columnas (no texto plano),
   para no desordenar celdas que ocupan dos líneas.
 - Campo con patrón sospechoso (texto pegado/duplicado) -> fila marcada
   para revisión en vez de guardarse tal cual.

Trabaja 100% en memoria (bytes in -> dicts out). No escribe nada a disco
salvo archivos temporales que limpia él mismo.
"""
import io
import os
import re
import tempfile

import extract_msg
import pdfplumber

COLUMNS = [
    "Fecha del correo", "Asunto del correo", "Tipo de acta (formato)",
    "Tipo de actividad", "Fecha de entrega/devolución", "Nombre",
    "DNI / CE", "Cargo", "Empresa", "Área o proyecto", "Gerente o Director",
    "Cod. Inventario", "Descripción del activo", "Marca", "Modelo", "Serie",
    "IMEI", "Número", "Observaciones", "Encargado de área (nombre)",
    "Encargado de área (DNI)",
]
REVIEW_COL = "Revisar"
SOURCE_FILE_COL = "Archivo origen"
PROCESSED_AT_COL = "Fecha de procesamiento"

# Orden final tal como se guarda en la base de datos / se exporta a Excel
ALL_COLUMNS = COLUMNS + [REVIEW_COL, SOURCE_FILE_COL, PROCESSED_AT_COL]

FORM_LABELS = {
    "TI-F-001": "TI-F-001 - Acta de entrega de celular",
    "TI-F-002": "TI-F-002 - Acta de entrega de equipo de cómputo",
    "MA-F-028": "MA-F-028 - Acta de entrega/devolución de activos (general)",
}
# formularios cuya tabla de activos no trae IMEI/Número, sino Placa/Estado
# y una columna de observación propia por fila
FORMS_SIN_IMEI = {"MA-F-028"}


# --------------------------------------------------------------------------
# .msg handling
# --------------------------------------------------------------------------

def extract_pdfs_from_msg_bytes(msg_bytes, filename_hint="correo.msg"):
    """
    Devuelve lista de (pdf_bytes, nombre_pdf, fecha_correo, asunto_correo).
    Descarta adjuntos que no son PDF (ej. imágenes de firma de correo).
    """
    results = []
    with tempfile.NamedTemporaryFile(suffix=".msg", delete=False) as tmp:
        tmp.write(msg_bytes)
        tmp_path = tmp.name
    try:
        m = extract_msg.Message(tmp_path)
        try:
            subject = (m.subject or "").strip()
            date = m.date
            fecha_correo = ""
            if date:
                try:
                    fecha_correo = date.strftime("%d/%m/%Y")
                except Exception:
                    fecha_correo = str(date)
            for a in m.attachments:
                name = a.longFilename or a.shortFilename or ""
                if name.lower().endswith(".pdf"):
                    results.append((a.data, os.path.basename(name), fecha_correo, subject))
        finally:
            m.close()
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
    return results


# --------------------------------------------------------------------------
# PDF parsing helpers (position-aware)
# --------------------------------------------------------------------------

def _get_words(page):
    return page.extract_words(use_text_flow=False, keep_blank_chars=False)


def _line_groups(words, tol=2.0):
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        placed = False
        for ln in lines:
            if abs(ln["top"] - w["top"]) <= tol:
                ln["words"].append(w)
                placed = True
                break
        if not placed:
            lines.append({"top": w["top"], "words": [w]})
    for ln in lines:
        ln["words"].sort(key=lambda w: w["x0"])
        ln["text"] = " ".join(w["text"] for w in ln["words"])
    lines.sort(key=lambda ln: ln["top"])
    return lines


def _find_line(lines, *needles):
    for ln in lines:
        if all(n.lower() in ln["text"].lower() for n in needles):
            return ln
    return None


def _detect_x_choice(lines, opt_a, opt_b):
    # Puede haber más de una línea que mencione ambas palabras (p.ej. el
    # título "ACTA DE ENTREGA O DEVOLUCIÓN DE ACTIVOS" también las trae);
    # nos quedamos con la primera que además tenga un token "X" suelto.
    candidatas = [ln for ln in lines
                  if opt_a.lower() in ln["text"].lower() and opt_b.lower() in ln["text"].lower()]
    ln = next((c for c in candidatas
               if any(w["text"].strip().upper() == "X" for w in c["words"])), None)
    if not ln:
        return None
    words = ln["words"]
    xs = [w for w in words if w["text"].strip().upper() == "X"]
    opt_a_w = next((w for w in words if opt_a.lower() in w["text"].lower()), None)
    opt_b_w = next((w for w in words if opt_b.lower() in w["text"].lower()), None)
    if not xs or not opt_a_w or not opt_b_w:
        return None
    best, best_dist = None, None
    for x in xs:
        for label, w in ((opt_a, opt_a_w), (opt_b, opt_b_w)):
            if x["x0"] <= w["x0"]:
                dist = w["x0"] - x["x0"]
                if best_dist is None or dist < best_dist:
                    best_dist = dist
                    best = label
    return best


def _looks_corrupted(value):
    if not value:
        return False
    v = value.replace(" ", "")
    n = len(v)
    if n < 6:
        return False
    for ln in range(5, n // 2 + 1):
        for i in range(0, n - ln):
            frag = v[i:i + ln]
            if frag and v.count(frag) > 1:
                return True
    return False


def _extract_signature_table(lines):
    """
    Extrae la tabla de firmas (sección 7) por POSICIÓN de columnas, en vez
    de texto plano — necesario para formularios como MA-F-028 que no
    traen el párrafo "Yo, ... con DNI...", así que los datos de la
    persona solo están aquí.

    Devuelve una lista de 2 o 3 dicts (uno por columna, de izquierda a
    derecha) con {"dni": ..., "nombre": ..., "cargo": ...}. Lista vacía
    si no se pudo ubicar la tabla.
    """
    dni_lines = [ln for ln in lines if ln["text"].count("DNI:") >= 1]
    if not dni_lines:
        return []
    dni_lines.sort(key=lambda ln: -ln["text"].count("DNI:"))
    dni_line = dni_lines[0]
    anchors = sorted(w["x0"] for w in dni_line["words"]
                      if w["text"].upper().rstrip(":") == "DNI")
    if not anchors:
        return []

    # Los bins van de una etiqueta a la siguiente (no al punto medio): un
    # nombre largo en una columna puede extenderse casi hasta la etiqueta
    # vecina, así que el punto medio corta nombres. Dejamos un pequeño
    # margen antes de la siguiente etiqueta para no comerle su columna.
    bins = []
    for i, x0 in enumerate(anchors):
        start = 0 if i == 0 else anchors[i - 1]
        end = 10_000 if i + 1 >= len(anchors) else anchors[i + 1] - 2
        bins.append((start, end))

    def collect(label_variants):
        candidates = [ln for ln in lines
                      if ln["top"] >= dni_line["top"] - 5
                      and any(lv in ln["text"].upper() for lv in label_variants)]
        if not candidates:
            return ["" for _ in bins]
        candidates.sort(key=lambda ln: ln["top"])
        ln = candidates[0]
        vals = ["" for _ in bins]
        for w in ln["words"]:
            wt = w["text"].upper().rstrip(":")
            if wt in (lv.rstrip(":") for lv in label_variants):
                continue
            for i, (s, e) in enumerate(bins):
                if s <= w["x0"] < e:
                    vals[i] = (vals[i] + " " + w["text"]).strip()
                    break
        return vals

    dni_vals = collect(["DNI:"])
    nombre_vals = collect(["NOMBRE:", "NOMBRES:"])
    cargo_vals = collect(["CARGO:"])
    return [{"dni": dni_vals[i], "nombre": nombre_vals[i], "cargo": cargo_vals[i]}
            for i in range(len(bins))]


def _extract_header_fields(lines, text):
    data = {}
    m = re.search(r"Yo,\s*(.+?)\s+con DNI\s*/\s*CE N[°º]\s*([\d]+)", text, re.S)
    if m:
        data["nombre"] = re.sub(r"\s+", " ", m.group(1)).strip()
        data["dni"] = m.group(2).strip()
    m = re.search(r"cargo de\s+(.+?)\s+señalo", text, re.S)
    if m:
        data["cargo"] = re.sub(r"\s+", " ", m.group(1)).strip()

    ln = _find_line(lines, "P.A PERÚ")
    empresa = None
    if ln:
        opts = ["P.A PERÚ", "P.E.I", "MEPAR", "AGHASO"]
        xs = [w for w in ln["words"] if w["text"].strip().upper() == "X"]
        best, best_dist = None, None
        for label in opts:
            w = next((w for w in ln["words"] if label.split()[0] in w["text"]), None)
            if not w:
                continue
            for x in xs:
                if x["x0"] <= w["x0"]:
                    dist = w["x0"] - x["x0"]
                    if best_dist is None or dist < best_dist:
                        best_dist = dist
                        best = label
        empresa = best
    data["empresa"] = empresa

    m = re.search(r"área o proyecto:\s*(.+)", text)
    if m:
        data["area"] = re.sub(r"\s+", " ", m.group(1)).strip().split("\n")[0]
    m = re.search(r"Gerente o Director:\s*(.+)", text)
    if m:
        data["gerente"] = re.sub(r"\s+", " ", m.group(1)).strip().split("\n")[0]

    m = re.search(r"Fecha de entrega o devoluci[oó]n:\s*([\d/]+)", text)
    if m:
        data["fecha_mov"] = m.group(1).strip()

    data["tipo_actividad"] = _detect_x_choice(lines, "Entrega", "Devolución")

    dni_matches = re.findall(r"DNI:\s*([\d]*)", text)
    nombre_matches = re.findall(r"NOMBRE:\s*(.*?)(?=\s*NOMBRE:|\s*CARGO:|\n|$)", text)
    nombre_matches = [n.strip() for n in nombre_matches]
    if len(dni_matches) >= 2 and dni_matches[-1].strip():
        data["encargado_dni"] = dni_matches[-1].strip()
    if nombre_matches and nombre_matches[-1]:
        data["encargado_nombre"] = nombre_matches[-1]

    m = re.search(r"OBSERVACIONES\s*\n(.*?)(?:FECHAS DE ENTREGA|$)", text, re.S)
    if m:
        obs = re.sub(r"\s+", " ", m.group(1)).strip()
        if obs:
            data["observaciones"] = obs

    return data


def _extract_asset_rows(page, lines):
    words = _get_words(page)
    rows = []

    header_lines = [ln for ln in lines
                    if ("Cod." in ln["text"] and "Inventario" in ln["text"]
                        and "DESCRIPCIÓN" in ln["text"])]
    obs_line = _find_line(lines, "OBSERVACIONES")
    block_bounds = []
    for i, h in enumerate(header_lines):
        start = h["top"]
        if i + 1 < len(header_lines):
            end = header_lines[i + 1]["top"]
        elif obs_line:
            end = obs_line["top"]
        else:
            end = start + 60
        block_bounds.append((h, start, end))

    for header_line, start_top, end_top in block_bounds:
        hw = header_line["words"]
        labels_seen = []
        skip_next = False
        for idx, w in enumerate(hw):
            if skip_next:
                skip_next = False
                continue
            t = w["text"]
            if t == "Cod." and idx + 1 < len(hw) and hw[idx + 1]["text"] == "Inventario":
                labels_seen.append(("cod", w["x0"]))
                skip_next = True
            elif t.upper().startswith("DESCRIP"):
                labels_seen.append(("descripcion", w["x0"]))
            elif t.upper() == "MARCA":
                labels_seen.append(("marca", w["x0"]))
            elif t.upper() == "MODELO":
                labels_seen.append(("modelo", w["x0"]))
            elif t.upper() == "SERIE":
                labels_seen.append(("serie", w["x0"]))
            elif t.upper() == "IMEI":
                labels_seen.append(("imei", w["x0"]))
            elif t.upper().startswith("N") and "MERO" in t.upper():
                labels_seen.append(("numero", w["x0"]))
            elif t.upper().startswith("CARACTER"):
                labels_seen.append(("caracteristicas", w["x0"]))
            elif t.upper().startswith("PLACA"):
                labels_seen.append(("placa", w["x0"]))
            elif t.upper().startswith("ESTADO"):
                labels_seen.append(("estado", w["x0"]))
            elif t.upper().startswith("OBSERVACI") and "ES" not in t.upper():
                # "OBSERVACIÓN" (columna, singular) — distinto de la sección
                # "OBSERVACIONES" (plural) que usan TI-F-001/002 fuera de la tabla
                labels_seen.append(("observacion_col", w["x0"]))

        if not labels_seen:
            continue
        labels_seen.sort(key=lambda t: t[1])
        bins = []
        for i, (name, x0) in enumerate(labels_seen):
            start = 0 if i == 0 else (labels_seen[i - 1][1] + x0) / 2
            end = 10_000 if i + 1 >= len(labels_seen) else (x0 + labels_seen[i + 1][1]) / 2
            bins.append((name, start, end))

        data_words = [w for w in words
                      if start_top + 2 < w["top"] < end_top and w["text"].strip()]
        if not data_words:
            continue

        col_text = {name: [] for name, _, _ in bins}
        for w in data_words:
            for name, x0, x1 in bins:
                if x0 <= w["x0"] < x1:
                    col_text[name].append(w)
                    break

        row = {}
        suspicious = False
        for name, ws in col_text.items():
            ws_sorted = sorted(ws, key=lambda w: (w["top"], w["x0"]))
            val = " ".join(w["text"] for w in ws_sorted).strip()
            val = re.sub(r"\s+", " ", val)
            if name == "cod" and _looks_corrupted(val):
                suspicious = True
            row[name] = val

        if any(row.get(k) for k in ("cod", "descripcion", "marca", "modelo", "serie")):
            row["_sospechoso"] = suspicious
            rows.append(row)

    return rows


def _empty_row():
    return dict.fromkeys(COLUMNS, "")


def process_pdf_bytes(pdf_bytes, fecha_correo, asunto_correo, source_filename):
    """Devuelve (rows, warnings) donde rows son dicts sobre COLUMNS."""
    warnings = []
    out_rows = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        words = _get_words(page)
        lines = _line_groups(words)

        m = re.search(r"C[oó]digo:\s*([A-Z]{2}-[A-Z]-\d{3})", text)
        codigo = m.group(1) if m else None

        base = {"Fecha del correo": fecha_correo, "Asunto del correo": asunto_correo}

        if codigo not in FORM_LABELS:
            row = _empty_row()
            row.update(base)
            row["Tipo de acta (formato)"] = (codigo or "DESCONOCIDO") + " - NO RECONOCIDO (completar a mano)"
            row[REVIEW_COL] = "SI - formulario no reconocido"
            warnings.append(f"{source_filename}: código '{codigo}' no reconocido, fila marcada para revisión manual.")
            out_rows.append(row)
            return out_rows, warnings

        header = _extract_header_fields(lines, text)

        if codigo in FORMS_SIN_IMEI:
            # Estos formularios no traen el párrafo "Yo, ... con DNI...",
            # así que los datos de la persona solo están en la tabla de
            # firmas (columna 2 = quien recibe/entrega, columna 3 =
            # responsable de área, igual que en TI-F-001/002).
            sig = _extract_signature_table(lines)
            if len(sig) >= 2:
                persona = sig[1]
                header["nombre"] = persona.get("nombre") or header.get("nombre", "")
                header["dni"] = persona.get("dni") or header.get("dni", "")
                header["cargo"] = persona.get("cargo") or header.get("cargo", "")
            if len(sig) >= 3:
                encargado = sig[2]
                header["encargado_nombre"] = encargado.get("nombre") or header.get("encargado_nombre", "")
                header["encargado_dni"] = encargado.get("dni") or header.get("encargado_dni", "")

        assets = _extract_asset_rows(page, lines)

        if not assets:
            row = _empty_row()
            row.update(base)
            row["Tipo de acta (formato)"] = FORM_LABELS[codigo] + " (sin activos detectados)"
            row[REVIEW_COL] = "SI - no se detectó tabla de activos"
            warnings.append(f"{source_filename}: no se detectaron activos, revisar manualmente.")
            out_rows.append(row)
            return out_rows, warnings

        for asset in assets:
            row = _empty_row()
            row.update(base)
            row["Tipo de acta (formato)"] = FORM_LABELS[codigo]
            row["Tipo de actividad"] = header.get("tipo_actividad") or ""
            row["Fecha de entrega/devolución"] = header.get("fecha_mov", "")
            row["Nombre"] = header.get("nombre", "")
            row["DNI / CE"] = header.get("dni", "")
            row["Cargo"] = header.get("cargo", "")
            row["Empresa"] = header.get("empresa", "")
            row["Área o proyecto"] = header.get("area", "")
            row["Gerente o Director"] = header.get("gerente", "")
            row["Cod. Inventario"] = asset.get("cod", "")
            row["Descripción del activo"] = asset.get("descripcion", "")
            row["Marca"] = asset.get("marca", "")
            row["Modelo"] = asset.get("modelo", "")
            row["Serie"] = asset.get("serie", "")
            if codigo in FORMS_SIN_IMEI:
                row["IMEI"] = ""
                row["Número"] = ""
                obs_parts = []
                if asset.get("placa"):
                    obs_parts.append(f"Placa: {asset['placa']}")
                if asset.get("estado"):
                    obs_parts.append(f"Estado: {asset['estado']}")
                if asset.get("observacion_col"):
                    obs_parts.append(asset["observacion_col"])
                row["Observaciones"] = " | ".join(obs_parts)
            else:
                row["IMEI"] = asset.get("imei", "")
                row["Número"] = asset.get("numero", "")
                row["Observaciones"] = header.get("observaciones", "")
            row["Encargado de área (nombre)"] = header.get("encargado_nombre", "")
            row["Encargado de área (DNI)"] = header.get("encargado_dni", "")

            flags = []
            if not header.get("tipo_actividad"):
                flags.append("tipo de actividad no detectado")
            if asset.get("_sospechoso"):
                flags.append("Cod. Inventario con patrón sospechoso")
            if not asset.get("cod"):
                flags.append("Cod. Inventario vacío")
            if codigo in FORMS_SIN_IMEI and not header.get("nombre"):
                flags.append("no se detectó el responsable de entrega")
            row[REVIEW_COL] = "; ".join(flags)
            if flags:
                warnings.append(f"{source_filename}: " + "; ".join(flags))
            out_rows.append(row)

    return out_rows, warnings


def process_uploaded_file(filename, file_bytes):
    """
    Punto de entrada único para la web: recibe un archivo subido (.msg o
    .pdf) y devuelve (rows, warnings). rows ya trae 'Archivo origen'
    puesto (la fecha de procesamiento la pone la capa de base de datos,
    no el parser, para que quede fijada en el momento exacto de guardar).
    """
    name_lower = filename.lower()
    all_rows, all_warnings = [], []

    if name_lower.endswith(".msg"):
        pdfs = extract_pdfs_from_msg_bytes(file_bytes, filename)
        if not pdfs:
            all_warnings.append(f"{filename}: no se encontró un PDF adjunto (Acta) dentro del correo.")
            return all_rows, all_warnings
        for pdf_bytes, pdf_name, fecha_correo, asunto in pdfs:
            try:
                rows, warnings = process_pdf_bytes(pdf_bytes, fecha_correo, asunto, f"{filename} -> {pdf_name}")
            except Exception as e:
                all_warnings.append(f"{filename} -> {pdf_name}: error al procesar el PDF ({e}). Revisar manualmente.")
                continue
            for r in rows:
                r[SOURCE_FILE_COL] = filename
            all_rows.extend(rows)
            all_warnings.extend(warnings)
    elif name_lower.endswith(".pdf"):
        try:
            rows, warnings = process_pdf_bytes(file_bytes, "", "", filename)
        except Exception as e:
            all_warnings.append(f"{filename}: error al procesar el PDF ({e}). Revisar manualmente.")
            return all_rows, all_warnings
        for r in rows:
            r[SOURCE_FILE_COL] = filename
        all_rows.extend(rows)
        all_warnings.extend(warnings)
    else:
        all_warnings.append(f"{filename}: extensión no soportada (solo .msg o .pdf).")

    return all_rows, all_warnings
