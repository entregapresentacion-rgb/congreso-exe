import os
import re
import sys
import json
import base64
import traceback
import platform
import threading
import subprocess
import webbrowser
import datetime as dt
import pandas as pd
from flask import Flask, render_template_string, request, jsonify

SISTEMA = platform.system()

INVALID_CHARS = set('<>:"/\\|?*') | {chr(i) for i in range(32)}
RESERVED = (
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)
MAX_NAME_BYTES = 240

ORDEN_COLUMNAS_PRIORITARIAS = [
    ("ESTADO PPT", ["estado ppt", "estado_ppt"]),
    ("CARPETA PONENCIA", ["carpeta ponencia", "carpeta_ponencia"]),
    ("SALA", ["sala"]),
    ("DÍA", ["dia", "fecha"]),
    ("HORA INICIO", ["hora inicio", "hora_inicio", "inicio"]),
    ("HORA FIN", ["hora fin", "hora_fin", "fin"]),
    ("TÍTULO SESIÓN ES", ["titulo sesion es", "titulo sesion", "titulo de sesion", "sesion"]),
]


def _cut_bytes(s, n=MAX_NAME_BYTES):
    b = s.encode("utf-8")
    if len(b) <= n:
        return s
    return b[:n].decode("utf-8", errors="ignore")


def sanitize_name(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return "_vacio"
    val = str(s).replace("\xa0", " ")
    clean = "".join(" " if ch in INVALID_CHARS else ch for ch in val)
    clean = re.sub(r"\s+", " ", clean).strip()
    clean = clean.rstrip(". ")
    if clean.upper() in RESERVED:
        clean += "_"
    clean = _cut_bytes(clean, MAX_NAME_BYTES).rstrip(". ").strip()
    return clean if clean else "_vacio"


def truncate50(s):
    if s is None:
        return ""
    s = str(s)
    return s if len(s) <= 50 else s[:50]


def formatear_dia_valor(valor, day_choice=0):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "_sin_fecha"
    if isinstance(valor, (dt.datetime, dt.date)):
        return f"DIA {valor.day:02d}"
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    texto = str(valor).replace("\xa0", " ").strip().strip("\"' ").strip()
    if not texto:
        return "_sin_fecha"
    if re.fullmatch(r"[0-9]{8}", texto):
        try:
            parsed = dt.datetime.strptime(texto, "%Y%m%d")
            return f"DIA {parsed.day:02d}"
        except ValueError:
            pass
    if re.fullmatch(r"[0-9]{1,2}", texto):
        d = int(texto)
        return f"DIA {d:02d}" if 1 <= d <= 31 else "DIA " + sanitize_name(texto)
    partes = [p for p in re.split(r"[/\-.\s]+", texto) if p != ""]
    if len(partes) == 3 and all(re.fullmatch(r"[0-9]{1,4}", p) for p in partes):
        idx = day_choice if 0 <= day_choice < 3 else 0
        try:
            d = int(partes[idx])
            if 1 <= d <= 31:
                return f"DIA {d:02d}"
        except ValueError:
            pass
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y%m%d", "%d %B %Y", "%B %d, %Y", "%d %b %Y"):
        try:
            parsed = dt.datetime.strptime(texto, fmt)
            return f"DIA {parsed.day:02d}"
        except ValueError:
            continue
    return "DIA " + sanitize_name(texto)


def obtener_ruta_base():
    if getattr(sys, "frozen", False):
        ruta = os.path.dirname(sys.executable)
        if SISTEMA == "Darwin" and ruta.endswith("Contents/MacOS"):
            ruta = os.path.abspath(os.path.join(ruta, "..", "..", ".."))
        return ruta
    return os.path.abspath(os.path.dirname(__file__))


def norm_col(nombre):
    s = str(nombre).replace("\xa0", " ").lower().strip()
    for a, b in [("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ñ", "n")]:
        s = s.replace(a, b)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def cargar_datos_congreso():
    base = obtener_ruta_base()
    posibles_rutas = [base, os.path.join(base, "PAQUETE_USB_CONGRESO")]
    ruta_activa = base
    for r in posibles_rutas:
        if os.path.exists(os.path.join(r, "config_congreso.json")) and os.path.exists(os.path.join(r, "programa.xlsx")):
            ruta_activa = r
            break

    ruta_json = os.path.join(ruta_activa, "config_congreso.json")
    ruta_excel = os.path.join(ruta_activa, "programa.xlsx")

    if not os.path.exists(ruta_json) or not os.path.exists(ruta_excel):
        return None, None, ruta_activa

    with open(ruta_json, "r", encoding="utf-8") as f:
        config = json.load(f)

    df = pd.read_excel(ruta_excel, engine="openpyxl").fillna("")

    ruta_puesto_sesion = os.path.join(ruta_activa, ".puesto_bloqueado_sesion.json")
    puesto_bloqueado = None
    if os.path.exists(ruta_puesto_sesion):
        try:
            with open(ruta_puesto_sesion, "r", encoding="utf-8") as fp:
                puesto_bloqueado = json.load(fp).get("puesto")
        except Exception:
            puesto_bloqueado = None

    config["_puesto_bloqueado"] = puesto_bloqueado
    config["_ruta_puesto_sesion"] = ruta_puesto_sesion
    return config, df, ruta_activa


def buscar_columna(columnas, candidatos):
    for cand in candidatos:
        ncand = norm_col(cand)
        for c in columnas:
            if norm_col(c) == ncand:
                return c
    for cand in candidatos:
        ncand = norm_col(cand)
        for c in columnas:
            if ncand and ncand in norm_col(c):
                return c
    return None


def ordenar_columnas_tabla(columnas_raw, config=None):
    cols = [str(c) for c in columnas_raw]
    ordenadas = []
    usadas = set()

    for etiqueta_defecto, aliases in ORDEN_COLUMNAS_PRIORITARIAS:
        encontrada = buscar_columna([c for c in cols if c not in usadas], aliases)
        if encontrada:
            ordenadas.append(encontrada)
            usadas.add(encontrada)
        else:
            ordenadas.append(etiqueta_defecto)

    if config and isinstance(config.get("orden_columnas_tabla"), list):
        for c_cfg in config["orden_columnas_tabla"]:
            if c_cfg in cols and c_cfg not in usadas:
                ordenadas.append(c_cfg)
                usadas.add(c_cfg)

    for c in cols:
        if c not in usadas:
            ordenadas.append(c)
            usadas.add(c)

    return ordenadas


def resolver_columnas_nivel_3(config, columnas_raw):
    alias_defecto = [
        ["hora inicio", "hora inicio sesion", "inicio"],
        ["hora fin", "hora final", "fin"],
        ["titulo sesion es", "titulo sesion", "titulo de sesion", "sesion", "titulo"],
        ["sala"],
        ["dia", "fecha"],
    ]

    n3_cfg = config.get("nivel_3_columnas") if config else None
    if not n3_cfg and config:
        for item in config.get("orden_carpetas", []):
            if isinstance(item, list) and item:
                first_norm = norm_col(item[0])
                if not any(k in first_norm for k in ("participante", "ponente", "autor", "speaker")):
                    n3_cfg = item
                    break

    resueltas = []
    usadas = set()

    if n3_cfg:
        for etiqueta_o_col in n3_cfg:
            if not etiqueta_o_col:
                continue
            if etiqueta_o_col in columnas_raw and etiqueta_o_col not in usadas:
                resueltas.append(etiqueta_o_col)
                usadas.add(etiqueta_o_col)
                continue
            encontrada = buscar_columna([c for c in columnas_raw if c not in usadas], [etiqueta_o_col])
            if encontrada:
                resueltas.append(encontrada)
                usadas.add(encontrada)

    if not resueltas:
        for grupo in alias_defecto:
            encontrada = buscar_columna([c for c in columnas_raw if c not in usadas], grupo)
            if encontrada:
                resueltas.append(encontrada)
                usadas.add(encontrada)

    return resueltas


def _formatear_nivel_3(row, cols, sep, max_chars):
    partes_limpias = []
    for c in cols:
        if c and c in row:
            raw_val = sanitize_name(row[c])
            partes_limpias.append(truncate50(raw_val)[:max_chars])
    combinado = sep.join(partes_limpias)
    return sanitize_name(combinado)


def construir_ruta_sesion(row, config, columnas_raw):
    partes = ["PPT"]
    sep_n3 = config.get("separador_nivel_3", " - ")
    max_n3 = int(config.get("max_chars_nivel_3", 50))
    formato_dia_nn = bool(config.get("formato_dia_nn", True))

    col_sala = config.get("nivel_1") or buscar_columna(columnas_raw, ["sala"])
    col_dia = config.get("nivel_2") or buscar_columna(columnas_raw, ["dia", "fecha"])
    n3_cols = resolver_columnas_nivel_3(config, columnas_raw)

    orden_carpetas = config.get("orden_carpetas", [])
    if orden_carpetas:
        for item in orden_carpetas:
            if isinstance(item, list):
                first_norm = norm_col(item[0]) if item else ""
                if any(k in first_norm for k in ("participante", "ponente", "autor", "speaker")):
                    continue
                partes.append(_formatear_nivel_3(row, n3_cols, sep_n3, max_n3))
            elif isinstance(item, str) and item in row:
                if col_dia and item == col_dia and formato_dia_nn:
                    partes.append(sanitize_name(formatear_dia_valor(row[item])))
                else:
                    partes.append(sanitize_name(row[item]))
        if n3_cols and not any(
            isinstance(x, list) and not any(k in norm_col(x[0] if x else "") for k in ("participante", "ponente", "autor", "speaker"))
            for x in orden_carpetas
        ):
            partes.append(_formatear_nivel_3(row, n3_cols, sep_n3, max_n3))
    else:
        if col_sala and col_sala in row:
            partes.append(sanitize_name(row[col_sala]))
        if col_dia and col_dia in row:
            val_dia = formatear_dia_valor(row[col_dia]) if formato_dia_nn else row[col_dia]
            partes.append(sanitize_name(val_dia))
        if n3_cols:
            partes.append(_formatear_nivel_3(row, n3_cols, sep_n3, max_n3))

    return os.path.join(*partes)


def _separar_nombres_celda(raw_val):
    texto = str(raw_val or "").replace("\xa0", " ").strip()
    if not texto:
        return []
    candidatos = re.split(r"[;|\n\r]+", texto)
    resultado = []
    for item in candidatos:
        parte = item.strip(" \t\r\n-–—")
        if parte:
            resultado.append(parte)
    return resultado


def obtener_participantes_fila(row, columnas, config=None):
    columnas_objetivo = []
    if config and isinstance(config.get("nivel_4_columnas"), list) and config["nivel_4_columnas"]:
        for col_cfg in config["nivel_4_columnas"]:
            if col_cfg in columnas:
                columnas_objetivo.append(col_cfg)
            else:
                encontrada = buscar_columna(columnas, [col_cfg])
                if encontrada:
                    columnas_objetivo.append(encontrada)

    if not columnas_objetivo:
        for c in columnas:
            nc = norm_col(c)
            if re.fullmatch(r"participante\s*\d*", nc) or re.fullmatch(r"ponente\s*\d*", nc):
                columnas_objetivo.append(c)

    participantes = []
    claves_vistas = set()

    for c in columnas_objetivo:
        val_celda = row[c] if c in row else ""
        for nombre in _separar_nombres_celda(val_celda):
            nombre_san = sanitize_name(nombre)
            if nombre_san == "_vacio":
                continue
            clave = re.sub(r"\s+", " ", nombre_san.lower()).strip()
            if clave not in claves_vistas:
                claves_vistas.add(clave)
                participantes.append(nombre.replace("\xa0", " ").strip())

    return participantes


def contar_archivos_en_carpeta(ruta_abs):
    if not os.path.exists(ruta_abs):
        return 0
    total = 0
    for _, _, files in os.walk(ruta_abs):
        for f in files:
            if not f.startswith("."):
                total += 1
    return total


def ordenar_columnas_por_nivel_3(columnas, n3_cols, config=None):
    base_ordenada = ordenar_columnas_tabla(columnas, config)
    ordenadas = []
    usadas = set()

    for c in n3_cols:
        if c in base_ordenada and c not in usadas:
            ordenadas.append(c)
            usadas.add(c)

    for c in base_ordenada:
        if c not in usadas:
            ordenadas.append(c)
            usadas.add(c)

    return ordenadas


app = Flask(__name__)

HTML_ERROR = """<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Error en Congreso_App</title></head>
<body style="background:#111726;color:#fff;font-family:sans-serif;padding:30px;">
  <h2 style="color:#f87171;">Se produjo un error al cargar Congreso_App</h2>
  <p>Ruta activa detectada: <code>{{ ruta_activa }}</code></p>
  <pre style="background:#1e293b;padding:16px;border-radius:8px;overflow:auto;color:#fca5a5;">{{ detalle_error }}</pre>
</body>
</html>"""

HTML_APP = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CongresoApp | Agenda por Sesiones</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-body: #111726;
            --bg-card: #1f2940;
            --bg-input: #151d30;
            --border-color: rgba(255, 255, 255, 0.08);
            --text-main: #ffffff;
            --text-muted: #94a3b8;
        }
        body {
            background-color: var(--bg-body);
            color: var(--text-main);
            font-family: 'Roboto', -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding-bottom: 50px;
        }
        .valex-topbar {
            background-color: var(--bg-card);
            border-bottom: 1px solid var(--border-color);
            padding: 12px 24px;
        }
        .brand-logo {
            display: flex;
            align-items: center;
            gap: 10px;
            font-weight: 700;
            font-size: 1.25rem;
            color: #fff;
            text-decoration: none;
        }
        .brand-icon {
            width: 32px;
            height: 32px;
            border-radius: 8px;
            background: linear-gradient(135deg, #0162e8, #00b9ff);
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .valex-search {
            background-color: var(--bg-input);
            border: 1px solid rgba(56, 189, 248, 0.25);
            border-radius: 20px;
            padding: 7px 16px;
            color: #fff;
            font-size: 0.88rem;
            width: 360px;
            max-width: 100%;
        }
        .valex-search:focus { outline: none; border-color: #00b9ff; box-shadow: 0 0 0 2px rgba(0, 185, 255, 0.18); }
        .kpi-card {
            border-radius: 8px;
            padding: 18px 20px 36px 20px;
            position: relative;
            overflow: hidden;
            color: #fff;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.25);
            height: 100%;
        }
        .kpi-blue { background: linear-gradient(to right, #0162e8 0%, #00b9ff 100%); }
        .kpi-pink { background: linear-gradient(to right, #f93a5a 0%, #f7778c 100%); }
        .kpi-green { background: linear-gradient(to right, #029666 0%, #48d6a8 100%); }
        .kpi-orange { background: linear-gradient(to right, #f76a2d 0%, #efa65f 100%); }
        .kpi-title { font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.8px; font-weight: 700; opacity: 0.9; margin-bottom: 6px; }
        .kpi-value { font-size: 1.55rem; font-weight: 700; line-height: 1.2; }
        .kpi-sub { font-size: 0.78rem; opacity: 0.85; margin-top: 4px; }
        .kpi-wave { position: absolute; bottom: 0; left: 0; width: 100%; height: 32px; opacity: 0.35; }
        .valex-card {
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);
        }
        .valex-card-header {
            padding: 16px 22px;
            border-bottom: 1px solid var(--border-color);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
        }
        .valex-card-title { font-size: 0.95rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin: 0; }
        .valex-card-subtitle { font-size: 0.8rem; color: var(--text-muted); margin-top: 2px; }
        .form-select, .form-control {
            background-color: var(--bg-input) !important;
            border: 1px solid rgba(255, 255, 255, 0.14) !important;
            color: #fff !important;
            font-size: 0.82rem;
        }
        .btn-valex-blue { background: linear-gradient(to right, #0162e8, #00b9ff); border: none; color: #fff; font-weight: 600; }
        .btn-valex-green { background: linear-gradient(to right, #029666, #48d6a8); border: none; color: #fff; font-weight: 600; }
        .btn-valex-blue:hover, .btn-valex-green:hover { color: #fff; opacity: 0.92; }
        .sesion-bloque {
            border-radius: 10px;
            margin-bottom: 18px;
            overflow: hidden;
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.28);
        }
        .sesion-bloque.tema-0 { background-color: #192642; border: 1px solid rgba(56, 189, 248, 0.35); border-left: 5px solid #00b9ff; }
        .sesion-bloque.tema-0 .sesion-header { background: linear-gradient(90deg, rgba(1, 98, 232, 0.34), rgba(0, 185, 255, 0.12)); }
        .sesion-bloque.tema-1 { background-color: #211e3d; border: 1px solid rgba(168, 85, 247, 0.35); border-left: 5px solid #a855f7; }
        .sesion-bloque.tema-1 .sesion-header { background: linear-gradient(90deg, rgba(126, 34, 206, 0.34), rgba(168, 85, 247, 0.12)); }
        .sesion-bloque.tema-2 { background-color: #162e36; border: 1px solid rgba(45, 212, 191, 0.35); border-left: 5px solid #2dd4bf; }
        .sesion-bloque.tema-2 .sesion-header { background: linear-gradient(90deg, rgba(13, 148, 136, 0.34), rgba(45, 212, 191, 0.12)); }
        .sesion-bloque.tema-3 { background-color: #2d2336; border: 1px solid rgba(244, 114, 182, 0.35); border-left: 5px solid #f472b6; }
        .sesion-bloque.tema-3 .sesion-header { background: linear-gradient(90deg, rgba(190, 24, 93, 0.32), rgba(244, 114, 182, 0.12)); }
        .sesion-header {
            padding: 14px 20px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.08);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
        }
        .sesion-meta-chip {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            background: rgba(15, 23, 42, 0.55);
            border: 1px solid rgba(255, 255, 255, 0.12);
            border-radius: 6px;
            padding: 3px 9px;
            font-size: 0.78rem;
            color: #e2e8f0;
        }
        .sesion-body { padding: 14px 20px; display: flex; flex-direction: column; gap: 10px; }
        .ponente-card {
            background-color: rgba(15, 23, 42, 0.52);
            border: 1px solid rgba(255, 255, 255, 0.09);
            border-radius: 8px;
            padding: 11px 16px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
        }
        .ponente-card:hover { background-color: rgba(15, 23, 42, 0.78); border-color: rgba(56, 189, 248, 0.35); }
        .cartel-estado {
            font-size: 0.78rem;
            font-weight: 700;
            letter-spacing: 0.5px;
            text-transform: uppercase;
            padding: 5px 12px;
            border-radius: 6px;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        .cartel-cargada { background-color: rgba(16, 185, 129, 0.22); color: #6ee7b7; border: 1px solid rgba(52, 211, 153, 0.5); }
        .cartel-pendiente { background-color: rgba(245, 158, 11, 0.22); color: #fcd34d; border: 1px solid rgba(251, 191, 36, 0.5); }
    </style>
</head>
<body>

    <header class="valex-topbar d-flex flex-wrap justify-content-between align-items-center gap-3">
        <div class="d-flex align-items-center gap-3">
            <a href="/" class="brand-logo">
                <div class="brand-icon"><i class="bi bi-easel2-fill"></i></div>
                <span>Congreso<span style="color:#00b9ff;">App</span></span>
            </a>
            <div class="d-flex align-items-center gap-2">
                <input type="text" id="buscadorGlobal" class="valex-search" placeholder="Filtrar por ponente, sesión, sala o día..." onkeyup="filtrarAgenda()" oninput="filtrarAgenda()">
                <button type="button" class="btn btn-sm btn-outline-info px-2 py-1" onclick="limpiarFiltros()" title="Limpiar filtro">
                    <i class="bi bi-x-lg"></i>
                </button>
            </div>
        </div>
        <div class="d-flex align-items-center gap-2">
            <span class="badge bg-dark border border-secondary px-3 py-2">
                <i class="bi bi-hdd-network text-info me-1"></i> {{ config.workgroup if config else 'N/A' }}
            </span>
            <span class="badge bg-dark border border-secondary px-3 py-2">
                <i class="bi bi-shield-lock text-warning me-1"></i>
                {% if puesto_activo %}
                    {{ puesto_activo }} ({{ 'ADMIN' if es_admin else 'SALA - SOLO LECTURA' }})
                {% else %}
                    Sin puesto fijado
                {% endif %}
            </span>
        </div>
    </header>

    <div class="container-fluid px-4 pt-4" style="max-width: 1440px;">

        {% if not config %}
            <div class="valex-card p-5 text-center my-5">
                <i class="bi bi-exclamation-diamond text-warning display-4 mb-3"></i>
                <h4 class="fw-bold">No se encontraron los archivos de configuración</h4>
                <p style="color: var(--text-muted); max-width: 560px; margin: 10px auto 20px;">
                    Coloca <code>programa.xlsx</code> y <code>config_congreso.json</code> en la misma carpeta que <strong>Congreso_App</strong>.
                </p>
                <a href="/" class="btn btn-valex-blue px-4">Reintentar</a>
            </div>
        {% else %}

            <div class="valex-card mb-4">
                <div class="valex-card-header">
                    <div>
                        <h5 class="valex-card-title">Puesto de Trabajo en el Congreso</h5>
                        <div class="valex-card-subtitle">
                            {% if puesto_bloqueado %}
                                Puesto fijado y bloqueado en <strong>{{ puesto_bloqueado }}</strong>. Para cambiarlo, vuelve a abrir el programa.
                            {% else %}
                                Selecciona <strong>ADM1/ADM2</strong> (puede cargar presentaciones) o una <strong>SALA</strong> (solo lectura y queda bloqueada).
                            {% endif %}
                        </div>
                    </div>
                    <div class="d-flex align-items-center gap-2">
                        <select id="selectPuesto" class="form-select form-select-sm" style="width: 260px;" {% if puesto_bloqueado %}disabled{% endif %}>
                            <option value="">-- Seleccionar puesto --</option>
                            {% for eq in config.equipos %}
                                <option value="{{ eq }}" {% if eq == puesto_activo %}selected{% endif %}>
                                    {{ eq }} {{ '(ADMIN - Carga)' if eq.startswith('ADM') else '(SALA - Solo ver)' }}
                                </option>
                            {% endfor %}
                        </select>
                        {% if not puesto_bloqueado %}
                            <button class="btn btn-sm btn-valex-green px-3" onclick="fijarPuesto()">Confirmar y Bloquear</button>
                        {% else %}
                            <span class="badge bg-secondary px-3 py-2"><i class="bi bi-lock-fill me-1"></i> Bloqueado</span>
                        {% endif %}
                    </div>
                </div>
            </div>

            <div class="row g-3 mb-4">
                <div class="col-sm-6 col-xl-3">
                    <div class="kpi-card kpi-blue">
                        <div class="kpi-title">Sesiones Visibles</div>
                        <div class="kpi-value" id="kpiSesiones">{{ sesiones_agrupadas|length }}</div>
                        <div class="kpi-sub">Total sesiones: {{ sesiones_agrupadas|length }}</div>
                        <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,30 L30,22 L60,28 L95,12 L130,24 L165,10 L200,18 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                    </div>
                </div>
                <div class="col-sm-6 col-xl-3">
                    <div class="kpi-card kpi-pink">
                        <div class="kpi-title">Ponentes Visibles</div>
                        <div class="kpi-value" id="kpiPonentes">{{ total_ponentes }}</div>
                        <div class="kpi-sub">Total en programa: {{ total_ponentes }}</div>
                        <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,25 L40,32 L80,15 L120,25 L160,12 L200,22 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                    </div>
                </div>
                <div class="col-sm-6 col-xl-3">
                    <div class="kpi-card kpi-green">
                        <div class="kpi-title">Presentaciones Cargadas</div>
                        <div class="kpi-value">{{ con_archivo }} / {{ total_ponentes }}</div>
                        <div class="kpi-sub">Por ponente en cada sesión</div>
                        <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,32 L35,18 L75,26 L115,10 L155,20 L200,8 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                    </div>
                </div>
                <div class="col-sm-6 col-xl-3">
                    <div class="kpi-card kpi-orange">
                        <div class="kpi-title">Columnas Prioritarias</div>
                        <div class="kpi-value">7 Campos</div>
                        <div class="kpi-sub">ESTADO PPT · CARPETA PONENCIA · SALA · DÍA · HORA INICIO · HORA FIN · TÍTULO SESIÓN ES</div>
                        <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,20 L45,28 L85,14 L130,26 L170,15 L200,24 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                    </div>
                </div>
            </div>

            <div id="alertaAccion" class="alert d-none mb-4"></div>

            <div class="valex-card mb-4">
                <div class="valex-card-header">
                    <div>
                        <h5 class="valex-card-title">Agenda Englobada por Sesiones</h5>
                        <div class="valex-card-subtitle">
                            Cada sesión agrupa su carpeta de ponencia y subcarpetas individuales por participante
                        </div>
                    </div>
                    <div class="d-flex flex-wrap align-items-center gap-2">
                        <select id="flt_dia" class="form-select form-select-sm" style="width: 150px;" onchange="filtrarAgenda()">
                            <option value="">Todos los días</option>
                            {% for d in dias_unicos %}
                                <option value="{{ d|lower }}">{{ d }}</option>
                            {% endfor %}
                        </select>
                        <select id="flt_sala" class="form-select form-select-sm" style="width: 210px;" onchange="filtrarAgenda()" {% if sala_restringida %}disabled{% endif %}>
                            <option value="">Todas las salas</option>
                            {% for s in salas_unicas %}
                                <option value="{{ s|lower }}" {% if sala_restringida and s|lower == sala_restringida|lower %}selected{% endif %}>{{ s }}</option>
                            {% endfor %}
                        </select>
                        <select id="flt_estado" class="form-select form-select-sm" style="width: 150px;" onchange="filtrarAgenda()">
                            <option value="">Estado: Todos</option>
                            <option value="cargado">Cargada</option>
                            <option value="pendiente">Pendiente</option>
                        </select>
                        {% if es_admin %}
                            <button class="btn btn-sm btn-valex-green px-3" onclick="crearArbolCarpetas()">
                                <i class="bi bi-folder-plus me-1"></i> Generar Carpetas
                            </button>
                        {% endif %}
                        <button class="btn btn-sm btn-valex-blue px-3" onclick="limpiarFiltros()">
                            <i class="bi bi-x-circle me-1"></i> Limpiar
                        </button>
                    </div>
                </div>

                <div class="p-3" id="contenedorSesiones">
                    {% for sesion in sesiones_agrupadas %}
                        <div class="sesion-bloque tema-{{ loop.index0 % 4 }}"
                             data-sala="{{ sesion.sala|lower }}"
                             data-dia="{{ sesion.dia|lower }}"
                             data-sesion-texto="{{ sesion.texto_busqueda|lower }}">

                            <div class="sesion-header">
                                <div>
                                    <div class="d-flex flex-wrap align-items-center gap-2 mb-2">
                                        {% for meta in sesion.campos_nivel_3 %}
                                            <span class="sesion-meta-chip">
                                                <strong class="text-info">{{ meta.etiqueta }}:</strong> {{ meta.valor }}
                                            </span>
                                        {% endfor %}
                                    </div>
                                    <div class="small" style="color: #cbd5e1;">
                                        <i class="bi bi-folder2 text-warning me-1"></i>
                                        <code>{{ sesion.ruta_sesion }}</code>
                                    </div>
                                </div>

                                <div class="d-flex align-items-center gap-2">
                                    <span class="badge bg-dark border border-secondary px-2 py-2">
                                        <i class="bi bi-people-fill text-info me-1"></i> {{ sesion.ponentes|length }} participante(s)
                                    </span>
                                    <button type="button" class="btn btn-sm btn-outline-light"
                                            onclick='abrirRuta({{ sesion.ruta_sesion|tojson }}, {{ "true" if es_admin else "false" }})'
                                            title="Abrir carpeta de la ponencia">
                                        <i class="bi bi-folder-symlink me-1"></i> Carpeta Ponencia
                                    </button>
                                </div>
                            </div>

                            <div class="sesion-body">
                                {% for p in sesion.ponentes %}
                                    <div class="ponente-card"
                                         data-estado="{{ 'cargado' if p.tiene_ppt else 'pendiente' }}"
                                         data-ponente-texto="{{ p.texto_busqueda|lower }}">

                                        <div class="d-flex align-items-center gap-3">
                                            <div style="width: 36px; height: 36px; border-radius: 50%; background: rgba(56, 189, 248, 0.16); border: 1px solid rgba(56, 189, 248, 0.35); display: flex; align-items: center; justify-content: center;">
                                                <i class="bi bi-person-fill text-info"></i>
                                            </div>
                                            <div>
                                                <div class="fw-bold text-white" style="font-size: 0.96rem;">{{ p.ponente }}</div>
                                                {% if p.detalle_extra %}
                                                    <div class="small" style="color: var(--text-muted);">{{ p.detalle_extra }}</div>
                                                {% endif %}
                                                <div class="small text-secondary"><code>{{ p.ruta_carpeta }}</code></div>
                                            </div>
                                        </div>

                                        <div class="d-flex align-items-center flex-wrap gap-2">
                                            {% if p.tiene_ppt %}
                                                <span class="cartel-estado cartel-cargada">
                                                    <i class="bi bi-check-circle-fill"></i> CARGADA ({{ p.num_archivos }})
                                                </span>
                                            {% else %}
                                                <span class="cartel-estado cartel-pendiente">
                                                    <i class="bi bi-hourglass-split"></i> PENDIENTE
                                                </span>
                                            {% endif %}

                                            <button type="button" class="btn btn-sm btn-valex-blue px-3"
                                                    onclick='abrirRuta({{ p.ruta_carpeta|tojson }}, {{ "true" if es_admin else "false" }})'>
                                                <i class="bi bi-folder2-open me-1"></i> Subcarpeta
                                            </button>

                                            {% if es_admin %}
                                                <label class="btn btn-sm btn-valex-green mb-0 px-3" title="Subir presentación a la subcarpeta de {{ p.ponente }}">
                                                    <i class="bi bi-upload me-1"></i> Subir
                                                    <input type="file" class="d-none" onchange='subirPresentacion(this, {{ p.ruta_carpeta|tojson }})'>
                                                </label>
                                            {% endif %}
                                        </div>
                                    </div>
                                {% endfor %}
                            </div>
                        </div>
                    {% endfor %}
                </div>
            </div>

        {% endif %}
    </div>

<script>
const SALA_BLOQUEADA = {{ sala_restringida|tojson if sala_restringida else 'null' }};

function mostrarAlerta(texto, ok = true) {
    const el = document.getElementById('alertaAccion');
    el.className = 'alert mb-4';
    el.style.background = ok ? 'rgba(2,150,102,0.2)' : 'rgba(249,58,90,0.2)';
    el.style.border = ok ? '1px solid #48d6a8' : '1px solid #f7778c';
    el.style.color = ok ? '#6ee7b7' : '#fda4af';
    el.innerHTML = `<i class="bi ${ok ? 'bi-check-circle-fill' : 'bi-exclamation-triangle-fill'} me-2"></i>${texto}`;
}

function fijarPuesto() {
    const puesto = document.getElementById('selectPuesto').value;
    if (!puesto) return;
    fetch('/api/fijar-puesto', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ puesto })
    })
    .then(r => r.json())
    .then(data => {
        mostrarAlerta(data.mensaje, data.ok);
        if (data.ok) setTimeout(() => location.reload(), 700);
    });
}

function crearArbolCarpetas() {
    fetch('/api/crear-carpetas', { method: 'POST' })
        .then(r => r.json())
        .then(data => {
            mostrarAlerta(data.mensaje, data.ok);
            if (data.ok) setTimeout(() => location.reload(), 900);
        });
}

function abrirRuta(rutaRelativa, crearSiNoExiste = true) {
    fetch('/api/abrir-ruta', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ruta: rutaRelativa, crear: crearSiNoExiste })
    })
    .then(r => r.json())
    .then(data => {
        if (!data.ok) mostrarAlerta(data.mensaje, false);
    });
}

function subirPresentacion(input, rutaRelativa) {
    if (!input.files || !input.files[0]) return;
    const fd = new FormData();
    fd.append('archivo', input.files[0]);
    fd.append('ruta', rutaRelativa);
    fetch('/api/subir-presentacion', { method: 'POST', body: fd })
        .then(r => r.json())
        .then(data => {
            mostrarAlerta(data.mensaje, data.ok);
            if (data.ok) setTimeout(() => location.reload(), 900);
        });
}

function filtrarAgenda() {
    const q = document.getElementById('buscadorGlobal').value.toLowerCase().trim();
    const fltDia = document.getElementById('flt_dia').value.toLowerCase().trim();
    const fltSala = document.getElementById('flt_sala').value.toLowerCase().trim();
    const fltEstado = document.getElementById('flt_estado').value.toLowerCase().trim();

    let sesionesVisibles = 0;
    let ponentesVisibles = 0;

    document.querySelectorAll('.sesion-bloque').forEach(bloque => {
        const salaBloque = bloque.getAttribute('data-sala') || '';
        const diaBloque = bloque.getAttribute('data-dia') || '';
        const textoSesion = bloque.getAttribute('data-sesion-texto') || '';

        if (SALA_BLOQUEADA && salaBloque !== SALA_BLOQUEADA.toLowerCase()) {
            bloque.style.display = 'none';
            return;
        }
        if (fltSala && salaBloque !== fltSala) {
            bloque.style.display = 'none';
            return;
        }
        if (fltDia && diaBloque !== fltDia) {
            bloque.style.display = 'none';
            return;
        }

        const coincideSesion = !q || textoSesion.includes(q);
        let ponentesEnBloque = 0;

        bloque.querySelectorAll('.ponente-card').forEach(card => {
            const estadoCard = card.getAttribute('data-estado') || '';
            const textoPonente = card.getAttribute('data-ponente-texto') || '';

            const okEstado = !fltEstado || estadoCard === fltEstado;
            const okTexto = coincideSesion || !q || textoPonente.includes(q);

            if (okEstado && okTexto) {
                card.style.display = '';
                ponentesEnBloque++;
                ponentesVisibles++;
            } else {
                card.style.display = 'none';
            }
        });

        if (ponentesEnBloque > 0) {
            bloque.style.display = '';
            sesionesVisibles++;
        } else {
            bloque.style.display = 'none';
        }
    });

    const kpiS = document.getElementById('kpiSesiones');
    const kpiP = document.getElementById('kpiPonentes');
    if (kpiS) kpiS.textContent = sesionesVisibles;
    if (kpiP) kpiP.textContent = ponentesVisibles;
}

function limpiarFiltros() {
    document.getElementById('buscadorGlobal').value = '';
    document.getElementById('flt_dia').value = '';
    if (!SALA_BLOQUEADA) {
        document.getElementById('flt_sala').value = '';
    }
    document.getElementById('flt_estado').value = '';
    filtrarAgenda();
}

filtrarAgenda();
</script>
</body>
</html>
"""


@app.route("/")
def index():
    ruta_activa = obtener_ruta_base()
    try:
        config, df, ruta_activa = cargar_datos_congreso()
        if config is None or df is None:
            return render_template_string(
                HTML_APP,
                config=None,
                puesto_activo=None,
                puesto_bloqueado=None,
                sala_restringida=None,
                es_admin=False,
            )

        columnas_raw = list(df.columns)
        n3_cols_resueltas = resolver_columnas_nivel_3(config, columnas_raw)
        columnas_ordenadas = ordenar_columnas_por_nivel_3(columnas_raw, n3_cols_resueltas, config)

        if n3_cols_resueltas:
            df = df.sort_values(
                by=n3_cols_resueltas,
                key=lambda s: s.astype(str).str.replace("\xa0", " ", regex=False).str.lower().str.strip(),
                kind="mergesort",
            )

        col_sala = config.get("nivel_1") or buscar_columna(columnas_raw, ["sala"])
        col_dia = config.get("nivel_2") or buscar_columna(columnas_raw, ["dia", "fecha"])
        col_titulo_ponencia = buscar_columna(columnas_raw, ["titulo ponencia", "ponencia", "comunicacion", "titulo charla"])
        formato_dia_nn = bool(config.get("formato_dia_nn", True))

        puesto_bloqueado = config.get("_puesto_bloqueado")
        puesto_activo = puesto_bloqueado or ""
        es_admin = str(puesto_activo).upper().startswith("ADM") or puesto_activo == ""
        sala_restringida = None if (not puesto_activo or es_admin) else sanitize_name(puesto_activo)

        sesiones_map = {}
        sesiones_agrupadas = []
        dias_unicos = []
        salas_unicas = []

        for _, row in df.iterrows():
            ruta_sesion = construir_ruta_sesion(row, config, columnas_raw)
            sala_val = sanitize_name(row[col_sala]) if (col_sala and col_sala in row and str(row[col_sala]).strip()) else ""
            if sala_val == "_vacio":
                sala_val = ""
            raw_dia = row[col_dia] if col_dia and col_dia in row else ""
            dia_val = formatear_dia_valor(raw_dia) if (col_dia and formato_dia_nn) else (str(raw_dia).strip() or "Sin Día")

            if dia_val and dia_val not in dias_unicos:
                dias_unicos.append(dia_val)
            if sala_val and sala_val not in salas_unicas:
                salas_unicas.append(sala_val)

            if ruta_sesion not in sesiones_map:
                campos_n3 = []
                for c in (n3_cols_resueltas or columnas_ordenadas[:5]):
                    if c in row:
                        val_c = str(row[c]).replace("\xa0", " ").strip()
                        if val_c:
                            campos_n3.append({"etiqueta": c, "valor": val_c})

                meta_textos = [f"{m['etiqueta']} {m['valor']}" for m in campos_n3]
                texto_sesion = " ".join([ruta_sesion, sala_val, dia_val] + meta_textos)
                sesion_obj = {
                    "ruta_sesion": ruta_sesion,
                    "sala": sala_val,
                    "dia": dia_val,
                    "campos_nivel_3": campos_n3,
                    "texto_busqueda": texto_sesion,
                    "ponentes": [],
                    "_ponentes_vistos": set(),
                }
                sesiones_map[ruta_sesion] = sesion_obj
                sesiones_agrupadas.append(sesion_obj)
            else:
                sesion_obj = sesiones_map[ruta_sesion]

            participantes = obtener_participantes_fila(row, columnas_raw, config) or ["Sin ponente"]
            detalle_extra = str(row[col_titulo_ponencia]).replace("\xa0", " ").strip() if (col_titulo_ponencia and col_titulo_ponencia in row) else ""
            datos_fila_txt = " ".join(str(row[c]).replace("\xa0", " ") for c in columnas_raw if c in row)

            for p in participantes:
                p_san = sanitize_name(p)
                clave_p = re.sub(r"\s+", " ", p_san.lower()).strip()
                if clave_p in sesion_obj["_ponentes_vistos"]:
                    continue
                sesion_obj["_ponentes_vistos"].add(clave_p)

                ruta_carpeta = os.path.join(ruta_sesion, p_san) if p_san != "_vacio" else ruta_sesion
                ruta_abs = os.path.join(ruta_activa, ruta_carpeta)
                num_archivos = contar_archivos_en_carpeta(ruta_abs)
                tiene_ppt = num_archivos > 0

                sesion_obj["ponentes"].append({
                    "ponente": p,
                    "detalle_extra": detalle_extra,
                    "ruta_carpeta": ruta_carpeta,
                    "tiene_ppt": tiene_ppt,
                    "num_archivos": num_archivos,
                    "texto_busqueda": f"{p} {detalle_extra} {datos_fila_txt}",
                })

        total_ponentes = sum(len(s["ponentes"]) for s in sesiones_agrupadas)
        con_archivo = sum(1 for s in sesiones_agrupadas for p in s["ponentes"] if p["tiene_ppt"])

        return render_template_string(
            HTML_APP,
            config=config,
            n3_cols_resueltas=n3_cols_resueltas,
            columnas_tabla=ordenar_columnas_tabla(columnas_raw, config),
            sesiones_agrupadas=sesiones_agrupadas,
            total_ponentes=total_ponentes,
            con_archivo=con_archivo,
            dias_unicos=dias_unicos,
            salas_unicas=salas_unicas,
            puesto_activo=puesto_activo,
            puesto_bloqueado=puesto_bloqueado,
            sala_restringida=sala_restringida,
            es_admin=es_admin,
        )
    except Exception:
        return render_template_string(
            HTML_ERROR,
            ruta_activa=ruta_activa,
            detalle_error=traceback.format_exc(),
        ), 500


@app.route("/api/fijar-puesto", methods=["POST"])
def api_fijar_puesto():
    config, _, _ = cargar_datos_congreso()
    if config is None:
        return jsonify({"ok": False, "mensaje": "No hay configuración cargada."})
    puesto = (request.json or {}).get("puesto", "").strip()
    if puesto not in config.get("equipos", []):
        return jsonify({"ok": False, "mensaje": "Puesto no válido."})

    with open(config["_ruta_puesto_sesion"], "w", encoding="utf-8") as fp:
        json.dump({"puesto": puesto}, fp)

    if not puesto.upper().startswith("ADM"):
        return jsonify({"ok": True, "mensaje": f"Sala '{puesto}' fijada y bloqueada (solo lectura)."})
    return jsonify({"ok": True, "mensaje": f"Puesto ADMIN '{puesto}' activado."})


@app.route("/api/crear-carpetas", methods=["POST"])
def api_crear_carpetas():
    config, df, ruta_activa = cargar_datos_congreso()
    if config is None:
        return jsonify({"ok": False, "mensaje": "Faltan archivos de configuración."})
    puesto = config.get("_puesto_bloqueado") or ""
    if puesto and not puesto.upper().startswith("ADM"):
        return jsonify({"ok": False, "mensaje": "Solo los puestos ADMIN pueden crear o modificar carpetas."})

    columnas_raw = list(df.columns)
    creadas = 0
    rutas_creadas = set()

    for _, row in df.iterrows():
        ruta_sesion = construir_ruta_sesion(row, config, columnas_raw)
        ruta_abs = os.path.join(ruta_activa, ruta_sesion)
        if ruta_abs not in rutas_creadas and not os.path.exists(ruta_abs):
            os.makedirs(ruta_abs, exist_ok=True)
            rutas_creadas.add(ruta_abs)
            creadas += 1

        for p in obtener_participantes_fila(row, columnas_raw, config):
            p_san = sanitize_name(p)
            if p_san == "_vacio":
                continue
            p_abs = os.path.join(ruta_abs, p_san)
            if p_abs not in rutas_creadas and not os.path.exists(p_abs):
                os.makedirs(p_abs, exist_ok=True)
                rutas_creadas.add(p_abs)
                creadas += 1

    return jsonify({"ok": True, "mensaje": f"Carpetas de ponencia y subcarpetas por participante listas ({creadas} nuevas)."})


@app.route("/api/abrir-ruta", methods=["POST"])
def api_abrir_ruta():
    config, _, ruta_activa = cargar_datos_congreso()
    data = request.json or {}
    ruta_rel = data.get("ruta", "PPT")
    crear = bool(data.get("crear", True))
    puesto = config.get("_puesto_bloqueado") or "" if config else ""
    es_admin = (not puesto) or puesto.upper().startswith("ADM")

    ruta_abs = os.path.join(ruta_activa, ruta_rel)
    if crear and es_admin:
        os.makedirs(ruta_abs, exist_ok=True)
    elif not os.path.exists(ruta_abs):
        return jsonify({"ok": False, "mensaje": "La carpeta aún no existe en disco."})

    if SISTEMA == "Windows":
        os.startfile(ruta_abs)
    elif SISTEMA == "Darwin":
        subprocess.Popen(["open", ruta_abs])
    else:
        subprocess.Popen(["xdg-open", ruta_abs])
    return jsonify({"ok": True})


@app.route("/api/subir-presentacion", methods=["POST"])
def api_subir_presentacion():
    config, _, ruta_activa = cargar_datos_congreso()
    puesto = config.get("_puesto_bloqueado") or "" if config else ""
    if puesto and not puesto.upper().startswith("ADM"):
        return jsonify({"ok": False, "mensaje": "Acceso denegado: las salas solo pueden ver presentaciones."})

    archivo = request.files.get("archivo")
    ruta_rel = request.form.get("ruta", "PPT")
    if not archivo or not archivo.filename:
        return jsonify({"ok": False, "mensaje": "No se seleccionó ningún archivo."})

    carpeta_destino = os.path.join(ruta_activa, ruta_rel)
    os.makedirs(carpeta_destino, exist_ok=True)
    nombre_base, ext = os.path.splitext(archivo.filename)
    nombre_seguro = sanitize_name(nombre_base) + sanitize_name(ext).replace("_vacio", "")
    archivo.save(os.path.join(carpeta_destino, nombre_seguro))
    return jsonify({"ok": True, "mensaje": f"Presentación '{nombre_seguro}' cargada en {ruta_rel}."})


def limpiar_bloqueo_al_salir(ruta_archivo_bloqueo):
    try:
        if ruta_archivo_bloqueo and os.path.exists(ruta_archivo_bloqueo):
            os.remove(ruta_archivo_bloqueo)
    except Exception:
        pass


if __name__ == "__main__":
    cfg, _, _ = cargar_datos_congreso()
    ruta_bloq = cfg.get("_ruta_puesto_sesion") if cfg else None
    limpiar_bloqueo_al_salir(ruta_bloq)
    try:
        threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:5051")).start()
    except Exception:
        pass
    try:
        app.run(host="127.0.0.1", port=5051, debug=False, use_reloader=False)
    finally:
        limpiar_bloqueo_al_salir(ruta_bloq)
