import os
import re
import sys
import json
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


def _cut_bytes(s, n=MAX_NAME_BYTES):
    b = s.encode("utf-8")
    if len(b) <= n:
        return s
    return b[:n].decode("utf-8", errors="ignore")


def sanitize_name(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return "_vacio"
    val = str(s)
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


def sanitizar_nombre_carpeta(texto, max_len=None):
    limpio = sanitize_name(texto)
    if max_len is not None and len(limpio) > max_len:
        limpio = sanitize_name(limpio[:max_len])
    return limpio


def formatear_dia_valor(valor, day_choice=0):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "_sin_fecha"
    if isinstance(valor, (dt.datetime, dt.date)):
        return f"DIA {valor.day:02d}"
    texto = str(valor).strip().strip("\"' ").strip()
    if not texto:
        return "_sin_fecha"
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
    if getattr(sys, 'frozen', False):
        ruta = os.path.dirname(sys.executable)
        if SISTEMA == 'Darwin' and ruta.endswith('Contents/MacOS'):
            ruta = os.path.abspath(os.path.join(ruta, '..', '..', '..'))
        return ruta
    return os.path.abspath(os.path.dirname(__file__))


def norm_col(nombre):
    s = str(nombre).lower().strip()
    for a, b in [('á', 'a'), ('é', 'e'), ('í', 'i'), ('ó', 'o'), ('ú', 'u'), ('ñ', 'n')]:
        s = s.replace(a, b)
    return re.sub(r'[^a-z0-9]+', ' ', s).strip()


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

    df = pd.read_excel(ruta_excel, engine='openpyxl').fillna("")

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
        for c in columnas:
            if norm_col(c) == cand:
                return c
    for cand in candidatos:
        for c in columnas:
            if cand in norm_col(c):
                return c
    return None


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

    n3_cols = config.get("nivel_3_columnas")
    if not n3_cols:
        n3_cols = [
            buscar_columna(columnas_raw, ["hora inicio", "hora inicio sesion"]),
            buscar_columna(columnas_raw, ["hora fin", "hora final"]),
            buscar_columna(columnas_raw, ["titulo sesion es", "titulo sesion", "titulo de sesion"]),
            buscar_columna(columnas_raw, ["sala"]),
            buscar_columna(columnas_raw, ["dia", "fecha"])
        ]
        n3_cols = [c for c in n3_cols if c]

    orden_carpetas = config.get("orden_carpetas", [])
    if orden_carpetas:
        for item in orden_carpetas:
            if isinstance(item, list):
                if item and item == config.get("nivel_4_columnas"):
                    continue
                partes.append(_formatear_nivel_3(row, item if item else n3_cols, sep_n3, max_n3))
            elif isinstance(item, str) and item in row:
                if col_dia and item == col_dia and formato_dia_nn:
                    partes.append(sanitize_name(formatear_dia_valor(row[item])))
                else:
                    partes.append(sanitize_name(row[item]))
        if not any(isinstance(x, list) for x in orden_carpetas) and n3_cols:
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


def obtener_participantes_fila(row, columnas, config=None):
    participantes = []
    cols_n4 = (config or {}).get("nivel_4_columnas") or []
    if cols_n4:
        for c in cols_n4:
            if c in row:
                val = str(row[c]).strip()
                if val and sanitize_name(val) != "_vacio" and val not in participantes:
                    participantes.append(val)
        return participantes

    for c in columnas:
        nc = norm_col(c)
        if nc.startswith("participante") or nc.startswith("ponente") or nc.startswith("speaker"):
            val = str(row[c]).strip()
            if val and sanitize_name(val) != "_vacio" and val not in participantes:
                participantes.append(val)
    return participantes


def contar_archivos_en_carpeta(ruta_abs):
    if not os.path.exists(ruta_abs):
        return 0
    total = 0
    for _, _, files in os.walk(ruta_abs):
        for f in files:
            if not f.startswith('.'):
                total += 1
    return total


def ordenar_columnas_solicitadas(columnas):
    grupos_prioridad = [
        ["hora inicio", "hora inicio sesion"],
        ["hora fin", "hora final"],
        ["titulo sesion es", "titulo sesion", "titulo de sesion"],
        ["sala"],
        ["dia", "fecha"],
        ["id"]
    ]
    ordenadas = []
    usadas = set()

    for grupo in grupos_prioridad:
        col_encontrada = buscar_columna([c for c in columnas if c not in usadas], grupo)
        if col_encontrada:
            ordenadas.append(col_encontrada)
            usadas.add(col_encontrada)

    for c in columnas:
        if c not in usadas:
            ordenadas.append(c)
            usadas.add(c)

    return ordenadas


app = Flask(__name__)

HTML_APP = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CongresoApp | Panel de Control</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-body: #141b2d;
            --bg-card: #1f2940;
            --bg-input: #171f32;
            --border-color: rgba(255, 255, 255, 0.08);
            --text-main: #ffffff;
            --text-muted: #8592a6;
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
            border: 1px solid var(--border-color);
            border-radius: 20px;
            padding: 6px 16px;
            color: #fff;
            font-size: 0.88rem;
            width: 300px;
        }
        .valex-search:focus { outline: none; border-color: #00b9ff; }
        .valex-subnav {
            background-color: #192236;
            border-bottom: 1px solid var(--border-color);
            padding: 0 24px;
            display: flex;
            gap: 24px;
            overflow-x: auto;
        }
        .subnav-item {
            padding: 12px 4px;
            color: var(--text-muted);
            font-size: 0.875rem;
            font-weight: 500;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 7px;
            border-bottom: 2px solid transparent;
            white-space: nowrap;
        }
        .subnav-item.active { color: #38bdf8; border-bottom-color: #0162e8; }
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
        .flt-input { min-width: 105px; }
        .btn-valex-blue { background: linear-gradient(to right, #0162e8, #00b9ff); border: none; color: #fff; font-weight: 600; }
        .btn-valex-green { background: linear-gradient(to right, #029666, #48d6a8); border: none; color: #fff; font-weight: 600; }
        .btn-valex-blue:hover, .btn-valex-green:hover { color: #fff; opacity: 0.92; }
        .table-valex { color: #e2e8f0; margin-bottom: 0; font-size: 0.84rem; }
        .table-valex thead tr.header-titles th {
            background-color: #171f32;
            color: #cbd5e1;
            text-transform: uppercase;
            font-size: 0.72rem;
            letter-spacing: 0.5px;
            border-bottom: 1px solid var(--border-color);
            padding: 11px 12px;
            white-space: nowrap;
        }
        .table-valex thead tr.header-filters th {
            background-color: #121829;
            border-bottom: 2px solid rgba(0, 185, 255, 0.25);
            padding: 7px 10px;
        }
        .table-valex tbody td {
            border-bottom: 1px solid var(--border-color);
            padding: 10px 12px;
            vertical-align: middle;
        }
        .table-valex tbody tr:hover { background-color: rgba(255, 255, 255, 0.04); }
        .badge-ppt-ok { background: rgba(2, 150, 102, 0.2); color: #48d6a8; border: 1px solid rgba(72, 214, 168, 0.35); }
        .badge-ppt-wait { background: rgba(247, 106, 45, 0.2); color: #efa65f; border: 1px solid rgba(239, 166, 95, 0.35); }
        .ponente-link {
            display: inline-block;
            background: rgba(1, 98, 232, 0.2);
            border: 1px solid rgba(0, 185, 255, 0.38);
            color: #7dd3fc;
            border-radius: 5px;
            padding: 3px 8px;
            font-size: 0.78rem;
            margin: 2px 3px 2px 0;
            cursor: pointer;
            text-decoration: none;
            white-space: nowrap;
        }
        .ponente-link:hover { background: rgba(0, 185, 255, 0.35); color: #fff; }
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
                <input type="text" id="buscadorPonente" class="valex-search" placeholder="Buscar ponente / participante..." onkeyup="buscarPonente()" onkeydown="if(event.key==='Enter') irACarpetaPonente()">
