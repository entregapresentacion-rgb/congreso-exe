import os
import re
import sys
import json
import socket
import platform
import threading
import subprocess
import webbrowser
import datetime as dt
import unicodedata
from pathlib import Path
from flask import Flask, render_template_string, request, jsonify

try:
    import openpyxl
except ImportError:
    openpyxl = None

SISTEMA = platform.system()

INVALID_CHARS = set('<>:"/\\|?*') | {chr(i) for i in range(32)}
RESERVED = (
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)
MAX_NAME_BYTES = 240


def obtener_ruta_base():
    if getattr(sys, 'frozen', False):
        ruta = os.path.dirname(sys.executable)
        if SISTEMA == 'Darwin' and ruta.endswith('Contents/MacOS'):
            ruta = os.path.abspath(os.path.join(ruta, '..', '..', '..'))
        return ruta
    return os.path.abspath(os.path.dirname(__file__))


def obtener_carpeta_ppt():
    escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.exists(escritorio):
        escritorio = os.path.expanduser("~")
    return os.path.join(escritorio, "PPT")


BASE_DIR = obtener_ruta_base()
RUTA_PPT = obtener_carpeta_ppt()
RUTA_CONFIG = os.path.join(BASE_DIR, "config_congreso.json")
RUTA_EXCEL = os.path.join(BASE_DIR, "programa.xlsx")


def _cut_bytes(s, n):
    b = s.encode("utf-8")
    if len(b) <= n:
        return s
    return b[:n].decode("utf-8", errors="ignore")


def sanitize_name(s):
    if s is None:
        return "_vacio"
    s = str(s)
    clean = "".join(" " if ch in INVALID_CHARS else ch for ch in s)
    clean = re.sub(r"\s+", " ", clean).strip()
    clean = clean.rstrip(". ")
    if clean.upper() in RESERVED:
        clean += "_"
    clean = _cut_bytes(clean, MAX_NAME_BYTES).rstrip(". ").strip()
    if not clean:
        clean = "_vacio"
    return clean


def truncate50(s):
    if s is None:
        return ""
    return s if len(s) <= 50 else s[:50]


class DateStr(str):
    day = None


def _make_datestr(d):
    s = DateStr(d.strftime("%d/%m/%Y"))
    s.day = d.day
    return s


def _parse_any_date(v):
    try:
        return dt.datetime.fromisoformat(v)
    except ValueError:
        pass
    for fmt in ("%Y%m%d", "%d %B %Y", "%B %d, %Y", "%d %b %Y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(v, fmt)
        except ValueError:
            continue
    raise ValueError("fecha no reconocida")


def parse_dia_nn(value, day_part_index=0):
    if isinstance(value, DateStr) and value.day:
        return "DIA {:02d}".format(value.day)
    if isinstance(value, (dt.datetime, dt.date)):
        return "DIA {:02d}".format(value.day)
    if value is None or str(value).strip() == "":
        return "_sin_fecha"

    v = str(value).strip().strip("\"' ").strip()
    if v.lower() in ("nan", "none", "null", "nat"):
        return "_sin_fecha"

    m = re.search(r"[/\-.\s]", v)
    day = None
    try:
        if m:
            parts = [p for p in re.split(r"[/\-.\s]+", v) if p != ""]
            if len(parts) == 3:
                idx = day_part_index if 0 <= day_part_index < 3 else 0
                day = int(parts[idx])
            elif len(parts) == 1 and re.fullmatch(r"[0-9]{1,2}", parts[0]):
                day = int(parts[0])
        elif re.fullmatch(r"[0-9]{1,2}", v):
            day = int(v)
        else:
            day = _parse_any_date(v).day
    except Exception:
        day = None

    if day is None or day < 1 or day > 31:
        return "DIA " + sanitize_name(v)
    return "DIA {:02d}".format(day)


def _cell_to_str(v):
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return _make_datestr(v)
    if isinstance(v, dt.time):
        return v.strftime("%H:%M")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _norm_key(s):
    txt = "".join(c for c in unicodedata.normalize("NFD", str(s or "").strip().lower()) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", txt).strip()


def cargar_config():
    defecto = {
        "workgroup": "CONGRESO",
        "orden_carpetas": [
            "Sala",
            "Día",
            ["Hora Inicio", "Hora Fin", "Título Sesión es", "Sala", "Día"]
        ],
        "nivel_3_columnas": ["Hora Inicio", "Hora Fin", "Título Sesión es", "Sala", "Día"],
        "separador_nivel_3": " - ",
        "max_chars_nivel_3": 50,
        "equipos": ["ADM1", "ADM2", "SALA_1", "SALA_2"],
        "day_part_index": 0
    }
    if os.path.exists(RUTA_CONFIG):
        try:
            with open(RUTA_CONFIG, "r", encoding="utf-8") as f:
                data = json.load(f)
                defecto.update(data)
        except Exception:
            pass
    return defecto


def load_xlsx_rows(path):
    if openpyxl is None:
        raise RuntimeError("Falta la librería openpyxl.")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        raw = [[_cell_to_str(c) for c in row] for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()

    if not raw:
        return [], []

    header_row = raw[0]
    width = 0
    for i, h in enumerate(header_row):
        if str(h).strip():
            width = i + 1
    if width == 0:
        return [], []

    headers, seen = [], {}
    for i in range(width):
        name = str(header_row[i]).strip() or "Columna {}".format(i + 1)
        if name in seen:
            seen[name] += 1
            name = "{} ({})".format(name, seen[name])
        else:
            seen[name] = 1
        headers.append(name)

    rows = []
    for r in raw[1:]:
        cells = list(r[:width]) + [""] * max(0, width - len(r))
        if not any(str(c).strip() for c in cells):
            continue
        rows.append(dict(zip(headers, cells)))

    return headers, rows


def match_col(headers, candidatos):
    mapa = {_norm_key(h): h for h in headers}
    for cand in candidatos:
        cn = _norm_key(cand)
        if cn in mapa:
            return mapa[cn]
    for cand in candidatos:
        cn = _norm_key(cand)
        for hn, h_real in mapa.items():
            if cn and cn in hn:
                return h_real
    return None


def detectar_columnas_l4(headers):
    palabras = [
        "ponente", "ponentes", "participante", "participantes",
        "speaker", "speakers", "autor", "autores", "orador", "oradores",
        "moderador", "moderadores", "chair", "coordinador", "coordinadores",
        "conferenciante", "conferenciantes", "presentador", "presentadores",
        "disertante", "disertantes", "invitado", "invitados"
    ]
    cols = []
    for h in headers:
        hn = _norm_key(h)
        if any(p in hn for p in palabras):
            cols.append(h)
    return cols


def construir_estructura_programa():
    if not os.path.exists(RUTA_EXCEL):
        return [], "No se encontró programa.xlsx"

    cfg = cargar_config()
    try:
        headers, rows = load_xlsx_rows(RUTA_EXCEL)
    except Exception as e:
        return [], f"Error leyendo programa.xlsx: {str(e)}"

    if not rows:
        return [], "El archivo programa.xlsx no tiene filas de datos."

    orden = cfg.get("orden_carpetas", [])
    n3_cols = cfg.get("nivel_3_columnas", [])
    sep = cfg.get("separador_nivel_3", " - ")
    if sep == "":
        sep = " - "
    max_chars = int(cfg.get("max_chars_nivel_3", 50))
    day_idx = int(cfg.get("day_part_index", 0))

    # Determinar columnas Nivel 1 (Sala), Nivel 2 (Fecha/Día) y Nivel 3 (Combo)
    col_l1 = None
    col_l2 = None
    combo_cols = []

    for item in orden:
        if isinstance(item, str) and item.strip():
            encontrada = match_col(headers, [item.strip()])
            if encontrada:
                if col_l1 is None:
                    col_l1 = encontrada
                elif col_l2 is None:
                    col_l2 = encontrada
        elif isinstance(item, list):
            for c_cfg in item:
                encontrada = match_col(headers, [c_cfg])
                if encontrada and encontrada not in combo_cols:
                    combo_cols.append(encontrada)

    if col_l1 is None:
        col_l1 = match_col(headers, ["sala"]) or headers[0]
    if col_l2 is None:
        col_l2 = match_col(headers, ["dia", "fecha"]) or (headers[1] if len(headers) > 1 else headers[0])
    if not combo_cols:
        for c_cfg in n3_cols:
            encontrada = match_col(headers, [c_cfg])
            if encontrada and encontrada not in combo_cols:
                combo_cols.append(encontrada)

    col_titulo = match_col(headers, ["titulo sesion es", "titulo sesion", "titulo de sesion", "sesion", "titulo"])
    col_h_ini = match_col(headers, ["hora inicio", "inicio"])
    col_h_fin = match_col(headers, ["hora fin", "fin"])
    l4_cols = detectar_columnas_l4(headers)

    # Rellenar celdas vacías/combinadas en columnas de sesión para agrupar bien todas las filas de cada sesión
    cols_ffill = [c for c in [col_l1, col_l2, col_titulo, col_h_ini, col_h_fin] + combo_cols if c]
    ultimos = {c: "" for c in cols_ffill}

    sesiones_map = {}
    orden_sesiones = []

    for r in rows:
        fila_proc = dict(r)
        for c in cols_ffill:
            val_actual = str(fila_proc.get(c, "")).strip()
            if val_actual:
                ultimos[c] = fila_proc.get(c, "")
            elif ultimos[c] != "":
                fila_proc[c] = ultimos[c]

        n1 = sanitize_name(fila_proc.get(col_l1, ""))
        n2 = sanitize_name(parse_dia_nn(fila_proc.get(col_l2, ""), day_part_index=day_idx))
        parts = [truncate50(sanitize_name(fila_proc.get(c, ""))) for c in combo_cols]
        n3 = sanitize_name(sep.join(parts))

        p1 = Path(RUTA_PPT) / n1
        p2 = p1 / n2
        p3 = p2 / n3
        ruta_sesion_str = str(p3)

        if ruta_sesion_str not in sesiones_map:
            orden_sesiones.append(ruta_sesion_str)
            sesiones_map[ruta_sesion_str] = {
                "ruta_sesion": ruta_sesion_str,
                "nivel_1": n1,
                "nivel_2": n2,
                "nivel_3": n3,
                "titulo": str(fila_proc.get(col_titulo, "")).strip() or n3,
                "hora_inicio": str(fila_proc.get(col_h_ini, "")).strip(),
                "hora_fin": str(fila_proc.get(col_h_fin, "")).strip(),
                "participantes": [],
                "_vistas": set()
            }

        ses = sesiones_map[ruta_sesion_str]

        # Extraer cada participante de las columnas marcadas de Nivel 4 (speakers)
        for c4 in l4_cols:
            raw_val = fila_proc.get(c4, "")
            if not str(raw_val).strip():
                continue
            # Separar si en una misma celda vienen varios nombres separados por ; o salto de línea o barra
            candidatos = [x.strip() for x in re.split(r'[;\n|]+', str(raw_val)) if x.strip()]
            for cand in candidatos:
                d = sanitize_name(cand)
                if not d or d == "_vacio":
                    continue
                clave = _norm_key(d)
                if clave not in ses["_vistas"]:
                    ses["_vistas"].add(clave)
                    ruta_p = str(p3 / d)
                    archivos = []
                    if os.path.isdir(ruta_p):
                        archivos = [
                            f for f in sorted(os.listdir(ruta_p))
                            if not f.startswith('.') and os.path.isfile(os.path.join(ruta_p, f))
                        ]
                    ses["participantes"].append({
                        "nombre": d,
                        "columna": c4,
                        "ruta": ruta_p,
                        "archivos": archivos,
                        "cargado": len(archivos) > 0
                    })

    resultado = []
    for k in orden_sesiones:
        s = sesiones_map[k]
        del s["_vistas"]
        resultado.append(s)

    return resultado, None


def crear_carpetas_estructura():
    sesiones, err = construir_estructura_programa()
    if err:
        return False, err, 0

    os.makedirs(RUTA_PPT, exist_ok=True)
    creadas = 0
    for ses in sesiones:
        p3 = Path(ses["ruta_sesion"])
        for part_dir in [p3.parent.parent, p3.parent, p3]:
            if not part_dir.exists():
                part_dir.mkdir(parents=True, exist_ok=True)
                creadas += 1
        for p in ses["participantes"]:
            pp = Path(p["ruta"])
            if not pp.exists():
                pp.mkdir(parents=True, exist_ok=True)
                creadas += 1

    return True, f"Estructura generada en {RUTA_PPT} ({len(sesiones)} sesiones, {creadas} carpetas nuevas).", creadas


def compartir_carpeta_red():
    os.makedirs(RUTA_PPT, exist_ok=True)
    if SISTEMA == "Windows":
        cmd = f'net share PPT="{RUTA_PPT}" /GRANT:Todos,FULL'
        res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if res.returncode == 0 or "ya se ha compartido" in (res.stderr + res.stdout).lower():
            return True, f"Carpeta {RUTA_PPT} compartida en red como \\\\{socket.gethostname()}\\PPT"
        return False, f"Ejecuta como Administrador para compartir en red: {res.stderr.strip()}"
    elif SISTEMA == "Darwin":
        subprocess.run(["chmod", "-R", "777", RUTA_PPT], check=False)
        return True, f"Permisos listos en {RUTA_PPT}. Activa 'Compartir archivos' en Ajustes de macOS."
    return True, f"Carpeta lista en {RUTA_PPT}."


def configurar_equipo_red(nombre_equipo, workgroup):
    nombre_limpio = sanitize_name(nombre_equipo)[:15].upper()
    wg_limpio = sanitize_name(workgroup)[:15].upper()
    if SISTEMA == "Windows":
        ps_script = (
            f"Rename-Computer -NewName '{nombre_limpio}' -Force -ErrorAction SilentlyContinue; "
            f"Add-Computer -WorkGroupName '{wg_limpio}' -Force -ErrorAction SilentlyContinue"
        )
        res = subprocess.run(["powershell", "-Command", ps_script], capture_output=True, text=True)
        if res.returncode == 0:
            return True, f"Equipo renombrado a {nombre_limpio} (Grupo: {wg_limpio}). Reinicia para aplicar."
        return False, "Se requieren permisos de Administrador en Windows para cambiar nombre/grupo de trabajo."
    elif SISTEMA == "Darwin":
        subprocess.run(["scutil", "--set", "ComputerName", nombre_limpio], check=False)
        subprocess.run(["scutil", "--set", "LocalHostName", nombre_limpio], check=False)
        return True, f"Nombre local ajustado a {nombre_limpio}."
    return False, "No soportado en este sistema."


def obtener_info_sistema():
    cfg = cargar_config()
    hostname = socket.gethostname()
    try:
        ip_local = socket.gethostbyname(hostname)
    except Exception:
        ip_local = "127.0.0.1"
    return {
        "sistema": SISTEMA,
        "hostname": hostname,
        "ip": ip_local,
        "workgroup": cfg.get("workgroup", "CONGRESO"),
        "equipos": cfg.get("equipos", ["ADM1", "ADM2"]),
        "ruta_ppt": RUTA_PPT
    }


def escanear_red_congreso():
    cfg = cargar_config()
    equipos = cfg.get("equipos", [])
    resultados = []
    for eq in equipos:
        online = False
        ip_eq = "-"
        try:
            ip_eq = socket.gethostbyname(eq)
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(0.35)
            if s.connect_ex((ip_eq, 445)) == 0 or s.connect_ex((ip_eq, 139)) == 0:
                online = True
            s.close()
        except Exception:
            online = False
        resultados.append({
            "nombre": eq,
            "ip": ip_eq,
            "online": online,
            "ruta_unc": f"\\\\{eq}\\PPT" if SISTEMA == "Windows" else f"smb://{eq}/PPT"
        })
    return resultados


app = Flask(__name__)

HTML_APP = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CongresosUSB | Presentaciones y Red</title>
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
            padding-bottom: 40px;
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
        .valex-subnav {
            background-color: #192236;
            border-bottom: 1px solid var(--border-color);
            padding: 0 24px;
            display: flex;
            gap: 20px;
        }
        .subnav-btn {
            background: none;
            border: none;
            border-bottom: 2px solid transparent;
            padding: 12px 6px;
            color: var(--text-muted);
            font-size: 0.9rem;
            font-weight: 500;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .subnav-btn.active {
            color: #38bdf8;
            border-bottom-color: #0162e8;
        }
        .valex-card {
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);
        }
        .valex-card-header {
            padding: 15px 20px;
            border-bottom: 1px solid var(--border-color);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .valex-card-title { font-size: 0.92rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin: 0; }
        .form-select, .form-control {
            background-color: var(--bg-input) !important;
            border: 1px solid rgba(255, 255, 255, 0.14) !important;
            color: #fff !important;
        }
        .btn-valex-primary {
            background: linear-gradient(to right, #0162e8 0%, #00b9ff 100%);
            border: none;
            color: #fff;
            font-weight: 600;
        }
        .btn-valex-green {
            background: linear-gradient(to right, #029666 0%, #48d6a8 100%);
            border: none;
            color: #fff;
            font-weight: 600;
        }
        .sesion-card {
            background-color: #171f32;
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 16px;
        }
        .part-row {
            background-color: #1f2940;
            border: 1px solid rgba(255,255,255,0.07);
            border-radius: 6px;
            padding: 10px 14px;
            margin-top: 8px;
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
        }
        .file-pill {
            background: rgba(0, 185, 255, 0.15);
            border: 1px solid rgba(0, 185, 255, 0.35);
            color: #38bdf8;
            border-radius: 4px;
            padding: 2px 8px;
            font-size: 0.78rem;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        .table-dark-valex {
            --bs-table-bg: transparent;
            --bs-table-color: #e2e8f0;
            --bs-table-border-color: rgba(255,255,255,0.08);
        }
    </style>
</head>
<body>
    <header class="valex-topbar d-flex justify-content-between align-items-center">
        <a href="/" class="brand-logo">
            <div class="brand-icon"><i class="bi bi-easel2-fill"></i></div>
            <span>Congresos<span style="color:#00b9ff;">USB</span> · Presentaciones y Red</span>
        </a>
        <div class="d-flex gap-2 align-items-center">
            <span class="badge bg-dark border border-secondary px-3 py-2">PC: {{ info.hostname }} ({{ info.ip }})</span>
            <span class="badge bg-dark border border-info text-info px-3 py-2">Grupo: {{ info.workgroup }}</span>
        </div>
    </header>

    <nav class="valex-subnav">
        <button class="subnav-btn active" id="tabBtnCargar" onclick="cambiarTab('cargar')">
            <i class="bi bi-file-earmark-slides-fill"></i> Presentaciones y Participantes
        </button>
        <button class="subnav-btn" id="tabBtnRed" onclick="cambiarTab('red')">
            <i class="bi bi-hdd-network-fill"></i> Equipos en Red y Carpetas Compartidas
        </button>
    </nav>

    <div class="container-fluid px-4 pt-4" style="max-width: 1320px;">
        <div id="alertaEstado" class="alert d-none mb-4"></div>

        <div id="panelCargar">
            <div class="valex-card mb-4">
                <div class="valex-card-header flex-wrap gap-2">
                    <div>
                        <h5 class="valex-card-title">Estructura de Sesiones y Carga de Presentaciones</h5>
                        <small class="text-secondary">Ruta: <code>{{ info.ruta_ppt }} / [Sala] / [DIA NN] / [Nivel 3] / [Participante]</code></small>
                    </div>
                    <div class="d-flex flex-wrap gap-2">
                        <button class="btn btn-valex-primary btn-sm px-3 py-2" onclick="crearCarpetas()">
                            <i class="bi bi-folder-plus me-1"></i> Crear Carpetas de Participantes
                        </button>
                        <button class="btn btn-outline-info btn-sm px-3 py-2" onclick="ejecutarAccion('/api/abrir_ppt')">
                            <i class="bi bi-folder2-open me-1"></i> Abrir Carpeta PPT
                        </button>
                    </div>
                </div>
                <div class="p-3 border-bottom" style="border-color: var(--border-color) !important;">
                    <div class="row g-2 align-items-center">
                        <div class="col-md-3">
                            <select id="filtroSala" class="form-select form-select-sm" onchange="renderizarSesiones()">
                                <option value="">Todas las Salas</option>
                            </select>
                        </div>
                        <div class="col-md-3">
                            <select id="filtroDia" class="form-select form-select-sm" onchange="renderizarSesiones()">
                                <option value="">Todos los Días</option>
                            </select>
                        </div>
                        <div class="col-md-4">
                            <input type="text" id="buscadorPrograma" class="form-control form-control-sm" placeholder="Buscar por participante, sesión o archivo..." oninput="renderizarSesiones()">
                        </div>
                        <div class="col-md-2 text-end">
                            <span class="badge bg-primary px-3 py-2" id="contadorResumen">0 / 0 cargadas</span>
                        </div>
                    </div>
                </div>
                <div class="p-3" id="contenedorSesiones">
                    <div class="text-center text-secondary py-4">Cargando estructura...</div>
                </div>
            </div>
        </div>

        <div id="panelRed" class="d-none">
            <div class="row g-4">
                <div class="col-lg-6">
                    <div class="valex-card h-100">
                        <div class="valex-card-header">
                            <h5 class="valex-card-title">1. Carpetas PPT en Red</h5>
                            <i class="bi bi-share-fill text-info"></i>
                        </div>
                        <div class="p-4">
                            <p class="small text-secondary mb-3">
                                Comparte la carpeta <code>{{ info.ruta_ppt }}</code> en la red local para que los equipos de sala accedan a las presentaciones.
                            </p>
                            <div class="d-grid gap-2">
                                <button class="btn btn-valex-green py-2" onclick="ejecutarAccion('/api/compartir_ppt')">
                                    <i class="bi bi-share-fill me-2"></i> COMPARTIR CARPETA PPT EN RED
                                </button>
                                <button class="btn btn-outline-info py-2" onclick="ejecutarAccion('/api/abrir_ppt')">
                                    <i class="bi bi-folder2-open me-2"></i> ABRIR CARPETA PPT LOCAL
                                </button>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-lg-6">
                    <div class="valex-card h-100">
                        <div class="valex-card-header">
                            <h5 class="valex-card-title">2. Identidad de este Equipo</h5>
                            <i class="bi bi-pc-display text-info"></i>
                        </div>
                        <div class="p-4">
                            <label class="form-label small text-secondary">Rol / Nombre asignado en el congreso:</label>
                            <select id="selectEquipo" class="form-select mb-3">
                                {% for eq in info.equipos %}
                                    <option value="{{ eq }}" {% if eq == info.hostname %}selected{% endif %}>{{ eq }}</option>
                                {% endfor %}
                            </select>
                            <button class="btn btn-valex-primary w-100 py-2" onclick="aplicarIdentidad()">
                                <i class="bi bi-check2-circle me-2"></i> ASIGNAR NOMBRE Y GRUPO ({{ info.workgroup }})
                            </button>
                        </div>
                    </div>
                </div>

                <div class="col-12">
                    <div class="valex-card">
                        <div class="valex-card-header">
                            <h5 class="valex-card-title">3. Equipos del Congreso en Red</h5>
                            <button class="btn btn-sm btn-outline-info" onclick="cargarRed()">
                                <i class="bi bi-arrow-clockwise me-1"></i> Actualizar Red
                            </button>
                        </div>
                        <div class="p-4">
                            <div class="table-responsive">
                                <table class="table table-dark-valex align-middle mb-0">
                                    <thead>
                                        <tr>
                                            <th>Equipo</th>
                                            <th>IP</th>
                                            <th>Estado</th>
                                            <th>Ruta Compartida</th>
                                            <th class="text-end">Acción</th>
                                        </tr>
                                    </thead>
                                    <tbody id="tablaEquipos">
                                        <tr><td colspan="5" class="text-center text-secondary">Escaneando red local...</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

<script>
let sesionesData = [];

function cambiarTab(tab) {
    const esCargar = tab === 'cargar';
    document.getElementById('panelCargar').classList.toggle('d-none', !esCargar);
    document.getElementById('panelRed').classList.toggle('d-none', esCargar);
    document.getElementById('tabBtnCargar').classList.toggle('active', esCargar);
    document.getElementById('tabBtnRed').classList.toggle('active', !esCargar);
    if (!esCargar) cargarRed();
}

function mostrarAlerta(ok, msg) {
    const el = document.getElementById('alertaEstado');
    el.className = 'alert mb-4 ' + (ok ? 'alert-success' : 'alert-danger');
    el.textContent = msg;
    el.classList.remove('d-none');
}

async function ejecutarAccion(url, payload = {}) {
    const res = await fetch(url, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(payload)
    });
    const data = await res.json();
    mostrarAlerta(data.ok, data.mensaje);
    return data;
}

async function crearCarpetas() {
    const res = await ejecutarAccion('/api/crear_carpetas');
    if (res.ok) await cargarPrograma();
}

async function cargarPrograma() {
    const res = await fetch('/api/programa');
    const data = await res.json();
    if (!data.ok) {
        document.getElementById('contenedorSesiones').innerHTML = `<div class="alert alert-warning mb-0">${data.mensaje}</div>`;
        return;
    }
    sesionesData = data.sesiones || [];

    const salas = [...new Set(sesionesData.map(s => s.nivel_1).filter(Boolean))].sort();
    const dias = [...new Set(sesionesData.map(s => s.nivel_2).filter(Boolean))].sort();

    const selSala = document.getElementById('filtroSala');
    const salaPrev = selSala.value;
    selSala.innerHTML = '<option value="">Todas las Salas</option>' + salas.map(s => `<option value="${s}">${s}</option>`).join('');
    selSala.value = salaPrev;

    const selDia = document.getElementById('filtroDia');
    const diaPrev = selDia.value;
    selDia.innerHTML = '<option value="">Todos los Días</option>' + dias.map(d => `<option value="${d}">${d}</option>`).join('');
    selDia.value = diaPrev;

    renderizarSesiones();
}

function renderizarSesiones() {
    const fSala = document.getElementById('filtroSala').value;
    const fDia = document.getElementById('filtroDia').value;
    const q = document.getElementById('buscadorPrograma').value.toLowerCase().trim();

    let totalParts = 0;
    let cargados = 0;
    sesionesData.forEach(s => {
        s.participantes.forEach(p => {
            totalParts++;
            if (p.cargado) cargados++;
        });
    });
    document.getElementById('contadorResumen').textContent = `${cargados} / ${totalParts} con presentación`;

    const filtradas = sesionesData.filter(s => {
        if (fSala && s.nivel_1 !== fSala) return false;
        if (fDia && s.nivel_2 !== fDia) return false;
        if (q) {
            const texto = `${s.nivel_1} ${s.nivel_2} ${s.nivel_3} ${s.titulo} ${s.participantes.map(p => p.nombre + ' ' + p.archivos.join(' ')).join(' ')}`.toLowerCase();
            if (!texto.includes(q)) return false;
        }
        return true;
    });

    const cont = document.getElementById('contenedorSesiones');
    if (filtradas.length === 0) {
        cont.innerHTML = '<div class="text-center text-secondary py-4">No hay sesiones que coincidan con el filtro.</div>';
        return;
    }

    cont.innerHTML = filtradas.map((s, idxS) => `
        <div class="sesion-card">
            <div class="d-flex flex-wrap justify-content-between align-items-start gap-2 mb-2">
                <div>
                    <div class="d-flex flex-wrap gap-2 mb-1">
                        <span class="badge bg-info text-dark">${s.nivel_1}</span>
                        <span class="badge bg-secondary">${s.nivel_2}</span>
                        ${(s.hora_inicio || s.hora_fin) ? `<span class="badge bg-dark border border-secondary">${s.hora_inicio} - ${s.hora_fin}</span>` : ''}
                    </div>
                    <h6 class="fw-bold mb-1 text-white">${s.nivel_3}</h6>
                    <small class="text-secondary font-monospace">${s.ruta_sesion}</small>
                </div>
                <button class="btn btn-sm btn-outline-light" onclick='ejecutarAccion("/api/abrir_ruta", {ruta: ${JSON.stringify(s.ruta_sesion)}})'>
                    <i class="bi bi-folder2-open me-1"></i> Abrir Carpeta Sesión
                </button>
            </div>
            <div>
                ${s.participantes.map((p, idxP) => {
                    const inputId = `file_${idxS}_${idxP}`;
                    return `
                    <div class="part-row">
                        <div>
                            <div class="d-flex align-items-center gap-2">
                                <span class="badge ${p.cargado ? 'bg-success' : 'bg-warning text-dark'}">
                                    ${p.cargado ? '<i class="bi bi-check-circle-fill"></i> Cargado' : '<i class="bi bi-clock"></i> Pendiente'}
                                </span>
                                <span class="fw-semibold text-white"><i class="bi bi-person-badge me-1 text-info"></i>${p.nombre}</span>
                                <small class="text-secondary">(${p.columna})</small>
                            </div>
                            <div class="mt-1 d-flex flex-wrap gap-1">
                                ${p.archivos.length > 0
                                    ? p.archivos.map(a => `
                                        <span class="file-pill">
                                            <i class="bi bi-file-earmark-ppt-fill"></i> ${a}
                                            <button type="button" class="btn btn-link btn-sm p-0 text-info" title="Abrir archivo" onclick='ejecutarAccion("/api/abrir_ruta", {ruta: ${JSON.stringify(p.ruta + "/" + a)}})'>
                                                <i class="bi bi-box-arrow-up-right"></i>
                                            </button>
                                        </span>`).join('')
                                    : '<small class="text-secondary">Sin presentaciones cargadas.</small>'}
                            </div>
                        </div>
                        <div class="d-flex align-items-center gap-2">
                            <input type="file" id="${inputId}" class="d-none" multiple onchange='subirPresentacion(this, ${JSON.stringify(p.ruta)})'>
                            <button class="btn btn-sm btn-valex-primary" onclick="document.getElementById('${inputId}').click()">
                                <i class="bi bi-upload me-1"></i> Cargar PPT
                            </button>
                            <button class="btn btn-sm btn-outline-secondary" title="Abrir carpeta del participante" onclick='ejecutarAccion("/api/abrir_ruta", {ruta: ${JSON.stringify(p.ruta)}})'>
                                <i class="bi bi-folder"></i>
                            </button>
                        </div>
                    </div>`;
                }).join('')}
            </div>
        </div>
    `).join('');
}

async function subirPresentacion(input, rutaDestino) {
    if (!input.files || input.files.length === 0) return;
    const formData = new FormData();
    formData.append('ruta_destino', rutaDestino);
    for (const file of input.files) {
        formData.append('archivos', file);
    }
    const res = await fetch('/api/subir_presentacion', {
        method: 'POST',
        body: formData
    });
    const data = await res.json();
    mostrarAlerta(data.ok, data.mensaje);
    input.value = '';
    if (data.ok) await cargarPrograma();
}

function aplicarIdentidad() {
    const eq = document.getElementById('selectEquipo').value;
    ejecutarAccion('/api/configurar_equipo', { equipo: eq });
}

async function cargarRed() {
    const res = await fetch('/api/red');
    const lista = await res.json();
    const tbody = document.getElementById('tablaEquipos');
    tbody.innerHTML = lista.map(item => `
        <tr>
            <td class="fw-bold">${item.nombre}</td>
            <td><code>${item.ip}</code></td>
            <td>${item.online
                ? '<span class="badge bg-success">En línea</span>'
                : '<span class="badge bg-secondary">Sin respuesta</span>'}</td>
            <td><code>${item.ruta_unc}</code></td>
            <td class="text-end">
                <button class="btn btn-sm btn-outline-light" onclick="ejecutarAccion('/api/abrir_ruta', {ruta: '${item.ruta_unc}'})">
                    Abrir PPT
                </button>
            </td>
        </tr>
    `).join('');
}

cargarPrograma();
</script>
</body>
</html>"""


@app.route('/')
def home():
    return render_template_string(HTML_APP, info=obtener_info_sistema())


@app.route('/api/programa')
def api_programa():
    sesiones, err = construir_estructura_programa()
    if err:
        return jsonify({"ok": False, "mensaje": err, "sesiones": []})
    return jsonify({"ok": True, "sesiones": sesiones})


@app.route('/api/crear_carpetas', methods=['POST'])
def api_crear_carpetas():
    ok, msg, _ = crear_carpetas_estructura()
    return jsonify({"ok": ok, "mensaje": msg})


@app.route('/api/subir_presentacion', methods=['POST'])
def api_subir_presentacion():
    ruta_destino = request.form.get('ruta_destino', '').strip()
    archivos = request.files.getlist('archivos')
    if not ruta_destino or not archivos:
        return jsonify({"ok": False, "mensaje": "Falta la carpeta de destino o el archivo."})

    ruta_abs_ppt = os.path.abspath(RUTA_PPT)
    ruta_abs_dest = os.path.abspath(ruta_destino)
    if not ruta_abs_dest.startswith(ruta_abs_ppt):
        return jsonify({"ok": False, "mensaje": "Ruta de destino no válida."})

    os.makedirs(ruta_abs_dest, exist_ok=True)
    guardados = []
    for f in archivos:
        if f and f.filename:
            base_n, ext_n = os.path.splitext(f.filename)
            nombre_limpio = sanitize_name(base_n) + ext_n
            f.save(os.path.join(ruta_abs_dest, nombre_limpio))
            guardados.append(nombre_limpio)

    return jsonify({
        "ok": True,
        "mensaje": f"Se cargaron {len(guardados)} archivo(s) en {os.path.basename(ruta_abs_dest)}: {', '.join(guardados)}"
    })


@app.route('/api/compartir_ppt', methods=['POST'])
def api_compartir_ppt():
    ok, msg = compartir_carpeta_red()
    return jsonify({"ok": ok, "mensaje": msg})


@app.route('/api/abrir_ppt', methods=['POST'])
def api_abrir_ppt():
    os.makedirs(RUTA_PPT, exist_ok=True)
    if SISTEMA == "Windows":
        os.startfile(RUTA_PPT)
    elif SISTEMA == "Darwin":
        subprocess.Popen(["open", RUTA_PPT])
    return jsonify({"ok": True, "mensaje": f"Abriendo {RUTA_PPT}"})


@app.route('/api/abrir_ruta', methods=['POST'])
def api_abrir_ruta():
    ruta = (request.json or {}).get("ruta", "")
    if not ruta:
        return jsonify({"ok": False, "mensaje": "Ruta vacía."})
    ruta_norm = os.path.normpath(ruta) if not ruta.startswith(("\\\\", "smb://")) else ruta
    if SISTEMA == "Windows":
        os.startfile(ruta_norm)
    elif SISTEMA == "Darwin":
        subprocess.Popen(["open", ruta_norm])
    return jsonify({"ok": True, "mensaje": f"Abriendo {ruta_norm}"})


@app.route('/api/configurar_equipo', methods=['POST'])
def api_configurar_equipo():
    cfg = cargar_config()
    equipo = (request.json or {}).get("equipo", "")
    ok, msg = configurar_equipo_red(equipo, cfg.get("workgroup", "CONGRESO"))
    return jsonify({"ok": ok, "mensaje": msg})


@app.route('/api/red')
def api_red():
    return jsonify(escanear_red_congreso())


if __name__ == '__main__':
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5055")).start()
    app.run(host='127.0.0.1', port=5055)
