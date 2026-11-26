import os
import re
import sys
import json
import time
import socket
import platform
import threading
import subprocess
import webbrowser
import unicodedata
import pandas as pd
from datetime import datetime, date
from flask import Flask, render_template_string, request, jsonify

SISTEMA = platform.system()


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


def normalizar_clave(texto):
    s = str(texto or '').strip().lower()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^a-z0-9]+', ' ', s).strip()


def formatear_valor_celda(val, es_hora=False):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, datetime):
        if es_hora or (val.year == 1899 and val.month == 12):
            return val.strftime("%H:%M")
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.strftime("%Y-%m-%d")
    if hasattr(val, "strftime"):
        return val.strftime("%H:%M")
    texto = str(val).strip()
    if es_hora and re.match(r'^\d{1,2}:\d{2}:\d{2}$', texto):
        return texto[:5]
    return texto


def sanitizar_nombre_carpeta(texto, max_len=None):
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        valor = "Sin_Especificar"
    else:
        valor = str(texto).strip()

    valor = re.sub(r'[\x00-\x1f\x7f]', '', valor)
    valor = re.sub(r'[\\/:*?"<>|]+', '-', valor)
    valor = re.sub(r'\s+', ' ', valor).strip(' .-_')

    if not valor:
        valor = "Sin_Especificar"

    if max_len is not None and len(valor) > max_len:
        valor = valor[:max_len].rstrip(' .-_') or "Sesion"

    return valor


def cargar_config():
    defecto = {
        "workgroup": "CONGRESO",
        "orden_carpetas": [
            "SALA",
            "DÍA",
            ["DÍA", "SALA", "HORA INICIO", "HORA FIN", "TÍTULO SESIÓN"]
        ],
        "nivel_3_columnas": ["DÍA", "SALA", "HORA INICIO", "HORA FIN", "TÍTULO SESIÓN"],
        "separador_nivel_3": " - ",
        "max_chars_nivel_3": 50,
        "equipos": ["ADM1", "ADM2", "SALA_1", "SALA_2"]
    }
    if os.path.exists(RUTA_CONFIG):
        try:
            with open(RUTA_CONFIG, "r", encoding="utf-8") as f:
                data = json.load(f)
                defecto.update(data)
        except Exception:
            pass
    return defecto


def detectar_columna(df_cols, candidatos):
    mapa = {normalizar_clave(c): c for c in df_cols}
    for cand in candidatos:
        cand_norm = normalizar_clave(cand)
        if cand_norm in mapa:
            return mapa[cand_norm]
    for cand in candidatos:
        cand_norm = normalizar_clave(cand)
        for col_norm, col_real in mapa.items():
            if cand_norm in col_norm:
                return col_real
    return None


def construir_nombre_nivel_compuesto(row, df_cols, cols_nivel, separador=" - ", max_len=50):
    partes = []
    for col_cfg in cols_nivel:
        col_real = detectar_columna(df_cols, [col_cfg])
        if not col_real:
            continue
        es_hora = "hora" in normalizar_clave(col_real)
        val_txt = formatear_valor_celda(row.get(col_real, ""), es_hora=es_hora)
        if val_txt:
            partes.append(val_txt)

    combinado = separador.join(partes) if partes else "Sesion"
    return sanitizar_nombre_carpeta(combinado, max_len=max_len)


def extraer_participantes_fila(row, df_cols):
    participantes = []
    vistas = set()

    col_ponente = detectar_columna(df_cols, ["ponente", "ponentes", "autor", "autores", "participante", "participantes", "orador", "speaker"])
    col_moderador = detectar_columna(df_cols, ["moderador", "moderadores", "chair", "coordinador"])
    col_nombre = detectar_columna(df_cols, ["nombre", "first name"])
    col_apellidos = detectar_columna(df_cols, ["apellidos", "apellido", "last name", "surname"])

    def agregar(nombre_raw):
        if not nombre_raw or (isinstance(nombre_raw, float) and pd.isna(nombre_raw)):
            return
        texto = str(nombre_raw).strip()
        if not texto or texto.lower() in ("nan", "none", "null", "-"):
            return
        trozos = re.split(r'[;/\n]+', texto)
        for t in trozos:
            limpio = sanitizar_nombre_carpeta(t, max_len=80)
            clave = normalizar_clave(limpio)
            if limpio and limpio != "Sin_Especificar" and clave not in vistas:
                vistas.add(clave)
                participantes.append(limpio)

    if col_nombre and col_apellidos:
        nom = formatear_valor_celda(row.get(col_nombre, ""))
        ape = formatear_valor_celda(row.get(col_apellidos, ""))
        completo = f"{nom} {ape}".strip()
        if completo:
            agregar(completo)

    if col_ponente:
        agregar(row.get(col_ponente, ""))

    if col_moderador:
        agregar(row.get(col_moderador, ""))

    if not participantes:
        participantes.append("General_Sesion")

    return participantes


def generar_estructura_desde_excel():
    if not os.path.exists(RUTA_EXCEL):
        return False, f"No se encontró el archivo programa.xlsx en {BASE_DIR}"

    cfg = cargar_config()
    orden_carpetas = cfg.get("orden_carpetas", [])
    sep_n3 = cfg.get("separador_nivel_3", " - ")
    max_n3 = int(cfg.get("max_chars_nivel_3", 50))

    try:
        df = pd.read_excel(RUTA_EXCEL, engine="openpyxl")
    except Exception as e:
        return False, f"Error al leer programa.xlsx: {str(e)}"

    df_cols = list(df.columns)
    os.makedirs(RUTA_PPT, exist_ok=True)

    sesiones_creadas = set()
    carpetas_participantes = 0

    for _, row in df.iterrows():
        if row.isna().all():
            continue

        ruta_actual = RUTA_PPT
        for nivel in orden_carpetas:
            if isinstance(nivel, list):
                if len(nivel) > 0:
                    nombre_carpeta = construir_nombre_nivel_compuesto(
                        row, df_cols, nivel, separador=sep_n3, max_len=max_n3
                    )
                    ruta_actual = os.path.join(ruta_actual, nombre_carpeta)
            elif isinstance(nivel, str) and nivel.strip():
                col_real = detectar_columna(df_cols, [nivel.strip()])
                if col_real:
                    es_hora = "hora" in normalizar_clave(col_real)
                    val = formatear_valor_celda(row.get(col_real, ""), es_hora=es_hora)
                    nombre_carpeta = sanitizar_nombre_carpeta(val)
                    ruta_actual = os.path.join(ruta_actual, nombre_carpeta)

        os.makedirs(ruta_actual, exist_ok=True)
        sesiones_creadas.add(ruta_actual)

        # Nivel 4 automático: participantes de cada sesión dentro de la carpeta de la sesión
        participantes = extraer_participantes_fila(row, df_cols)
        for part in participantes:
            ruta_part = os.path.join(ruta_actual, part)
            if not os.path.exists(ruta_part):
                os.makedirs(ruta_part, exist_ok=True)
                carpetas_participantes += 1

    return True, f"Estructura creada en {RUTA_PPT}: {len(sesiones_creadas)} sesiones y {carpetas_participantes} subcarpetas de participantes."


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
    nombre_limpio = sanitizar_nombre_carpeta(nombre_equipo, max_len=15).upper()
    wg_limpio = sanitizar_nombre_carpeta(workgroup, max_len=15).upper()
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
        "ruta_ppt": RUTA_PPT,
        "excel_existe": os.path.exists(RUTA_EXCEL)
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
    <title>CongresosUSB | Panel de Sala y Red</title>
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
        }
        .valex-card-title { font-size: 0.95rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin: 0; }
        .form-select, .form-control {
            background-color: var(--bg-input) !important;
            border: 1px solid rgba(255, 255, 255, 0.12) !important;
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
            <div class="brand-icon"><i class="bi bi-hdd-network-fill"></i></div>
            <span>Congresos<span style="color:#00b9ff;">USB</span> · Agente Local</span>
        </a>
        <div class="d-flex gap-2">
            <span class="badge bg-dark border border-secondary px-3 py-2">PC: {{ info.hostname }} ({{ info.ip }})</span>
            <span class="badge bg-dark border border-info text-info px-3 py-2">Grupo: {{ info.workgroup }}</span>
        </div>
    </header>

    <div class="container-fluid px-4 pt-4" style="max-width: 1200px;">
        <div id="alertaEstado" class="alert d-none mb-4"></div>

        <div class="row g-4">
            <div class="col-lg-6">
                <div class="valex-card h-100">
                    <div class="valex-card-header">
                        <h5 class="valex-card-title">1. Carpetas PPT y Participantes</h5>
                        <i class="bi bi-folder-fill text-warning fs-5"></i>
                    </div>
                    <div class="p-4">
                        <p class="small text-secondary mb-3">
                            Crea en el Escritorio (<code>{{ info.ruta_ppt }}</code>) las carpetas de sesión y, dentro de cada una, las subcarpetas de sus participantes.
                        </p>
                        <div class="d-grid gap-2">
                            <button class="btn btn-valex-primary py-2" onclick="ejecutarAccion('/api/crear_carpetas')">
                                <i class="bi bi-folder-plus me-2"></i> CREAR ESTRUCTURA DE CARPETAS PPT
                            </button>
                            <button class="btn btn-valex-green py-2" onclick="ejecutarAccion('/api/compartir_ppt')">
                                <i class="bi bi-share-fill me-2"></i> COMPARTIR CARPETA PPT EN RED
                            </button>
                            <button class="btn btn-outline-info py-2" onclick="ejecutarAccion('/api/abrir_ppt')">
                                <i class="bi bi-folder2-open me-2"></i> ABRIR CARPETA PPT EN EXPLORADOR
                            </button>
                        </div>
                    </div>
                </div>
            </div>

            <div class="col-lg-6">
                <div class="valex-card h-100">
                    <div class="valex-card-header">
                        <h5 class="valex-card-title">2. Identidad de este Equipo</h5>
                        <i class="bi bi-pc-display text-info fs-5"></i>
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

<script>
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

cargarRed();
</script>
</body>
</html>"""


@app.route('/')
def home():
    return render_template_string(HTML_APP, info=obtener_info_sistema())


@app.route('/api/crear_carpetas', methods=['POST'])
def api_crear_carpetas():
    ok, msg = generar_estructura_desde_excel()
    return jsonify({"ok": ok, "mensaje": msg})


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
    if SISTEMA == "Windows":
        os.startfile(ruta)
    elif SISTEMA == "Darwin":
        subprocess.Popen(["open", ruta])
    return jsonify({"ok": True, "mensaje": f"Abriendo {ruta}"})


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
