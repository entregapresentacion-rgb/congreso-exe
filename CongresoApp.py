import os
import sys
import json
import time
import ctypes
import socket
import platform
import threading
import subprocess
import webbrowser
import requests
import pandas as pd
from datetime import datetime
from flask import Flask, render_template_string, request, send_from_directory, redirect, flash, jsonify
from werkzeug.utils import secure_filename

SISTEMA = platform.system()

def es_admin():
    if SISTEMA == 'Windows':
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except Exception:
            return False
    else:
        return os.geteuid() == 0

def obtener_ruta_base():
    if getattr(sys, 'frozen', False):
        ruta = os.path.dirname(sys.executable)
        if SISTEMA == 'Darwin' and ruta.endswith('Contents/MacOS'):
            return os.path.abspath(os.path.join(ruta, '..', '..', '..'))
        return ruta
    return os.path.abspath(os.path.dirname(__file__))

BASE_DIR = obtener_ruta_base()
CONFIG_FILE = os.path.join(BASE_DIR, 'config_congreso.json')
EXCEL_FILE = os.path.join(BASE_DIR, 'programa.xlsx')
CARPETA_PPT = r"C:\PPT" if SISTEMA == 'Windows' else "/Users/Shared/PPT"

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"workgroup": "CONGRESO", "orden_carpetas": ["Sala", "Sesion"], "equipos": ["ADM1"]}

CONFIG = cargar_config()
EQUIPO_ACTUAL = None
ES_SERVIDOR_CENTRAL = False

def configurar_entorno_sistema(nombre_equipo):
    workgroup = CONFIG.get("workgroup", "CONGRESO")
    os.makedirs(CARPETA_PPT, exist_ok=True)
    pc_actual = socket.gethostname().split('.')[0].upper()

    if SISTEMA == 'Windows':
        escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
        ruta_lnk = os.path.join(escritorio, "PPT.lnk")
        ps_shortcut = f"""
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut('{ruta_lnk}')
        $Shortcut.TargetPath = '{CARPETA_PPT}'
        $Shortcut.Save()
        """
        subprocess.run(["powershell", "-NoProfile", "-Command", ps_shortcut], creationflags=0x08000000)
        subprocess.run('net share PPT="C:\\PPT" /GRANT:Everyone,FULL', shell=True, creationflags=0x08000000)
        subprocess.run('icacls "C:\\PPT" /grant Everyone:(OI)(CI)F /T', shell=True, creationflags=0x08000000)

        ps_net = f"""
        $cs = Get-WmiObject Win32_ComputerSystem
        if ($cs.Workgroup -ne '{workgroup}') {{
            $cs.JoinDomainOrWorkgroup('{workgroup}') | Out-Null
        }}
        if ($env:COMPUTERNAME -ne '{nombre_equipo.upper()}') {{
            Rename-Computer -NewName '{nombre_equipo.upper()}' -Force -ErrorAction SilentlyContinue
        }}
        """
        subprocess.run(["powershell", "-NoProfile", "-Command", ps_net], creationflags=0x08000000)
        return pc_actual != nombre_equipo.upper()

    elif SISTEMA == 'Darwin':
        os.chmod(CARPETA_PPT, 0o777)
        usuario_real = os.environ.get('SUDO_USER') or os.environ.get('USER')
        escritorio_mac = os.path.expanduser(f"~{usuario_real}/Desktop")
        enlace_mac = os.path.join(escritorio_mac, "PPT")
        if not os.path.exists(enlace_mac):
            try:
                os.symlink(CARPETA_PPT, enlace_mac)
            except OSError:
                pass

        subprocess.run(["sharing", "-a", CARPETA_PPT, "-S", "PPT"], stderr=subprocess.DEVNULL)

        nombre_limpio = nombre_equipo.replace('_', '-').replace(' ', '-')
        if pc_actual != nombre_equipo.upper():
            subprocess.run(["scutil", "--set", "ComputerName", nombre_equipo.upper()])
            subprocess.run(["scutil", "--set", "LocalHostName", nombre_limpio])
            subprocess.run(["scutil", "--set", "HostName", nombre_limpio])
            subprocess.run([
                "defaults", "write", "/Library/Preferences/SystemConfiguration/com.apple.smb.server",
                "Workgroup", workgroup
            ])
            return True
        return False

def limpiar_nombre_carpeta(texto):
    invalids = '<>:"/\\|?*'
    res = "".join(c for c in str(texto).strip() if c not in invalids)
    return res if res else "General"

def construir_ruta_destino(fila_ponencia):
    orden = CONFIG.get("orden_carpetas", ["Sala", "Sesion"])
    subrutas = [limpiar_nombre_carpeta(fila_ponencia.get(col, "General")) for col in orden if col in fila_ponencia]
    ruta_final = os.path.join(CARPETA_PPT, *subrutas)
    os.makedirs(ruta_final, exist_ok=True)
    if SISTEMA == 'Darwin':
        try:
            os.chmod(ruta_final, 0o777)
        except OSError:
            pass
    return ruta_final

def abrir_pantalla_completa(ruta_archivo):
    if not os.path.exists(ruta_archivo):
        return False
    ext = os.path.splitext(ruta_archivo)[1].lower()
    if SISTEMA == 'Windows':
        try:
            if ext in ['.ppt', '.pptx', '.pps', '.ppsx']:
                ps_ppt = f"""
                try {{
                    $ppt = New-Object -ComObject PowerPoint.Application
                    $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
                    $pres = $ppt.Presentations.Open('{ruta_archivo}', $true, $false, $true)
                    $pres.SlideShowSettings.Run() | Out-Null
                }} catch {{
                    Start-Process '{ruta_archivo}'
                }}
                """
                subprocess.Popen(["powershell", "-NoProfile", "-Command", ps_ppt], creationflags=0x08000000)
            else:
                os.startfile(ruta_archivo)
            return True
        except Exception:
            os.startfile(ruta_archivo)
            return True
    elif SISTEMA == 'Darwin':
        try:
            if ext in ['.ppt', '.pptx', '.pps', '.ppsx']:
                script_mac = f'''
                tell application "Microsoft PowerPoint"
                    activate
                    open POSIX file "{ruta_archivo}"
                    delay 1
                    run slide show slide show settings of active presentation
                end tell
                '''
                res = subprocess.run(["osascript", "-e", script_mac], stderr=subprocess.DEVNULL)
                if res.returncode != 0:
                    subprocess.Popen(["open", ruta_archivo])
            elif ext == '.key':
                script_key = f'''
                tell application "Keynote"
                    activate
                    open POSIX file "{ruta_archivo}"
                    delay 1
                    start the front document
                end tell
                '''
                subprocess.run(["osascript", "-e", script_key], stderr=subprocess.DEVNULL)
            elif ext == '.pdf':
                script_pdf = f'''
                tell application "Preview"
                    activate
                    open POSIX file "{ruta_archivo}"
                    delay 0.8
                    tell application "System Events" to keystroke "f" using {{command down, control down}}
                end tell
                '''
                subprocess.run(["osascript", "-e", script_pdf], stderr=subprocess.DEVNULL)
            else:
                subprocess.Popen(["open", ruta_archivo])
            return True
        except Exception:
            subprocess.Popen(["open", ruta_archivo])
            return True

app = Flask(__name__)
app.secret_key = 'congreso_unificado_key'
lock_archivos = threading.Lock()
ALLOWED_EXTENSIONS = {'pptx', 'ppt', 'pdf', 'key', 'mp4', 'zip', 'mov', 'avi'}

SELECTOR_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Inicializar Puesto - Congreso</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-dark d-flex align-items-center justify-content-center" style="min-height: 100vh;">
<div class="card shadow-lg border-0 rounded-4 p-4" style="max-width: 460px; width: 100%;">
    <div class="text-center mb-3">
        <h4 class="fw-bold mb-1">🖥 Configuración de Puesto</h4>
        <span class="badge bg-success">Grupo de Trabajo: {{ workgroup }}</span>
    </div>
    <p class="text-muted text-center small">Selecciona el nombre asignado a este ordenador. El sistema configurará la red y la carpeta <strong>PPT</strong> automáticamente.</p>
    <form method="POST" action="/seleccionar_equipo">
        <div class="mb-3">
            <select name="equipo" class="form-select form-select-lg fw-bold" required>
                {% for eq in equipos %}
                    <option value="{{ eq }}">{{ eq }}</option>
                {% endfor %}
            </select>
        </div>
        <button type="submit" class="btn btn-primary btn-lg w-100 fw-bold">INICIAR SISTEMA EN ESTE EQUIPO</button>
    </form>
</div>
</body>
</html>"""

DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestión de Congreso - {{ equipo }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        body { background-color: #f0f2f5; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
        .session-card { border: none; border-radius: 12px; margin-bottom: 12px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }
        .session-header { background: #ffffff; padding: 16px 22px; cursor: pointer; transition: background 0.2s; border-left: 5px solid #0d6efd; }
        .session-header:hover { background: #f8faff; }
        .btn-play-active { background-color: #198754; color: white; border: none; width: 42px; height: 42px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: 1.2rem; transition: transform 0.15s, background 0.2s; box-shadow: 0 3px 8px rgba(25,135,84,0.35); }
        .btn-play-active:hover { background-color: #146c43; transform: scale(1.08); color: white; }
        .btn-play-disabled { background-color: #dee2e6; color: #adb5bd; border: none; width: 42px; height: 42px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: 1.2rem; cursor: not-allowed; pointer-events: none; }
        .upload-wrapper { position: relative; overflow: hidden; display: inline-block; }
        .upload-wrapper input[type=file] { font-size: 100px; position: absolute; left: 0; top: 0; opacity: 0; cursor: pointer; }
    </style>
</head>
<body>

<nav class="navbar navbar-dark bg-dark shadow-sm mb-4">
    <div class="container-fluid px-4">
        <span class="navbar-brand fw-bold d-flex align-items-center gap-2">
            <i class="bi bi-display text-success fs-4"></i>
            <span>CONGRESO LIVE</span>
            <span class="badge {% if es_admin %}bg-primary{% else %}bg-success{% endif %} ms-2">{{ equipo }}</span>
            <span class="badge bg-secondary fw-normal">Grupo: {{ workgroup }}</span>
        </span>
        <div class="d-flex align-items-center gap-3">
            <span class="text-light small"><i class="bi bi-folder-check text-warning"></i> Carpeta Local: <strong>{{ ruta_ppt }}</strong></span>
            <button onclick="location.reload()" class="btn btn-outline-light btn-sm rounded-pill px-3"><i class="bi bi-arrow-clockwise"></i> Actualizar</button>
        </div>
    </div>
</nav>

<div class="container-fluid px-4 pb-5">

    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for category, message in messages %}
          <div class="alert alert-{{ category }} alert-dismissible fade show shadow-sm">{{ message }}<button type="button" class="btn-close" data-bs-dismiss="alert"></button></div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <div class="row g-3 mb-4">
        <div class="col-md-3">
            <div class="card border-0 shadow-sm rounded-3 p-3">
                <small class="text-muted fw-bold text-uppercase">Sesiones Mostradas</small>
                <h2 class="fw-bold mb-0">{{ sesiones|length }}</h2>
            </div>
        </div>
        <div class="col-md-3">
            <div class="card border-0 shadow-sm rounded-3 p-3">
                <small class="text-muted fw-bold text-uppercase">Total Ponencias</small>
                <h2 class="fw-bold mb-0">{{ total_ponencias }}</h2>
            </div>
        </div>
        <div class="col-md-3">
            <div class="card border-0 shadow-sm rounded-3 p-3">
                <small class="text-muted fw-bold text-uppercase">Listas en Local (PPT)</small>
                <h2 class="fw-bold text-success mb-0">{{ total_listas }}</h2>
            </div>
        </div>
        <div class="col-md-3">
            <div class="card border-0 shadow-sm rounded-3 p-3">
                <small class="text-muted fw-bold text-uppercase">Pendientes</small>
                <h2 class="fw-bold text-secondary mb-0">{{ total_ponencias - total_listas }}</h2>
            </div>
        </div>
    </div>

    <div class="card border-0 shadow-sm rounded-3 p-3 mb-4">
        <form method="GET" action="/" class="row g-2 align-items-center">
            {% if es_admin %}
            <div class="col-md-3">
                <select name="sala" class="form-select form-select-sm" onchange="this.form.submit()">
                    <option value="">-- Todas las Salas --</option>
                    {% for s in salas_disponibles %}
                        <option value="{{ s }}" {% if s == sala_filtro %}selected{% endif %}>{{ s }}</option>
                    {% endfor %}
                </select>
            </div>
            {% endif %}
            <div class="col-md-5">
                <div class="input-group input-group-sm">
                    <span class="input-group-text bg-white"><i class="bi bi-search"></i></span>
                    <input type="text" name="q" value="{{ busqueda }}" class="form-control" placeholder="Buscar sesión, ponente o título...">
                    <button type="submit" class="btn btn-primary">Filtrar</button>
                </div>
            </div>
            {% if sala_filtro or busqueda %}
            <div class="col-auto"><a href="/" class="btn btn-sm btn-outline-danger">Limpiar</a></div>
            {% endif %}
        </form>
    </div>

    <h6 class="text-uppercase text-muted fw-bold small mb-3"><i class="bi bi-layers-fill me-1"></i> Sesiones del Programa (Haz clic en una sesión para desplegar sus ponencias)</h6>
    
    <div class="accordion" id="accordionSesiones">
        {% for sesion in sesiones %}
        <div class="session-card">
            <div class="session-header d-flex justify-content-between align-items-center" data-bs-toggle="collapse" data-bs-target="#collapse_{{ loop.index }}">
                <div>
                    <div class="d-flex align-items-center gap-2 mb-1">
                        <span class="badge bg-primary">{{ sesion.sala }}</span>
                        <span class="text-muted small fw-semibold"><i class="bi bi-calendar-event"></i> {{ sesion.fecha }} | {{ sesion.horario }}</span>
                    </div>
                    <h5 class="mb-0 fw-bold text-dark">{{ sesion.nombre }}</h5>
                </div>
                <div class="d-flex align-items-center gap-3">
                    <span class="badge {% if sesion.listas == sesion.items|length %}bg-success{% else %}bg-dark{% endif %} rounded-pill px-3 py-2">
                        {{ sesion.listas }} / {{ sesion.items|length }} listas
                    </span>
                    <i class="bi bi-chevron-down fs-5 text-muted"></i>
                </div>
            </div>

            <div id="collapse_{{ loop.index }}" class="collapse {% if loop.first or busqueda %}show{% endif %}" data-bs-parent="#accordionSesiones">
                <div class="bg-white border-top p-0">
                    <table class="table table-hover align-middle mb-0">
                        <thead class="table-light small text-muted">
                            <tr>
                                <th class="ps-4" style="width:80px;">ID</th>
                                <th style="width:110px;">HORA</th>
                                <th style="width:240px;">PONENTE</th>
                                <th>TÍTULO DE LA PRESENTACIÓN</th>
                                <th class="text-center" style="width:140px;">PROYECTAR</th>
                                {% if es_admin %}
                                <th class="text-end pe-4" style="width:180px;">CARGAR ARCHIVO</th>
                                {% endif %}
                            </tr>
                        </thead>
                        <tbody>
                            {% for p in sesion.items %}
                            <tr>
                                <td class="ps-4"><span class="badge bg-light text-dark border">#{{ p.ID }}</span></td>
                                <td class="fw-bold text-primary">{{ p.Hora }}</td>
                                <td class="fw-semibold">{{ p.Ponente }}</td>
                                <td>{{ p.Titulo }}</td>
                                <td class="text-center">
                                    {% if p.en_local %}
                                        <button onclick="lanzarPresentacion('{{ p.ID }}')" class="btn-play-active" title="Abrir a Pantalla Completa">
                                            <i class="bi bi-play-fill"></i>
                                        </button>
                                    {% else %}
                                        <span class="btn-play-disabled" title="Presentación aún no cargada">
                                            <i class="bi bi-play-fill"></i>
                                        </span>
                                    {% endif %}
                                </td>
                                {% if es_admin %}
                                <td class="text-end pe-4">
                                    <form action="/subir/{{ p.ID }}" method="POST" enctype="multipart/form-data" class="m-0">
                                        <div class="upload-wrapper">
                                            <button type="button" class="btn btn-sm {% if p.en_local %}btn-outline-success{% else %}btn-primary{% endif %} rounded-pill px-3">
                                                <i class="bi bi-cloud-upload me-1"></i>{% if p.en_local %}Reemplazar{% else %}Subir{% endif %}
                                            </button>
                                            <input type="file" name="archivo" onchange="this.form.submit()" required>
                                        </div>
                                    </form>
                                </td>
                                {% endif %}
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        {% else %}
        <div class="alert alert-light text-center py-5 shadow-sm">No hay sesiones disponibles para mostrar.</div>
        {% endfor %}
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
<script>
function lanzarPresentacion(id) {
    fetch('/play/' + id, { method: 'POST' })
        .then(r => r.json())
        .then(data => {
            if(!data.ok) alert('No se pudo abrir la presentación en local.');
        });
}
setInterval(() => {
    if (document.activeElement.type !== 'file' && document.activeElement.type !== 'text') {
        location.reload();
    }
}, 12000);
</script>
</body>
</html>"""

def buscar_archivo_en_ppt(ponencia_id):
    prefijo = f"ID_{ponencia_id}_"
    encontrados = []
    for root, _, files in os.walk(CARPETA_PPT):
        for f in files:
            if f.startswith(prefijo):
                encontrados.append(os.path.join(root, f))
    if encontrados:
        encontrados.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        return encontrados[0]
    return None

def obtener_servidor_url():
    for host in ["http://ADM1:5000", "http://ADM1.local:5000"]:
        try:
            r = requests.get(f"{host}/api/ping", timeout=1.5)
            if r.status_code == 200:
                return host
        except Exception:
            pass
    return "http://ADM1:5000"

def obtener_datos_completos():
    registros = []
    if ES_SERVIDOR_CENTRAL:
        if os.path.exists(EXCEL_FILE):
            df = pd.read_excel(EXCEL_FILE, engine='openpyxl').fillna('')
            registros = df.to_dict(orient='records')
    else:
        url_srv = obtener_servidor_url()
        try:
            r = requests.get(f"{url_srv}/api/raw_excel", timeout=3)
            if r.status_code == 200:
                registros = r.json()
                pd.DataFrame(registros).to_excel(EXCEL_FILE, index=False)
        except Exception:
            if os.path.exists(EXCEL_FILE):
                df = pd.read_excel(EXCEL_FILE, engine='openpyxl').fillna('')
                registros = df.to_dict(orient='records')

    for reg in registros:
        id_val = str(reg.get('ID', '')).strip().replace('.0', '')
        reg['ID'] = id_val
        ruta_loc = buscar_archivo_en_ppt(id_val)
        reg['ruta_local'] = ruta_loc
        reg['en_local'] = bool(ruta_loc)
        reg['nombre_archivo'] = os.path.basename(ruta_loc) if ruta_loc else None
    return registros

def sincronizador_segundo_plano():
    while True:
        if EQUIPO_ACTUAL and not ES_SERVIDOR_CENTRAL:
            url_srv = obtener_servidor_url()
            try:
                r = requests.get(f"{url_srv}/api/raw_excel", timeout=5)
                if r.status_code == 200:
                    datos = r.json()
                    es_admin = EQUIPO_ACTUAL.upper().startswith("ADM")
                    for p in datos:
                        if not es_admin and str(p.get('Sala', '')).strip().lower() != EQUIPO_ACTUAL.replace('_', ' ').lower():
                            continue
                        arch_remoto = p.get('nombre_archivo')
                        id_p = str(p.get('ID', '')).strip().replace('.0', '')
                        if arch_remoto:
                            carpeta_dest = construir_ruta_destino(p)
                            ruta_dest = os.path.join(carpeta_dest, arch_remoto)
                            if not os.path.exists(ruta_dest):
                                with requests.get(f"{url_srv}/descargar_id/{id_p}", stream=True, timeout=30) as rf:
                                    if rf.status_code == 200:
                                        with open(ruta_dest, "wb") as f:
                                            for chunk in rf.iter_content(8192):
                                                f.write(chunk)
                                        for root, _, files in os.walk(CARPETA_PPT):
                                            for f_old in files:
                                                if f_old.startswith(f"ID_{id_p}_") and f_old != arch_remoto:
                                                    try: os.remove(os.path.join(root, f_old))
                                                    except OSError: pass
            except Exception:
                pass
        time.sleep(6)

@app.route('/api/ping')
def api_ping():
    return jsonify({"status": "ok", "equipo": EQUIPO_ACTUAL})

@app.route('/seleccionar_equipo', methods=['POST'])
def seleccionar_equipo():
    global EQUIPO_ACTUAL, ES_SERVIDOR_CENTRAL
    eq = request.form.get('equipo')
    if eq:
        EQUIPO_ACTUAL = eq
        ES_SERVIDOR_CENTRAL = (EQUIPO_ACTUAL.upper() == "ADM1")
        reinicio = configurar_entorno_sistema(EQUIPO_ACTUAL)
        if reinicio:
            flash(f"Equipo renombrado a '{EQUIPO_ACTUAL}' en grupo '{CONFIG.get('workgroup')}'. Reinicia el equipo si es la primera vez para aplicar el nombre en red.", "warning")
    return redirect('/')

@app.route('/')
def index():
    if not EQUIPO_ACTUAL:
        return render_template_string(SELECTOR_HTML, workgroup=CONFIG.get("workgroup", "CONGRESO"), equipos=CONFIG.get("equipos", ["ADM1"]))

    es_admin = EQUIPO_ACTUAL.upper().startswith("ADM")
    sala_filtro = request.args.get('sala', '') if es_admin else EQUIPO_ACTUAL.replace('_', ' ')
    busqueda = request.args.get('q', '').lower().strip()

    datos = obtener_datos_completos()
    salas_disponibles = sorted(list(set(str(d.get('Sala', '')).strip() for d in datos if str(d.get('Sala', '')).strip())))

    if sala_filtro:
        datos = [d for d in datos if str(d.get('Sala', '')).strip().lower() == sala_filtro.lower()]
    if busqueda:
        datos = [d for d in datos if busqueda in str(d.get('Ponente','')).lower() or busqueda in str(d.get('Titulo','')).lower() or busqueda in str(d.get('Sesion','')).lower()]

    sesiones_dict = {}
    for d in datos:
        nombre_sesion = str(d.get('Sesion', '')).strip() or f"Sesión General - {d.get('Sala', '')}"
        clave = f"{d.get('Fecha', '')}_{d.get('Sala', '')}_{nombre_sesion}"
        if clave not in sesiones_dict:
            sesiones_dict[clave] = {
                'nombre': nombre_sesion,
                'sala': d.get('Sala', ''),
                'fecha': d.get('Fecha', ''),
                'horario': d.get('Hora', ''),
                'items': [],
                'listas': 0
            }
        sesiones_dict[clave]['items'].append(d)
        if d['en_local']:
            sesiones_dict[clave]['listas'] += 1

    sesiones_lista = list(sesiones_dict.values())
    total_ponencias = len(datos)
    total_listas = sum(1 for d in datos if d['en_local'])

    return render_template_string(
        DASHBOARD_HTML,
        equipo=EQUIPO_ACTUAL,
        workgroup=CONFIG.get("workgroup", "CONGRESO"),
        ruta_ppt=CARPETA_PPT,
        es_admin=es_admin,
        sesiones=sesiones_lista,
        salas_disponibles=salas_disponibles,
        sala_filtro=sala_filtro,
        busqueda=busqueda,
        total_ponencias=total_ponencias,
        total_listas=total_listas
    )

@app.route('/play/<ponencia_id>', methods=['POST'])
def play_presentacion(ponencia_id):
    ruta = buscar_archivo_en_ppt(ponencia_id)
    if ruta and abrir_pantalla_completa(ruta):
        return jsonify({"ok": True})
    return jsonify({"ok": False})

@app.route('/subir/<ponencia_id>', methods=['POST'])
def subir_presentacion(ponencia_id):
    file = request.files.get('archivo')
    if not file or file.filename == '':
        return redirect(request.referrer)
    ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
    if ext not in ALLOWED_EXTENSIONS:
        flash('Formato no soportado.', 'warning')
        return redirect(request.referrer)

    if ES_SERVIDOR_CENTRAL:
        with lock_archivos:
            datos = obtener_datos_completos()
            ponencia = next((p for p in datos if str(p['ID']) == str(ponencia_id)), None)
            if ponencia:
                for root, _, files in os.walk(CARPETA_PPT):
                    for f_old in files:
                        if f_old.startswith(f"ID_{ponencia_id}_"):
                            try: os.remove(os.path.join(root, f_old))
                            except OSError: pass
                carpeta_dest = construir_ruta_destino(ponencia)
                ts = datetime.now().strftime("%H%M%S")
                hora_l = str(ponencia.get('Hora', '')).replace(':', '-')
                nombre = secure_filename(f"ID_{ponencia_id}_{hora_l}_{ponencia.get('Ponente','')}_v{ts}") + f".{ext}"
                file.save(os.path.join(carpeta_dest, nombre))
                flash(f'✔ Presentación guardada en {CARPETA_PPT} para {ponencia.get("Ponente")}', 'success')
    else:
        url_srv = obtener_servidor_url()
        try:
            files = {'archivo': (file.filename, file.stream, file.mimetype)}
            r = requests.post(f"{url_srv}/subir/{ponencia_id}", files=files, timeout=30)
            if r.status_code == 200:
                flash('✔ Enviado al servidor central ADM1 y replicado.', 'success')
        except Exception:
            flash('❌ Error conectando con ADM1.', 'danger')

    return redirect(request.referrer)

@app.route('/api/raw_excel')
def api_raw_excel():
    return jsonify(obtener_datos_completos())

@app.route('/descargar_id/<ponencia_id>')
def descargar_id(ponencia_id):
    ruta = buscar_archivo_en_ppt(ponencia_id)
    if ruta:
        return send_from_directory(os.path.dirname(ruta), os.path.basename(ruta), as_attachment=True)
    return "No encontrado", 404

if __name__ == '__main__':
    if not es_admin():
        if SISTEMA == 'Windows':
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()
        elif SISTEMA == 'Darwin':
            cmd_mac = f'do shell script "\\"{sys.executable}\\"" with administrator privileges'
            subprocess.Popen(["osascript", "-e", cmd_mac])
            sys.exit()

    threading.Thread(target=sincronizador_segundo_plano, daemon=True).start()
    threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    app.run(host='0.0.0.0', port=5000, threaded=True)
