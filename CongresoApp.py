import os
import re
import sys
import json
import platform
import threading
import subprocess
import webbrowser
import pandas as pd
from flask import Flask, render_template_string, request, jsonify

SISTEMA = platform.system()

def obtener_ruta_base():
    if getattr(sys, 'frozen', False):
        ruta = os.path.dirname(sys.executable)
        if SISTEMA == 'Darwin' and ruta.endswith('Contents/MacOS'):
            ruta = os.path.abspath(os.path.join(ruta, '..', '..', '..'))
        return ruta
    return os.path.abspath(os.path.dirname(__file__))

def limpiar_nombre_carpeta(texto):
    if pd.isna(texto) or texto is None:
        return "Sin_Especificar"
    s = str(texto).strip()
    s = re.sub(r'[\\/*?:"<>|]', '-', s)
    return s[:85].strip() or "Sin_Especificar"

def cargar_datos_congreso():
    base = obtener_ruta_base()
    # Busca los archivos en la misma carpeta o dentro de PAQUETE_USB_CONGRESO
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
    return config, df, ruta_activa

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
        .valex-search {
            background-color: var(--bg-input);
            border: 1px solid var(--border-color);
            border-radius: 20px;
            padding: 6px 16px;
            color: #fff;
            font-size: 0.88rem;
            width: 320px;
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
            border: 1px solid rgba(255, 255, 255, 0.12) !important;
            color: #fff !important;
            font-size: 0.88rem;
        }
        .btn-valex-blue { background: linear-gradient(to right, #0162e8, #00b9ff); border: none; color: #fff; font-weight: 600; }
        .btn-valex-green { background: linear-gradient(to right, #029666, #48d6a8); border: none; color: #fff; font-weight: 600; }
        .btn-valex-blue:hover, .btn-valex-green:hover { color: #fff; opacity: 0.92; }
        .table-valex { color: #e2e8f0; margin-bottom: 0; font-size: 0.86rem; }
        .table-valex thead th {
            background-color: #171f32;
            color: var(--text-muted);
            text-transform: uppercase;
            font-size: 0.73rem;
            letter-spacing: 0.6px;
            border-bottom: 1px solid var(--border-color);
            padding: 12px 16px;
        }
        .table-valex tbody td {
            border-bottom: 1px solid var(--border-color);
            padding: 12px 16px;
            vertical-align: middle;
        }
        .table-valex tbody tr:hover { background-color: rgba(255, 255, 255, 0.03); }
        .badge-ppt-ok { background: rgba(2, 150, 102, 0.2); color: #48d6a8; border: 1px solid rgba(72, 214, 168, 0.35); }
        .badge-ppt-wait { background: rgba(247, 106, 45, 0.2); color: #efa65f; border: 1px solid rgba(239, 166, 95, 0.35); }
    </style>
</head>
<body>

    <header class="valex-topbar d-flex flex-wrap justify-content-between align-items-center gap-3">
        <div class="d-flex align-items-center gap-4">
            <a href="/" class="brand-logo">
                <div class="brand-icon"><i class="bi bi-easel2-fill"></i></div>
                <span>Congreso<span style="color:#00b9ff;">App</span></span>
            </a>
            <input type="text" id="buscadorGlobal" class="valex-search" placeholder="Buscar ponente, sesión o sala..." onkeyup="filtrarTabla()">
        </div>
        <div class="d-flex align-items-center gap-3">
            <span class="badge bg-dark border border-secondary px-3 py-2">
                <i class="bi bi-hdd-network text-info me-1"></i> Red: {{ config.workgroup if config else 'N/A' }}
            </span>
            <span class="badge bg-dark border border-secondary px-3 py-2">
                <i class="bi bi-pc-display text-success me-1"></i> {{ equipo_actual }}
            </span>
        </div>
    </header>

    <nav class="valex-subnav">
        <span class="subnav-item active" onclick="filtrarPorSala('')"><i class="bi bi-grid-1x2"></i> Todas las Salas</span>
        {% if config %}
            {% for eq in config.equipos %}
                {% if not eq.startswith('ADM') %}
                    <span class="subnav-item" onclick="filtrarPorSala('{{ eq }}')"><i class="bi bi-door-open"></i> {{ eq }}</span>
                {% endif %}
            {% endfor %}
        {% endif %}
    </nav>

    <div class="container-fluid px-4 pt-4" style="max-width: 1360px;">

        {% if not config %}
            <div class="valex-card p-5 text-center my-5">
                <i class="bi bi-exclamation-diamond text-warning display-4 mb-3"></i>
                <h4 class="fw-bold">No se encontraron los archivos de configuración</h4>
                <p style="color: var(--text-muted); max-width: 560px; margin: 10px auto 20px;">
                    Asegúrate de colocar <code>programa.xlsx</code> y <code>config_congreso.json</code> en la misma carpeta que <strong>CongresoApp</strong> (dentro de <code>PAQUETE_USB_CONGRESO</code>).
                </p>
                <a href="/" class="btn btn-valex-blue px-4">Reintentar lectura</a>
            </div>
        {% else %}

        <!-- Cabecera + Botones de Acción Rápida -->
        <div class="d-flex flex-wrap justify-content-between align-items-center mb-4 gap-3">
            <div>
                <h4 class="fw-bold mb-1">Panel de Gestión de Ponencias y Salas</h4>
                <p class="mb-0" style="color: var(--text-muted); font-size: 0.88rem;">
                    Ruta activa: <code>{{ ruta_activa }}</code>
                </p>
            </div>
            <div class="d-flex flex-wrap gap-2">
                <button class="btn btn-valex-green px-3 py-2" onclick="crearArbolCarpetas()">
                    <i class="bi bi-folder-plus me-1"></i> Crear Estructura de Carpetas PPT
                </button>
                <button class="btn btn-valex-blue px-3 py-2" onclick="abrirCarpetaBase()">
                    <i class="bi bi-folder2-open me-1"></i> Abrir Carpeta PPT
                </button>
            </div>
        </div>

        <div id="alertaAccion" class="alert d-none mb-4"></div>

        <!-- 4 Tarjetas KPI -->
        <div class="row g-3 mb-4">
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-blue">
                    <div class="kpi-title">Total Ponencias</div>
                    <div class="kpi-value">{{ total_ponencias }}</div>
                    <div class="kpi-sub">Cargadas desde programa.xlsx</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,30 L30,22 L60,28 L95,12 L130,24 L165,10 L200,18 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-pink">
                    <div class="kpi-title">Equipos Configurados</div>
                    <div class="kpi-value">{{ config.equipos|length }} Equipos</div>
                    <div class="kpi-sub">{{ config.equipos|join(', ') }}</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,25 L40,32 L80,15 L120,25 L160,12 L200,22 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-green">
                    <div class="kpi-title">Presentaciones Listas</div>
                    <div class="kpi-value">{{ con_archivo }} / {{ total_ponencias }}</div>
                    <div class="kpi-sub">Archivos detectados en PPT</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,32 L35,18 L75,26 L115,10 L155,20 L200,8 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-orange">
                    <div class="kpi-title">Jerarquía Activa</div>
                    <div class="kpi-value">{{ config.orden_carpetas|length }} Niveles</div>
                    <div class="kpi-sub">{{ config.orden_carpetas|join(' / ') }}</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,20 L45,28 L85,14 L130,26 L170,15 L200,24 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
        </div>

        <!-- Configuración rápida de Puesto + Tabla de Ponencias -->
        <div class="valex-card mb-4">
            <div class="valex-card-header">
                <div>
                    <h5 class="valex-card-title">Identificación de este Ordenador en el Congreso</h5>
                    <div class="valex-card-subtitle">Selecciona qué puesto es este equipo según <code>config_congreso.json</code></div>
                </div>
                <div class="d-flex align-items-center gap-2">
                    <select id="selectEquipo" class="form-select form-select-sm" style="width: 220px;">
                        {% for eq in config.equipos %}
                            <option value="{{ eq }}" {% if eq == equipo_actual %}selected{% endif %}>{{ eq }}</option>
                        {% endfor %}
                    </select>
                    <button class="btn btn-sm btn-valex-blue px-3" onclick="aplicarEquipo()">Aplicar Puesto</button>
                </div>
            </div>
        </div>

        <div class="valex-card">
            <div class="valex-card-header">
                <div>
                    <h5 class="valex-card-title">Programa del Congreso y Carpetas de Ponencia</h5>
                    <div class="valex-card-subtitle">Pulsa en "Abrir Carpeta" para guardar o lanzar el PowerPoint de cada ponente</div>
                </div>
                <div class="d-flex gap-2">
                    <select id="filtroSala" class="form-select form-select-sm" style="width: 180px;" onchange="filtrarTabla()">
                        <option value="">Todas las Salas</option>
                        {% for eq in config.equipos %}
                            {% if not eq.startswith('ADM') %}
                                <option value="{{ eq }}">{{ eq }}</option>
                            {% endif %}
                        {% endfor %}
                    </select>
                </div>
            </div>
            <div class="table-responsive">
                <table class="table table-valex" id="tablaPonencias">
                    <thead>
                        <tr>
                            {% for col in columnas %}
                                <th>{{ col }}</th>
                            {% endfor %}
                            <th>Estado PPT</th>
                            <th class="text-end">Carpeta Ponencia</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for fila in filas %}
                        <tr data-sala="{{ fila.sala }}">
                            {% for col in columnas %}
                                <td>{{ fila.datos[col] }}</td>
                            {% endfor %}
                            <td>
                                {% if fila.tiene_ppt %}
                                    <span class="badge badge-ppt-ok"><i class="bi bi-check-circle me-1"></i>{{ fila.num_archivos }} archivo(s)</span>
                                {% else %}
                                    <span class="badge badge-ppt-wait"><i class="bi bi-clock me-1"></i>Pendiente</span>
                                {% endif %}
                            </td>
                            <td class="text-end">
                                <button class="btn btn-sm btn-valex-blue" onclick='abrirRuta({{ fila.ruta_relativa | tojson }})'>
                                    <i class="bi bi-folder2-open me-1"></i> Abrir Carpeta
                                </button>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
        {% endif %}
    </div>

<script>
function mostrarAlerta(texto, ok = true) {
    const el = document.getElementById('alertaAccion');
    el.className = 'alert mb-4';
    el.style.background = ok ? 'rgba(2,150,102,0.2)' : 'rgba(249,58,90,0.2)';
    el.style.border = ok ? '1px solid #48d6a8' : '1px solid #f7778c';
    el.style.color = ok ? '#6ee7b7' : '#fda4af';
    el.innerHTML = `<i class="bi ${ok ? 'bi-check-circle-fill' : 'bi-exclamation-triangle-fill'} me-2"></i>${texto}`;
}

function crearArbolCarpetas() {
    fetch('/api/crear-carpetas', { method: 'POST' })
        .then(r => r.json())
        .then(data => {
            mostrarAlerta(data.mensaje, data.ok);
            if (data.ok) setTimeout(() => location.reload(), 1200);
        });
}

function abrirCarpetaBase() {
    fetch('/api/abrir-ruta', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ruta: 'PPT' })
    });
}

function abrirRuta(rutaRelativa) {
    fetch('/api/abrir-ruta', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ruta: rutaRelativa })
    });
}

function aplicarEquipo() {
    const equipo = document.getElementById('selectEquipo').value;
    fetch('/api/configurar-equipo', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ equipo })
    })
    .then(r => r.json())
    .then(data => mostrarAlerta(data.mensaje, data.ok));
}

function filtrarPorSala(sala) {
    document.getElementById('filtroSala').value = sala;
    filtrarTabla();
}

function filtrarTabla() {
    const texto = document.getElementById('buscadorGlobal').value.toLowerCase();
    const sala = document.getElementById('filtroSala').value.toLowerCase();
    const filas = document.querySelectorAll('#tablaPonencias tbody tr');
    filas.forEach(tr => {
        const contenido = tr.textContent.toLowerCase();
        const salaFila = (tr.getAttribute('data-sala') || '').toLowerCase();
        const coincideTexto = contenido.includes(texto);
        const coincideSala = !sala || salaFila === sala;
        tr.style.display = (coincideTexto && coincideSala) ? '' : 'none';
    });
}
</script>
</body>
</html>"""

def construir_ruta_fila(row, orden_carpetas):
    partes = ["PPT"]
    for col in orden_carpetas:
        if col in row:
            partes.append(limpiar_nombre_carpeta(row[col]))
    return os.path.join(*partes)

@app.route('/')
def index():
    config, df, ruta_activa = cargar_datos_congreso()
    if config is None or df is None:
        return render_template_string(HTML_APP, config=None, equipo_actual=platform.node())

    columnas = list(df.columns)
    col_sala = next((c for c in columnas if str(c).strip().lower() == 'sala'), None)
    orden = config.get("orden_carpetas", [])

    filas = []
    con_archivo = 0
    for _, row in df.iterrows():
        ruta_rel = construir_ruta_fila(row, orden)
        ruta_abs = os.path.join(ruta_activa, ruta_rel)
        archivos = [f for f in os.listdir(ruta_abs) if not f.startswith('.')] if os.path.exists(ruta_abs) else []
        tiene_ppt = len(archivos) > 0
        if tiene_ppt:
            con_archivo += 1
        filas.append({
            "datos": {c: str(row[c]) for c in columnas},
            "sala": str(row[col_sala]).strip() if col_sala else "",
            "ruta_relativa": ruta_rel,
            "tiene_ppt": tiene_ppt,
            "num_archivos": len(archivos)
        })

    return render_template_string(
        HTML_APP,
        config=config,
        columnas=columnas,
        filas=filas,
        total_ponencias=len(filas),
        con_archivo=con_archivo,
        ruta_activa=ruta_activa,
        equipo_actual=platform.node()
    )

@app.route('/api/crear-carpetas', methods=['POST'])
def api_crear_carpetas():
    config, df, ruta_activa = cargar_datos_congreso()
    if config is None:
        return jsonify({"ok": False, "mensaje": "Faltan programa.xlsx o config_congreso.json"})
    orden = config.get("orden_carpetas", [])
    creadas = 0
    for _, row in df.iterrows():
        ruta_abs = os.path.join(ruta_activa, construir_ruta_fila(row, orden))
        if not os.path.exists(ruta_abs):
            os.makedirs(ruta_abs, exist_ok=True)
            creadas += 1
    return jsonify({"ok": True, "mensaje": f"Estructura verificada. Se crearon {creadas} carpetas nuevas dentro de PPT/."})

@app.route('/api/abrir-ruta', methods=['POST'])
def api_abrir_ruta():
    _, _, ruta_activa = cargar_datos_congreso()
    ruta_rel = request.json.get("ruta", "PPT")
    ruta_abs = os.path.join(ruta_activa, ruta_rel)
    os.makedirs(ruta_abs, exist_ok=True)
    if SISTEMA == 'Windows':
        os.startfile(ruta_abs)
    elif SISTEMA == 'Darwin':
        subprocess.Popen(["open", ruta_abs])
    else:
        subprocess.Popen(["xdg-open", ruta_abs])
    return jsonify({"ok": True})

@app.route('/api/configurar-equipo', methods=['POST'])
def api_configurar_equipo():
    config, _, _ = cargar_datos_congreso()
    equipo = request.json.get("equipo", "").strip()
    wg = config.get("workgroup", "CONGRESO") if config else "CONGRESO"
    try:
        if SISTEMA == 'Windows':
            cmd = f'powershell -Command "Rename-Computer -NewName \'{equipo}\' -Force; Add-Computer -WorkgroupName \'{wg}\' -Force"'
            subprocess.Popen(cmd, shell=True)
            return jsonify({"ok": True, "mensaje": f"Orden enviada en Windows: Equipo '{equipo}' en grupo '{wg}' (reinicia para aplicar)."})
        elif SISTEMA == 'Darwin':
            subprocess.Popen(["scutil", "--set", "ComputerName", equipo])
            subprocess.Popen(["scutil", "--set", "LocalHostName", re.sub(r'[^a-zA-Z0-9-]', '-', equipo)])
            return jsonify({"ok": True, "mensaje": f"Nombre de equipo configurado en macOS como '{equipo}'."})
    except Exception as e:
        return jsonify({"ok": False, "mensaje": f"No se pudo cambiar automáticamente (requiere permisos de administrador): {e}"})
    return jsonify({"ok": True, "mensaje": f"Puesto activo establecido: {equipo} ({wg})."})

if __name__ == '__main__':
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5051")).start()
    app.run(host='127.0.0.1', port=5051)
