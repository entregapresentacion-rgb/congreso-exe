import os
import sys
import json
import shutil
import platform
import threading
import subprocess
import webbrowser
import pandas as pd
from flask import Flask, render_template_string, request

SISTEMA = platform.system()

def obtener_ruta_base():
    if getattr(sys, 'frozen', False):
        ruta = os.path.dirname(sys.executable)
        if SISTEMA == 'Darwin' and ruta.endswith('Contents/MacOS'):
            ruta = os.path.abspath(os.path.join(ruta, '..', '..', '..'))
        if 'AppTranslocation' in ruta or not os.access(ruta, os.W_OK):
            return os.path.join(os.path.expanduser("~"), "Desktop")
        return ruta
    return os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)

HTML_CONFIG = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Valex Congreso | Configurador USB</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
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
            gap: 24px;
        }
        .subnav-item {
            padding: 12px 4px;
            color: var(--text-muted);
            font-size: 0.875rem;
            font-weight: 500;
            text-decoration: none;
            display: flex;
            align-items: center;
            gap: 7px;
            border-bottom: 2px solid transparent;
        }
        .subnav-item.active {
            color: #38bdf8;
            border-bottom-color: #0162e8;
        }
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
        .kpi-value { font-size: 1.5rem; font-weight: 700; line-height: 1.2; }
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
        }
        .valex-card-title { font-size: 0.95rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin: 0; }
        .valex-card-subtitle { font-size: 0.8rem; color: var(--text-muted); margin-top: 2px; }
        .form-label { color: #cbd5e1; font-size: 0.85rem; font-weight: 500; margin-bottom: 6px; }
        .form-control, .form-select {
            background-color: var(--bg-input) !important;
            border: 1px solid rgba(255, 255, 255, 0.12) !important;
            color: #ffffff !important;
            border-radius: 6px;
            padding: 10px 14px;
            font-size: 0.9rem;
        }
        .form-control:focus, .form-select:focus {
            border-color: #0162e8 !important;
            box-shadow: 0 0 0 0.2rem rgba(1, 98, 232, 0.25) !important;
        }
        input[type="file"]::file-selector-button {
            background: linear-gradient(to right, #0162e8, #00b9ff);
            color: white;
            border: none;
            padding: 6px 14px;
            border-radius: 4px;
            margin-right: 12px;
            font-weight: 500;
            cursor: pointer;
        }
        .nivel-box { background-color: #171f32; border: 1px solid var(--border-color); border-radius: 6px; padding: 10px; }
        .ruta-preview {
            background-color: #101624;
            border: 1px solid rgba(0, 185, 255, 0.25);
            color: #38bdf8;
            font-family: monospace;
            padding: 12px 16px;
            border-radius: 6px;
            font-size: 0.86rem;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .btn-valex-primary {
            background: linear-gradient(to right, #0162e8 0%, #00b9ff 100%);
            border: none;
            color: #fff;
            font-weight: 600;
            padding: 13px 20px;
            border-radius: 6px;
            box-shadow: 0 4px 15px rgba(1, 98, 232, 0.35);
        }
        .btn-valex-primary:hover { color: #fff; opacity: 0.95; }
        .badge-sala {
            background-color: rgba(0, 185, 255, 0.15);
            color: #38bdf8;
            border: 1px solid rgba(0, 185, 255, 0.3);
            padding: 5px 10px;
            border-radius: 4px;
            font-size: 0.78rem;
        }
        .step-badge {
            width: 26px;
            height: 26px;
            border-radius: 50%;
            background: linear-gradient(135deg, #0162e8, #00b9ff);
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            font-size: 0.8rem;
            flex-shrink: 0;
        }
    </style>
</head>
<body>
    <header class="valex-topbar d-flex justify-content-between align-items-center">
        <a href="/" class="brand-logo">
            <div class="brand-icon"><i class="bi bi-diagram-3-fill"></i></div>
            <span>Congresos<span style="color:#00b9ff;">USB</span></span>
        </a>
        <span class="badge bg-dark border border-secondary text-light px-3 py-2">
            <i class="bi bi-pc-display me-1 text-info"></i> Sistema: {{ sistema }}
        </span>
    </header>

    <nav class="valex-subnav">
        <span class="subnav-item active"><i class="bi bi-sliders"></i> Paso 1: Configurador USB</span>
        <span class="subnav-item"><i class="bi bi-play-circle"></i> Paso 2: Ejecutar CongresoApp en los equipos</span>
    </nav>

    <div class="container-fluid px-4 pt-4" style="max-width: 1220px;">
        <div class="mb-4">
            <h4 class="fw-bold mb-1">Configurador de Paquete USB</h4>
            <p class="mb-0" style="color: var(--text-muted); font-size: 0.88rem;">
                Prepara los archivos maestros (<code>programa.xlsx</code> y <code>config_congreso.json</code>) que usará <strong>CongresoApp</strong> en cada ordenador
            </p>
        </div>

        <!-- 4 Tarjetas KPI -->
        <div class="row g-3 mb-4">
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-blue">
                    <div class="kpi-title">Archivo Programa</div>
                    <div class="kpi-value" id="kpiArchivo">Sin cargar</div>
                    <div class="kpi-sub" id="kpiFilas">Selecciona un .xlsx</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,30 L30,22 L60,28 L95,12 L130,24 L165,10 L200,18 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-pink">
                    <div class="kpi-title">Salas Detectadas</div>
                    <div class="kpi-value" id="kpiSalas">0 Salas</div>
                    <div class="kpi-sub">Detección automática</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,25 L40,32 L80,15 L120,25 L160,12 L200,22 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-green">
                    <div class="kpi-title">Niveles Carpeta PPT</div>
                    <div class="kpi-value" id="kpiNiveles">4 Niveles</div>
                    <div class="kpi-sub">Árbol de ponencias</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,32 L35,18 L75,26 L115,10 L155,20 L200,8 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-orange">
                    <div class="kpi-title">Puestos Admin</div>
                    <div class="kpi-value" id="kpiAdmins">2 Equipos</div>
                    <div class="kpi-sub" id="kpiWorkgroup">Grupo: CONGRESO</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,20 L45,28 L85,14 L130,26 L170,15 L200,24 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
        </div>

        {% if mensaje %}
            <div class="alert mb-4" style="background: rgba(2,150,102,0.2); border: 1px solid #48d6a8; color: #6ee7b7;">
                <div class="fw-bold mb-1"><i class="bi bi-check-circle-fill me-2"></i>{{ mensaje }}</div>
                <small>Ahora copia toda la carpeta <strong>PAQUETE_USB_CONGRESO</strong> a tu pendrive y abre <strong>CongresoApp</strong> desde ahí en cada ordenador.</small>
            </div>
        {% endif %}
        {% if error %}
            <div class="alert mb-4" style="background: rgba(249,58,90,0.2); border: 1px solid #f7778c; color: #fda4af;">
                <i class="bi bi-exclamation-triangle-fill me-2"></i>{{ error }}
            </div>
        {% endif %}

        <form method="POST" enctype="multipart/form-data">
            <div class="row g-4">
                <div class="col-lg-8">
                    <div class="valex-card">
                        <div class="valex-card-header">
                            <div>
                                <h5 class="valex-card-title">Parámetros del Congreso</h5>
                                <div class="valex-card-subtitle">Configura las columnas del Excel y los equipos de red</div>
                            </div>
                            <i class="bi bi-sliders text-info fs-5"></i>
                        </div>
                        <div class="p-4">
                            <div class="mb-4">
                                <label class="form-label">1. Archivo Excel del Congreso (.xlsx / .xls)</label>
                                <input type="file" id="inputExcel" name="excel" class="form-control" accept=".xlsx,.xls" required onchange="leerColumnasExcel(this)">
                            </div>

                            <div class="row g-3 mb-4">
                                <div class="col-md-6">
                                    <label class="form-label">2. Grupo de Trabajo (Red LAN)</label>
                                    <input type="text" name="workgroup" id="inputWorkgroup" value="CONGRESO" class="form-control" required oninput="actualizarKPIs()">
                                </div>
                                <div class="col-md-6">
                                    <label class="form-label">3. Nº de ordenadores Admin (ADM1, ADM2...)</label>
                                    <input type="number" name="num_admins" id="inputAdmins" value="2" min="1" max="20" class="form-control" required oninput="actualizarKPIs()">
                                </div>
                            </div>

                            <div class="mb-4">
                                <label class="form-label">4. Orden de subcarpetas dentro de PPT</label>
                                <div class="row g-2 mb-3">
                                    <div class="col-md-3">
                                        <div class="nivel-box">
                                            <small class="d-block text-info fw-semibold mb-1" style="font-size:0.74rem;">NIVEL 1</small>
                                            <select name="nivel_1" id="nivel_1" class="form-select form-select-sm" onchange="actualizarVistaPrevia()">
                                                <option value="Sala">Sala</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div class="col-md-3">
                                        <div class="nivel-box">
                                            <small class="d-block text-info fw-semibold mb-1" style="font-size:0.74rem;">NIVEL 2</small>
                                            <select name="nivel_2" id="nivel_2" class="form-select form-select-sm" onchange="actualizarVistaPrevia()">
                                                <option value="Dia">Dia</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div class="col-md-3">
                                        <div class="nivel-box">
                                            <small class="d-block text-info fw-semibold mb-1" style="font-size:0.74rem;">NIVEL 3</small>
                                            <select name="nivel_3" id="nivel_3" class="form-select form-select-sm" onchange="actualizarVistaPrevia()">
                                                <option value="Titulo de sesion">Titulo de sesion</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div class="col-md-3">
                                        <div class="nivel-box">
                                            <small class="d-block text-info fw-semibold mb-1" style="font-size:0.74rem;">NIVEL 4</small>
                                            <select name="nivel_4" id="nivel_4" class="form-select form-select-sm" onchange="actualizarVistaPrevia()">
                                                <option value="Participantes">Participantes</option>
                                            </select>
                                        </div>
                                    </div>
                                </div>
                                <div class="ruta-preview" id="vistaPrevia">
                                    <i class="bi bi-folder-fill text-warning"></i>
                                    <span>PPT / Sala / Dia / Titulo de sesion / Participantes / archivo.pptx</span>
                                </div>
                            </div>

                            <div class="mb-4">
                                <label class="form-label">5. Equipos o Salas adicionales (Opcional)</label>
                                <textarea name="equipos_extra" rows="2" class="form-control" placeholder="Escribe otros nombres de equipo si los necesitas (uno por línea)..."></textarea>
                            </div>

                            <button type="submit" class="btn btn-valex-primary w-100">
                                <i class="bi bi-box-seam-fill me-2"></i> GENERAR CARPETA LISTA PARA PENDRIVE
                            </button>
                        </div>
                    </div>
                </div>

                <!-- Columna Derecha: Flujo de uso + Salas -->
                <div class="col-lg-4">
                    <div class="valex-card h-100">
                        <div class="valex-card-header">
                            <div>
                                <h5 class="valex-card-title">¿Cómo usar los archivos?</h5>
                                <div class="valex-card-subtitle">Guía rápida en 3 pasos</div>
                            </div>
                            <i class="bi bi-info-circle text-info fs-5"></i>
                        </div>
                        <div class="p-4 d-flex flex-column gap-3" style="font-size: 0.85rem;">
                            <div class="d-flex gap-3">
                                <span class="step-badge">1</span>
                                <div>
                                    <strong class="text-white d-block">Genera el paquete</strong>
                                    <span style="color: var(--text-muted);">Se creará <code>PAQUETE_USB_CONGRESO</code> con <code>programa.xlsx</code>, <code>config_congreso.json</code> y <code>CongresoApp</code>.</span>
                                </div>
                            </div>
                            <div class="d-flex gap-3">
                                <span class="step-badge">2</span>
                                <div>
                                    <strong class="text-white d-block">Copia al Pendrive USB</strong>
                                    <span style="color: var(--text-muted);">Mete toda esa carpeta en tu USB y llévalo a cada ordenador del congreso (ADM1, Sala 1, Sala 2...).</span>
                                </div>
                            </div>
                            <div class="d-flex gap-3">
                                <span class="step-badge">3</span>
                                <div>
                                    <strong class="text-white d-block">Abre CongresoApp en cada PC</strong>
                                    <span style="color: var(--text-muted);">CongresoApp leerá automáticamente el Excel y el JSON para crear las carpetas PPT, configurar la red y gestionar las presentaciones.</span>
                                </div>
                            </div>

                            <hr style="border-color: var(--border-color); my-1">

                            <div>
                                <small class="text-uppercase fw-bold d-block mb-2" style="color: var(--text-muted); font-size: 0.72rem;">
                                    Salas detectadas en tu Excel
                                </small>
                                <div id="listaSalasPreview" class="d-flex flex-wrap gap-2">
                                    <span style="color: var(--text-muted); font-size: 0.82rem;">Sube el Excel para verlas aquí.</span>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </form>
    </div>

<script>
function normalizar(texto) {
    return texto.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}
function actualizarKPIs() {
    const admins = document.getElementById('inputAdmins').value || 2;
    const wg = (document.getElementById('inputWorkgroup').value || 'CONGRESO').toUpperCase();
    document.getElementById('kpiAdmins').textContent = admins + (admins == 1 ? ' Equipo' : ' Equipos');
    document.getElementById('kpiWorkgroup').textContent = 'Grupo: ' + wg;
}
function leerColumnasExcel(input) {
    if (!input.files || !input.files[0]) return;
    const file = input.files[0];
    document.getElementById('kpiArchivo').textContent = file.name.length > 16 ? file.name.substring(0, 14) + '...' : file.name;
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1});
        if (rows.length > 0) {
            document.getElementById('kpiFilas').textContent = (rows.length - 1) + ' ponencias/filas';
            const headers = rows[0].filter(h => h !== undefined && h !== null && String(h).trim() !== '').map(h => String(h).trim());
            rellenarDesplegables(headers);
            const idxSala = headers.findIndex(h => normalizar(h) === 'sala');
            const contenedorSalas = document.getElementById('listaSalasPreview');
            if (idxSala !== -1) {
                const salasSet = new Set();
                for (let i = 1; i < rows.length; i++) {
                    if (rows[i][idxSala]) salasSet.add(String(rows[i][idxSala]).trim());
                }
                const salas = Array.from(salasSet).filter(Boolean).sort();
                document.getElementById('kpiSalas').textContent = salas.length + (salas.length === 1 ? ' Sala' : ' Salas');
                contenedorSalas.innerHTML = salas.length > 0
                    ? salas.map(s => `<span class="badge-sala"><i class="bi bi-door-open me-1"></i>${s}</span>`).join('')
                    : '<span style="color: var(--text-muted); font-size:0.82rem;">Sin salas rellenas.</span>';
            }
        }
    };
    reader.readAsArrayBuffer(file);
}
function rellenarDesplegables(columnas) {
    const selectores = [
        { id: 'nivel_1', claves: ['sala'] },
        { id: 'nivel_2', claves: ['dia', 'fecha'] },
        { id: 'nivel_3', claves: ['titulo de sesion', 'sesion', 'titulo sesion'] },
        { id: 'nivel_4', claves: ['participantes', 'ponente', 'ponentes', 'autor'] }
    ];
    selectores.forEach(sel => {
        const el = document.getElementById(sel.id);
        el.innerHTML = '<option value="">-- Ninguno (Omitir) --</option>';
        let seleccionada = false;
        columnas.forEach(col => {
            const opt = document.createElement('option');
            opt.value = col;
            opt.textContent = col;
            const colNorm = normalizar(col);
            if (!seleccionada && sel.claves.some(k => colNorm === k || colNorm.includes(k))) {
                opt.selected = true;
                seleccionada = true;
            }
            el.appendChild(opt);
        });
    });
    actualizarVistaPrevia();
}
function actualizarVistaPrevia() {
    const niveles = ['nivel_1', 'nivel_2', 'nivel_3', 'nivel_4']
        .map(id => document.getElementById(id).value)
        .filter(val => val !== '');
    document.getElementById('kpiNiveles').textContent = niveles.length + (niveles.length === 1 ? ' Nivel' : ' Niveles');
    const ruta = niveles.length > 0 ? 'PPT / ' + niveles.join(' / ') + ' / archivo.pptx' : 'PPT / archivo.pptx';
    document.getElementById('vistaPrevia').innerHTML = `<i class="bi bi-folder-fill text-warning"></i><span>${ruta}</span>`;
}
</script>
</body>
</html>"""

@app.route('/', methods=['GET', 'POST'])
def index():
    mensaje = None
    error = None
    if request.method == 'POST':
        archivo = request.files.get('excel')
        if not archivo or archivo.filename == '':
            error = "Debes seleccionar un archivo Excel."
        else:
            try:
                base = obtener_ruta_base()
                salida = os.path.join(base, "PAQUETE_USB_CONGRESO")
                shutil.rmtree(salida, ignore_errors=True)
                os.makedirs(salida, exist_ok=True)

                ruta_excel_dest = os.path.join(salida, "programa.xlsx")
                archivo.save(ruta_excel_dest)

                df = pd.read_excel(ruta_excel_dest, engine='openpyxl')
                num_admins = int(request.form.get('num_admins', 2))
                equipos = [f"ADM{i}" for i in range(1, num_admins + 1)]

                col_sala = next((c for c in df.columns if str(c).strip().lower() == 'sala'), None)
                if col_sala:
                    salas = sorted(list(set(str(s).strip() for s in df[col_sala].dropna() if str(s).strip())))
                    equipos.extend(salas)

                extras = [l.strip() for l in request.form.get('equipos_extra', '').splitlines() if l.strip()]
                for ex in extras:
                    if ex not in equipos:
                        equipos.append(ex)

                orden_carpetas = []
                for n in ['nivel_1', 'nivel_2', 'nivel_3', 'nivel_4']:
                    val = request.form.get(n, '').strip()
                    if val and val not in orden_carpetas:
                        orden_carpetas.append(val)

                config_data = {
                    "workgroup": request.form.get('workgroup', 'CONGRESO').strip().upper(),
                    "orden_carpetas": orden_carpetas,
                    "equipos": equipos
                }

                with open(os.path.join(salida, "config_congreso.json"), "w", encoding="utf-8") as f:
                    json.dump(config_data, f, indent=4, ensure_ascii=False)

                # Copiar CongresoApp (.exe, .app o .py) si está en la misma carpeta
                rutas_busqueda = [base, os.path.dirname(os.path.abspath(sys.executable))]
                if SISTEMA == 'Darwin':
                    rutas_busqueda.append(os.path.abspath(os.path.join(os.path.dirname(sys.executable), '..', '..', '..')))

                for r_busq in rutas_busqueda:
                    for nombre_app in ["CongresoApp.exe", "congreso_app.py"]:
                        src = os.path.join(r_busq, nombre_app)
                        dst = os.path.join(salida, nombre_app)
                        if os.path.exists(src) and not os.path.exists(dst):
                            shutil.copy2(src, dst)
                    app_mac = os.path.join(r_busq, "CongresoApp.app")
                    if os.path.exists(app_mac) and not os.path.exists(os.path.join(salida, "CongresoApp.app")):
                        shutil.copytree(app_mac, os.path.join(salida, "CongresoApp.app"))

                mensaje = f"¡Paquete generado con éxito en: {salida}!"
                if SISTEMA == 'Windows':
                    os.startfile(salida)
                elif SISTEMA == 'Darwin':
                    subprocess.Popen(["open", salida])
            except Exception as e:
                error = f"Error procesando el archivo: {str(e)}"

    return render_template_string(HTML_CONFIG, mensaje=mensaje, error=error, sistema=SISTEMA)

if __name__ == '__main__':
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5050")).start()
    app.run(host='127.0.0.1', port=5050)
