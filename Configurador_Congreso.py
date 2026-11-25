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
        # Evitar error Read-only file system si macOS ejecuta en AppTranslocation
        if 'AppTranslocation' in ruta or not os.access(ruta, os.W_OK):
            return os.path.join(os.path.expanduser("~"), "Desktop")
        return ruta
    return os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)

HTML_CONFIG = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Configurador de Congreso</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        body { background: #f0f2f5; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
        .ruta-preview { background: #1e293b; color: #38bdf8; font-family: monospace; padding: 10px 14px; border-radius: 8px; font-size: 0.9rem; }
    </style>
</head>
<body class="py-4">
<div class="container" style="max-width: 780px;">
    <div class="card shadow border-0 rounded-4 overflow-hidden">
        <div class="bg-dark text-white p-4">
            <h4 class="mb-0 fw-bold">⚙️ Configurador de Paquete USB - Congreso</h4>
            <small class="text-secondary">Carga el Excel para detectar las columnas y configurar la estructura de carpetas</small>
        </div>
        <div class="card-body p-4">
            {% if mensaje %}
                <div class="alert alert-success fw-bold">{{ mensaje }}</div>
            {% endif %}
            {% if error %}
                <div class="alert alert-danger fw-bold">{{ error }}</div>
            {% endif %}

            <form method="POST" enctype="multipart/form-data">
                <div class="mb-3">
                    <label class="form-label fw-bold">1. Archivo Excel del Congreso (.xlsx):</label>
                    <input type="file" id="inputExcel" name="excel" class="form-control" accept=".xlsx,.xls" required onchange="leerColumnasExcel(this)">
                </div>

                <div class="row g-3 mb-4">
                    <div class="col-md-6">
                        <label class="form-label fw-bold">2. Grupo de Trabajo (Red):</label>
                        <input type="text" name="workgroup" value="CONGRESO" class="form-control" required>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-bold">3. Nº de ordenadores Admin (ADM1, ADM2...):</label>
                        <input type="number" name="num_admins" value="2" min="1" max="20" class="form-control" required>
                    </div>
                </div>

                <!-- 4. Desplegables dinámicos de jerarquía de carpetas -->
                <div class="mb-3">
                    <label class="form-label fw-bold">4. Orden de subcarpetas dentro de PPT (Selecciona las columnas del Excel):</label>
                    <div class="row g-2 mb-2">
                        <div class="col-md-3">
                            <label class="small text-muted fw-semibold">Nivel 1 (Raíz PPT)</label>
                            <select name="nivel_1" id="nivel_1" class="form-select form-select-sm col-selector" onchange="actualizarVistaPrevia()">
                                <option value="Sala">Sala</option>
                            </select>
                        </div>
                        <div class="col-md-3">
                            <label class="small text-muted fw-semibold">Nivel 2</label>
                            <select name="nivel_2" id="nivel_2" class="form-select form-select-sm col-selector" onchange="actualizarVistaPrevia()">
                                <option value="Dia">Dia</option>
                            </select>
                        </div>
                        <div class="col-md-3">
                            <label class="small text-muted fw-semibold">Nivel 3</label>
                            <select name="nivel_3" id="nivel_3" class="form-select form-select-sm col-selector" onchange="actualizarVistaPrevia()">
                                <option value="Titulo de sesion">Titulo de sesion</option>
                            </select>
                        </div>
                        <div class="col-md-3">
                            <label class="small text-muted fw-semibold">Nivel 4 (Interior)</label>
                            <select name="nivel_4" id="nivel_4" class="form-select form-select-sm col-selector" onchange="actualizarVistaPrevia()">
                                <option value="Participantes">Participantes</option>
                            </select>
                        </div>
                    </div>
                    <div class="ruta-preview" id="vistaPrevia">
                        📁 Estructura resultante: PPT / Sala / Dia / Titulo de sesion / Participantes / archivo.pptx
                    </div>
                </div>

                <div class="mb-4">
                    <label class="form-label fw-bold">5. Salas adicionales (opcional, las del Excel se detectan solas):</label>
                    <textarea name="equipos_extra" rows="2" class="form-control" placeholder="Escribe aquí otros nombres de equipo si los necesitas (uno por línea)"></textarea>
                </div>

                <button type="submit" class="btn btn-success btn-lg w-100 fw-bold">📦 GENERAR CARPETA LISTA PARA PENDRIVE</button>
            </form>
        </div>
    </div>
</div>

<script>
function normalizar(texto) {
    return texto.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

function leerColumnasExcel(input) {
    if (!input.files || !input.files[0]) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
        
        if (rows.length > 0) {
            const headers = rows[0].filter(h => h !== undefined && h !== null && String(h).trim() !== '').map(h => String(h).trim());
            rellenarDesplegables(headers);
        }
    };
    reader.readAsArrayBuffer(input.files[0]);
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
    
    const ruta = niveles.length > 0 ? 'PPT / ' + niveles.join(' / ') + ' / archivo.pptx' : 'PPT / archivo.pptx';
    document.getElementById('vistaPrevia').textContent = '📁 Estructura resultante: ' + ruta;
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

                # Detectar columna de Sala sin importar mayúsculas/tildes
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

                # Buscar CongresoApp en la carpeta de la app o en el Escritorio
                rutas_busqueda = [base, os.path.dirname(os.path.abspath(sys.executable))]
                if SISTEMA == 'Darwin':
                    rutas_busqueda.append(os.path.abspath(os.path.join(os.path.dirname(sys.executable), '..', '..', '..')))

                for r_busq in rutas_busqueda:
                    exe_win = os.path.join(r_busq, "CongresoApp.exe")
                    app_mac = os.path.join(r_busq, "CongresoApp.app")
                    if os.path.exists(exe_win) and not os.path.exists(os.path.join(salida, "CongresoApp.exe")):
                        shutil.copy2(exe_win, os.path.join(salida, "CongresoApp.exe"))
                    if os.path.exists(app_mac) and not os.path.exists(os.path.join(salida, "CongresoApp.app")):
                        shutil.copytree(app_mac, os.path.join(salida, "CongresoApp.app"))

                mensaje = f"✔ ¡Paquete generado con éxito en: {salida}! Estructura: {' > '.join(orden_carpetas)}"
                if SISTEMA == 'Windows':
                    os.startfile(salida)
                elif SISTEMA == 'Darwin':
                    subprocess.Popen(["open", salida])
            except Exception as e:
                error = f"Error procesando el archivo: {str(e)}"

    return render_template_string(HTML_CONFIG, mensaje=mensaje, error=error)

if __name__ == '__main__':
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5050")).start()
    app.run(host='127.0.0.1', port=5050)
