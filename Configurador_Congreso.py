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
            return os.path.abspath(os.path.join(ruta, '..', '..', '..'))
        return ruta
    return os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)

HTML_CONFIG = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Configurador de Congreso</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>body { background: #f0f2f5; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }</style>
</head>
<body class="py-4">
<div class="container" style="max-width: 760px;">
    <div class="card shadow border-0 rounded-4 overflow-hidden">
        <div class="bg-dark text-white p-4">
            <h4 class="mb-0 fw-bold">⚙️ Configurador de Paquete USB - Congreso</h4>
            <small class="text-secondary">Genera la configuración y estructura de carpetas desde tu Excel</small>
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
                    <input type="file" name="excel" class="form-control" accept=".xlsx,.xls" required>
                </div>

                <div class="row g-3 mb-3">
                    <div class="col-md-6">
                        <label class="form-label fw-bold">2. Grupo de Trabajo (Red):</label>
                        <input type="text" name="workgroup" value="CONGRESO" class="form-control" required>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-bold">3. Nº de ordenadores Admin (ADM1, ADM2...):</label>
                        <input type="number" name="num_admins" value="2" min="1" max="20" class="form-control" required>
                    </div>
                </div>

                <div class="mb-3">
                    <label class="form-label fw-bold">4. Orden de subcarpetas dentro de PPT (separadas por coma):</label>
                    <input type="text" name="orden_carpetas" value="Sala, Sesion" class="form-control" placeholder="Ej: Fecha, Sala, Sesion">
                    <small class="text-muted">Escribe los nombres exactos de las columnas de tu Excel en el orden deseado.</small>
                </div>

                <div class="mb-4">
                    <label class="form-label fw-bold">5. Salas adicionales (opcional, detectadas automáticamente del Excel):</label>
                    <textarea name="equipos_extra" rows="2" class="form-control" placeholder="Si quieres añadir algún equipo extra manualmente, escríbelo aquí (uno por línea)"></textarea>
                </div>

                <button type="submit" class="btn btn-success btn-lg w-100 fw-bold">📦 GENERAR CARPETA LISTA PARA PENDRIVE</button>
            </form>
        </div>
    </div>
</div>
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

                if 'Sala' in df.columns:
                    salas = sorted(list(set(str(s).strip() for s in df['Sala'].dropna() if str(s).strip())))
                    equipos.extend(salas)

                extras = [l.strip() for l in request.form.get('equipos_extra', '').splitlines() if l.strip()]
                for ex in extras:
                    if ex not in equipos:
                        equipos.append(ex)

                orden_raw = request.form.get('orden_carpetas', 'Sala, Sesion')
                orden_carpetas = [c.strip() for c in orden_raw.split(',') if c.strip()]

                config_data = {
                    "workgroup": request.form.get('workgroup', 'CONGRESO').strip().upper(),
                    "orden_carpetas": orden_carpetas,
                    "equipos": equipos
                }

                with open(os.path.join(salida, "config_congreso.json"), "w", encoding="utf-8") as f:
                    json.dump(config_data, f, indent=4, ensure_ascii=False)

                exe_win = os.path.join(base, "CongresoApp.exe")
                app_mac = os.path.join(base, "CongresoApp.app")
                if os.path.exists(exe_win):
                    shutil.copy2(exe_win, os.path.join(salida, "CongresoApp.exe"))
                if os.path.exists(app_mac):
                    shutil.copytree(app_mac, os.path.join(salida, "CongresoApp.app"))

                mensaje = f"✔ ¡Paquete generado con éxito en: {salida}! Equipos configurados: {', '.join(equipos)}"
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
