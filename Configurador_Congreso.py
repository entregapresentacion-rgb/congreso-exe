import os
import re
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
            padding: 9px 12px;
            font-size: 0.88rem;
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
        .nivel-box {
            background-color: #171f32;
            border: 1px solid var(--border-color);
            border-radius: 6px;
            padding: 10px;
            height: 100%;
        }
        .multi-cols-box {
            max-height: 155px;
            overflow-y: auto;
            background-color: #141b2d;
            border: 1px solid rgba(255, 255, 255, 0.12);
            border-radius: 6px;
            padding: 6px 10px;
        }
        .form-check-input:checked {
            background-color: #0162e8;
            border-color: #00b9ff;
        }
        .ruta-preview {
            background-color: #101624;
            border: 1px solid rgba(0, 185, 255, 0.25);
            color: #38bdf8;
            font-family: monospace;
            padding: 12px 16px;
            border-radius: 6px;
            font-size: 0.85rem;
            display: flex;
            align-items: flex-start;
            gap: 10px;
            word-break: break-word;
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
        <span class="subnav-item active"><i class="bi bi-sliders"></i> Configurador USB</span>
        <span class="subnav-item"><i class="bi bi-folder-symlink"></i> Nivel 3 (≤50 car.) + Participantes</span>
    </nav>

    <div class="container-fluid px-4 pt-4" style="max-width: 1280px;">
        <div class="mb-4">
            <h4 class="fw-bold mb-1">Configurador de Paquete USB</h4>
            <p class="mb-0" style="color: var(--text-muted); font-size: 0.88rem;">
                Nivel 3 por defecto con <strong>Hora inicio</strong>, <strong>Hora fin</strong>, <strong>Título sesión</strong>, <strong>Sala</strong> y <strong>Día</strong> · Sanitizado, máx. <strong>50 caracteres</strong> y subcarpeta por participante
            </p>
        </div>

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
                    <div class="kpi-title">Nivel 3 (Máx. 50 car.)</div>
                    <div class="kpi-value" id="kpiNivel3Cols">5 Campos</div>
                    <div class="kpi-sub" id="kpiNivel3Resumen">Hora inicio · Hora fin · Título sesión · Sala · Día</div>
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
                <small>Copia toda la carpeta <strong>PAQUETE_USB_CONGRESO</strong> al pendrive y abre <strong>CongresoApp</strong> en cada equipo.</small>
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
                                <div class="valex-card-subtitle">Configura las columnas y el nombre sanitizado de sesión (máx. 50 caracteres)</div>
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
                                <label class="form-label">4. Estructura de carpetas dentro de PPT</label>
                                <div class="row g-3 mb-3">
                                    <div class="col-md-3">
                                        <div class="nivel-box">
                                            <small class="d-block text-info fw-semibold mb-1" style="font-size:0.74rem;">NIVEL 1 (Raíz)</small>
                                            <select name="nivel_1" id="nivel_1" class="form-select form-select-sm" onchange="actualizarVistaPrevia()">
                                                <option value="SALA">SALA</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div class="col-md-3">
                                        <div class="nivel-box">
                                            <small class="d-block text-info fw-semibold mb-1" style="font-size:0.74rem;">NIVEL 2</small>
                                            <select name="nivel_2" id="nivel_2" class="form-select form-select-sm" onchange="actualizarVistaPrevia()">
                                                <option value="DÍA">DÍA</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div class="col-md-3">
                                        <div class="nivel-box" style="border-color: rgba(0, 185, 255, 0.35);">
                                            <small class="d-block text-info fw-semibold mb-1" style="font-size:0.74rem;">NIVEL 3 (Sesión · ≤50 car.)</small>
                                            <div id="nivel3Checks" class="multi-cols-box mb-2">
                                                <div class="form-check">
                                                    <input class="form-check-input nivel3-chk" type="checkbox" name="nivel_3_cols" value="HORA INICIO" id="n3_def_1" checked onchange="actualizarVistaPrevia()">
                                                    <label class="form-check-label small" for="n3_def_1">HORA INICIO</label>
                                                </div>
                                                <div class="form-check">
                                                    <input class="form-check-input nivel3-chk" type="checkbox" name="nivel_3_cols" value="HORA FIN" id="n3_def_2" checked onchange="actualizarVistaPrevia()">
                                                    <label class="form-check-label small" for="n3_def_2">HORA FIN</label>
                                                </div>
                                                <div class="form-check">
                                                    <input class="form-check-input nivel3-chk" type="checkbox" name="nivel_3_cols" value="TÍTULO SESIÓN" id="n3_def_3" checked onchange="actualizarVistaPrevia()">
                                                    <label class="form-check-label small" for="n3_def_3">TÍTULO SESIÓN</label>
                                                </div>
                                                <div class="form-check">
                                                    <input class="form-check-input nivel3-chk" type="checkbox" name="nivel_3_cols" value="SALA" id="n3_def_4" checked onchange="actualizarVistaPrevia()">
                                                    <label class="form-check-label small" for="n3_def_4">SALA</label>
                                                </div>
                                                <div class="form-check">
                                                    <input class="form-check-input nivel3-chk" type="checkbox" name="nivel_3_cols" value="DÍA" id="n3_def_5" checked onchange="actualizarVistaPrevia()">
                                                    <label class="form-check-label small" for="n3_def_5">DÍA</label>
                                                </div>
                                            </div>
                                            <label class="small text-secondary d-block mb-1" style="font-size:0.7rem;">Separador:</label>
                                            <input type="text" name="separador_nivel_3" id="separador_nivel_3" value=" - " class="form-control form-control-sm" oninput="actualizarVistaPrevia()">
                                        </div>
                                    </div>
                                    <div class="col-md-3">
                                        <div class="nivel-box">
                                            <small class="d-block text-info fw-semibold mb-1" style="font-size:0.74rem;">NIVEL 4 (Participante)</small>
                                            <select name="nivel_4" id="nivel_4" class="form-select form-select-sm" onchange="actualizarVistaPrevia()">
                                                <option value="PARTICIPANTE 1">PARTICIPANTE 1</option>
                                            </select>
                                            <small class="d-block text-secondary mt-2" style="font-size:0.7rem;">Además se crean carpetas para cada participante detectado.</small>
                                        </div>
                                    </div>
                                </div>

                                <div class="ruta-preview" id="vistaPrevia">
                                    <i class="bi bi-folder-fill text-warning mt-1"></i>
                                    <span>PPT / SALA / DÍA / [HORA INICIO - HORA FIN - TÍTULO SESIÓN - SALA - DÍA · ≤50 car.] / [Carpeta por Participante] / archivo.pptx</span>
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

                <div class="col-lg-4">
                    <div class="valex-card h-100">
                        <div class="valex-card-header">
                            <div>
                                <h5 class="valex-card-title">Salas Detectadas</h5>
                                <div class="valex-card-subtitle">Previsualización desde el Excel</div>
                            </div>
                            <i class="bi bi-door-open text-info fs-5"></i>
                        </div>
                        <div class="p-4">
                            <div id="listaSalasPreview" class="d-flex flex-wrap gap-2 mb-4">
                                <span style="color: var(--text-muted); font-size: 0.82rem;">Sube el Excel para ver las salas aquí.</span>
                            </div>
                            <hr style="border-color: var(--border-color);">
                            <p class="small mb-2 text-white fw-semibold">Reglas aplicadas automáticamente:</p>
                            <ul class="small mb-0 ps-3" style="color: var(--text-muted);">
                                <li>Sanitización de caracteres especiales en carpetas.</li>
                                <li>Nivel 3 truncado a un máximo de <strong>50 caracteres</strong>.</li>
                                <li>Carpeta individual por cada participante de la sesión.</li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
        </form>
    </div>

<script>
const DEFECTO_N3 = [
    ['hora inicio', 'hora inicio sesion', 'inicio'],
    ['hora fin', 'hora final', 'fin'],
    ['titulo sesion', 'titulo de sesion', 'titulo sesion es', 'sesion', 'titulo'],
    ['sala'],
    ['dia', 'fecha']
];

function normalizar(texto) {
    return String(texto || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, " ").trim();
}

function coincideConLista(colNorm, patrones) {
    return patrones.some(p => colNorm === p || colNorm.includes(p));
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
            document.getElementById('kpiFilas').textContent = (rows.length - 1) + ' filas';
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
        { id: 'nivel_4', claves: ['participante 1', 'participante', 'ponente', 'ponentes'] }
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
            if (!seleccionada && coincideConLista(colNorm, sel.claves)) {
                opt.selected = true;
                seleccionada = true;
            }
            el.appendChild(opt);
        });
    });

    const n3Container = document.getElementById('nivel3Checks');
    let marcadaAlguna = false;
    n3Container.innerHTML = columnas.map((col, idx) => {
        const colNorm = normalizar(col);
        const esPorDefecto = DEFECTO_N3.some(grupo => coincideConLista(colNorm, grupo));
        if (esPorDefecto) marcadaAlguna = true;
        return `
            <div class="form-check">
                <input class="form-check-input nivel3-chk" type="checkbox" name="nivel_3_cols" value="${col.replace(/"/g, '&quot;')}" id="n3_${idx}" ${esPorDefecto ? 'checked' : ''} onchange="actualizarVistaPrevia()">
                <label class="form-check-label small" for="n3_${idx}">${col}</label>
            </div>
        `;
    }).join('');

    if (!marcadaAlguna && n3Container.querySelector('.nivel3-chk')) {
        n3Container.querySelector('.nivel3-chk').checked = true;
    }

    actualizarVistaPrevia();
}

function obtenerNivel3Seleccionado() {
    const cols = Array.from(document.querySelectorAll('.nivel3-chk:checked')).map(chk => chk.value);
    const sep = document.getElementById('separador_nivel_3').value;
    let texto = cols.join(sep);
    if (texto.length > 50) {
        texto = texto.substring(0, 50) + '…';
    }
    return { cols, sep, texto };
}

function actualizarVistaPrevia() {
    const n1 = document.getElementById('nivel_1').value;
    const n2 = document.getElementById('nivel_2').value;
    const n3 = obtenerNivel3Seleccionado();
    const n4 = document.getElementById('nivel_4').value;

    document.getElementById('kpiNivel3Cols').textContent = n3.cols.length + (n3.cols.length === 1 ? ' Campo' : ' Campos');
    document.getElementById('kpiNivel3Resumen').textContent = n3.cols.length > 0 ? n3.cols.join(' · ') : 'Omitido';

    const partes = ['PPT'];
    if (n1) partes.push(n1);
    if (n2) partes.push(n2);
    if (n3.cols.length > 0) partes.push(`[${n3.texto} · ≤50 car.]`);
    partes.push(n4 ? `[Participantes / ${n4}]` : '[Carpeta por Participante]');
    partes.push('archivo.pptx');

    document.getElementById('vistaPrevia').innerHTML = `<i class="bi bi-folder-fill text-warning mt-1"></i><span>${partes.join(' / ')}</span>`;
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

                col_sala = next((c for c in df.columns if re.sub(r'[^a-z0-9]+', '', str(c).lower()) == 'sala'), None)
                if col_sala:
                    salas = sorted(list(set(sanitizar_nombre_carpeta(s) for s in df[col_sala].dropna() if str(s).strip())))
                    equipos.extend(salas)

                extras = [sanitizar_nombre_carpeta(l) for l in request.form.get('equipos_extra', '').splitlines() if l.strip()]
                for ex in extras:
                    if ex not in equipos:
                        equipos.append(ex)

                n1 = request.form.get('nivel_1', '').strip()
                n2 = request.form.get('nivel_2', '').strip()
                n3_cols = [c.strip() for c in request.form.getlist('nivel_3_cols') if c.strip()]
                sep_n3 = request.form.get('separador_nivel_3', ' - ')
                n4 = request.form.get('nivel_4', '').strip()

                orden_carpetas = []
                for item in [n1, n2, n3_cols if n3_cols else "", n4]:
                    if item and item not in orden_carpetas:
                        orden_carpetas.append(item)

                config_data = {
                    "workgroup": sanitizar_nombre_carpeta(request.form.get('workgroup', 'CONGRESO').strip().upper(), max_len=15),
                    "orden_carpetas": orden_carpetas,
                    "nivel_3_columnas": n3_cols,
                    "separador_nivel_3": sep_n3,
                    "max_chars_nivel_3": 50,
                    "equipos": equipos
                }

                with open(os.path.join(salida, "config_congreso.json"), "w", encoding="utf-8") as f:
                    json.dump(config_data, f, indent=4, ensure_ascii=False)

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
