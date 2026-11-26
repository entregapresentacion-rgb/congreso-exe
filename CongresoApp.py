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


def construir_ruta_sesion(row, config, columnas_raw):
    partes = ["PPT"]
    orden_carpetas = config.get("orden_carpetas", [])
    sep_n3 = config.get("separador_nivel_3", " - ")
    max_n3 = int(config.get("max_chars_nivel_3", 50))

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

    tiene_lista_n3 = any(isinstance(x, list) for x in orden_carpetas)

    if orden_carpetas:
        for item in orden_carpetas:
            if isinstance(item, list):
                valores = [str(row[c]).strip() for c in item if c in row and str(row[c]).strip()]
                combinado = sep_n3.join(valores) if valores else "Sin_Especificar"
                partes.append(sanitizar_nombre_carpeta(combinado, max_len=max_n3))
            elif item in row:
                partes.append(sanitizar_nombre_carpeta(row[item]))
        if not tiene_lista_n3 and n3_cols:
            valores = [str(row[c]).strip() for c in n3_cols if c in row and str(row[c]).strip()]
            combinado = sep_n3.join(valores) if valores else "Sin_Especificar"
            partes.append(sanitizar_nombre_carpeta(combinado, max_len=max_n3))
    else:
        col_sala = buscar_columna(columnas_raw, ["sala"])
        col_dia = buscar_columna(columnas_raw, ["dia", "fecha"])
        if col_sala:
            partes.append(sanitizar_nombre_carpeta(row[col_sala]))
        if col_dia:
            partes.append(sanitizar_nombre_carpeta(row[col_dia]))
        if n3_cols:
            valores = [str(row[c]).strip() for c in n3_cols if c in row and str(row[c]).strip()]
            combinado = sep_n3.join(valores) if valores else "Sin_Especificar"
            partes.append(sanitizar_nombre_carpeta(combinado, max_len=max_n3))

    return os.path.join(*partes)


def obtener_participantes_fila(row, columnas):
    participantes = []
    for c in columnas:
        nc = norm_col(c)
        if nc.startswith("participante") or nc.startswith("ponente"):
            val = str(row[c]).strip()
            if val and val not in participantes:
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
                <button class="btn btn-sm btn-valex-blue px-3 py-2" onclick="irACarpetaPonente()">
                    <i class="bi bi-folder-symlink me-1"></i> Ir a Carpeta Ponente
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

    <nav class="valex-subnav">
        <span class="subnav-item active" onclick="cambiarVista('tabla')"><i class="bi bi-table"></i> Tabla y Filtros por Campo</span>
        <span class="subnav-item" onclick="cambiarVista('visor')"><i class="bi bi-collection-play"></i> Visor por Días y Sesiones</span>
    </nav>

    <div class="container-fluid px-4 pt-4" style="max-width: 1480px;">

        {% if not config %}
            <div class="valex-card p-5 text-center my-5">
                <i class="bi bi-exclamation-diamond text-warning display-4 mb-3"></i>
                <h4 class="fw-bold">No se encontraron los archivos de configuración</h4>
                <p style="color: var(--text-muted); max-width: 560px; margin: 10px auto 20px;">
                    Coloca <code>programa.xlsx</code> y <code>config_congreso.json</code> en la misma carpeta que <strong>CongresoApp</strong>.
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
                            Selecciona <strong>ADM...</strong> (puede cargar presentaciones) o una <strong>SALA</strong> (solo lectura y queda bloqueada).
                        {% endif %}
                    </div>
                </div>
                <div class="d-flex align-items-center gap-2">
                    <select id="selectPuesto" class="form-select form-select-sm" style="width: 240px;" {% if puesto_bloqueado %}disabled{% endif %}>
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
                    <div class="kpi-title">Ponencias Visibles</div>
                    <div class="kpi-value" id="kpiVisibles">{{ filas|length }}</div>
                    <div class="kpi-sub">Total en programa: {{ filas|length }}</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,30 L30,22 L60,28 L95,12 L130,24 L165,10 L200,18 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-pink">
                    <div class="kpi-title">Permisos del Equipo</div>
                    <div class="kpi-value">{{ 'ADMIN' if es_admin else 'SOLO VER' }}</div>
                    <div class="kpi-sub">{{ 'Puede cargar presentaciones' if es_admin else 'Solo visualización de sala' }}</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,25 L40,32 L80,15 L120,25 L160,12 L200,22 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-green">
                    <div class="kpi-title">Presentaciones Cargadas</div>
                    <div class="kpi-value">{{ con_archivo }} / {{ filas|length }}</div>
                    <div class="kpi-sub">Detectadas en carpetas PPT</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,32 L35,18 L75,26 L115,10 L155,20 L200,8 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
            <div class="col-sm-6 col-xl-3">
                <div class="kpi-card kpi-orange">
                    <div class="kpi-title">Carpeta Sesión (Nivel 3)</div>
                    <div class="kpi-value">≤ {{ config.max_chars_nivel_3 | default(50) }} Car.</div>
                    <div class="kpi-sub">Sanitizada + Subcarpetas Ponente</div>
                    <svg class="kpi-wave" viewBox="0 0 200 40" preserveAspectRatio="none"><path d="M0,20 L45,28 L85,14 L130,26 L170,15 L200,24 L200,40 L0,40 Z" fill="#ffffff"/></svg>
                </div>
            </div>
        </div>

        <div id="alertaAccion" class="alert d-none mb-4"></div>
        <div id="resultadosPonente" class="valex-card p-3 mb-4 d-none"></div>

        <div id="vistaTabla" class="valex-card">
            <div class="valex-card-header">
                <div>
                    <h5 class="valex-card-title">Programación y Carpetas de Ponentes</h5>
                    <div class="valex-card-subtitle">Abrir Carpeta · Estado · Hora Inicio · Hora Fin · Título Sesión · Sala · Día · ID · Resto</div>
                </div>
                <div class="d-flex gap-2">
                    {% if es_admin %}
                        <button class="btn btn-sm btn-valex-green px-3" onclick="crearArbolCarpetas()">
                            <i class="bi bi-folder-plus me-1"></i> Generar Todas las Carpetas PPT
                        </button>
                    {% endif %}
                    <button class="btn btn-sm btn-valex-blue px-3" onclick="limpiarFiltros()">
                        <i class="bi bi-x-circle me-1"></i> Limpiar Filtros
                    </button>
                </div>
            </div>

            <div class="table-responsive">
                <table class="table table-valex" id="tablaPonencias">
                    <thead>
                        <tr class="header-titles">
                            <th>Abrir Carpeta</th>
                            <th>Estado de Carga</th>
                            {% for col in columnas_ordenadas %}
                                <th>{{ col }}</th>
                            {% endfor %}
                            <th>Carpetas por Ponente</th>
                        </tr>
                        <tr class="header-filters">
                            <th><span class="text-secondary small">Acción</span></th>
                            <th>
                                <select id="flt_estado" class="form-select form-select-sm" style="min-width:110px;" onchange="aplicarFiltros()">
                                    <option value="">Todos</option>
                                    <option value="cargado">Cargada</option>
                                    <option value="pendiente">Pendiente</option>
                                </select>
                            </th>
                            {% for col in columnas_ordenadas %}
                                <th>
                                    <input type="text" class="form-control form-control-sm flt-input flt-col" data-idx="{{ loop.index0 }}" placeholder="Filtrar {{ col }}..." onkeyup="aplicarFiltros()">
                                </th>
                            {% endfor %}
                            <th>
                                <input type="text" id="flt_ponente_col" class="form-control form-control-sm flt-input" placeholder="Filtrar ponente..." onkeyup="aplicarFiltros()">
                            </th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for fila in filas %}
                        <tr
                            data-estado="{{ 'cargado' if fila.tiene_ppt else 'pendiente' }}"
                            data-sala="{{ fila.sala }}"
                            data-ponentes="{{ fila.participantes|join(' | ')|lower }}"
                            {% for col in columnas_ordenadas %}
                                data-c_{{ loop.index0 }}="{{ fila.datos[col]|lower }}"
                            {% endfor %}
                        >
                            <td>
                                {% if es_admin %}
                                    <div class="d-flex flex-nowrap gap-1 align-items-center">
                                        <button class="btn btn-sm btn-valex-blue" onclick='abrirRuta({{ fila.ruta_relativa | tojson }}, true)' title="Abrir carpeta con subcarpetas de ponentes">
                                            <i class="bi bi-folder2-open"></i> Abrir
                                        </button>
                                        <label class="btn btn-sm btn-valex-green mb-0" title="Cargar archivo en esta sesión">
                                            <i class="bi bi-upload"></i>
                                            <input type="file" class="d-none" onchange='subirPresentacion(this, {{ fila.ruta_relativa | tojson }})'>
                                        </label>
                                    </div>
                                {% else %}
                                    <button class="btn btn-sm btn-outline-info" onclick='abrirRuta({{ fila.ruta_relativa | tojson }}, false)'>
                                        <i class="bi bi-eye"></i> Ver
                                    </button>
                                {% endif %}
                            </td>
                            <td>
                                {% if fila.tiene_ppt %}
                                    <span class="badge badge-ppt-ok"><i class="bi bi-check-circle me-1"></i>Cargada ({{ fila.num_archivos }})</span>
                                {% else %}
                                    <span class="badge badge-ppt-wait"><i class="bi bi-clock me-1"></i>Pendiente</span>
                                {% endif %}
                            </td>
                            {% for col in columnas_ordenadas %}
                                <td>{{ fila.datos[col] }}</td>
                            {% endfor %}
                            <td>
                                {% for p in fila.participantes_carpetas %}
                                    <span class="ponente-link" onclick='abrirRuta({{ p.ruta | tojson }}, {{ "true" if es_admin else "false" }})' title="Abrir carpeta de {{ p.nombre }}">
                                        <i class="bi bi-person-folder me-1"></i>{{ p.nombre }}
                                    </span>
                                {% endfor %}
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>

        <div id="vistaVisor" class="d-none">
            {% for dia, sesiones in visor_por_dia.items() %}
                <div class="valex-card mb-4">
                    <div class="valex-card-header">
                        <div>
                            <h5 class="valex-card-title"><i class="bi bi-calendar3 text-info me-2"></i>Día: {{ dia }}</h5>
                            <div class="valex-card-subtitle">{{ sesiones|length }} sesión(es) en este día</div>
                        </div>
                    </div>
                    <div class="p-4">
                        <div class="row g-3">
                            {% for sesion_nombre, items in sesiones.items() %}
                                <div class="col-md-6 col-xl-4">
                                    <div class="p-3 rounded" style="background:#171f32; border:1px solid var(--border-color); height:100%;">
                                        <div class="fw-bold text-info mb-2" style="font-size:0.9rem;">
                                            <i class="bi bi-easel me-1"></i> {{ sesion_nombre }}
                                        </div>
                                        <div class="d-flex flex-column gap-2">
                                            {% for it in items %}
                                                <div class="d-flex justify-content-between align-items-center p-2 rounded" style="background:#1f2940; font-size:0.82rem;">
                                                    <div class="me-2">
                                                        <div class="fw-semibold text-white">{{ it.titulo_corto }}</div>
                                                        <small style="color:var(--text-muted);">{{ it.participantes|join(', ') if it.participantes else 'Sin participante' }}</small>
                                                    </div>
                                                    {% if it.tiene_ppt %}
                                                        <span class="badge badge-ppt-ok">Cargada</span>
                                                    {% else %}
                                                        <span class="badge badge-ppt-wait">Pendiente</span>
                                                    {% endif %}
                                                </div>
                                            {% endfor %}
                                        </div>
                                    </div>
                                </div>
                            {% endfor %}
                        </div>
                    </div>
                </div>
            {% endfor %}
        </div>
        {% endif %}
    </div>

<script>
const SALA_BLOQUEADA = {{ sala_restringida | tojson }};

function mostrarAlerta(texto, ok = true) {
    const el = document.getElementById('alertaAccion');
    el.className = 'alert mb-4';
    el.style.background = ok ? 'rgba(2,150,102,0.2)' : 'rgba(249,58,90,0.2)';
    el.style.border = ok ? '1px solid #48d6a8' : '1px solid #f7778c';
    el.style.color = ok ? '#6ee7b7' : '#fda4af';
    el.innerHTML = `<i class="bi ${ok ? 'bi-check-circle-fill' : 'bi-exclamation-triangle-fill'} me-2"></i>${texto}`;
}

function cambiarVista(vista) {
    const items = document.querySelectorAll('.valex-subnav .subnav-item');
    items.forEach(i => i.classList.remove('active'));
    if (vista === 'tabla') {
        items[0].classList.add('active');
        document.getElementById('vistaTabla').classList.remove('d-none');
        document.getElementById('vistaVisor').classList.add('d-none');
    } else {
        items[1].classList.add('active');
        document.getElementById('vistaTabla').classList.add('d-none');
        document.getElementById('vistaVisor').classList.remove('d-none');
    }
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

function aplicarFiltros() {
    const estado = document.getElementById('flt_estado').value;
    const ponenteFiltro = document.getElementById('flt_ponente_col').value.toLowerCase().trim();
    const inputsCol = Array.from(document.querySelectorAll('.flt-col'));
    const filas = document.querySelectorAll('#tablaPonencias tbody tr');
    let visibles = 0;

    filas.forEach(tr => {
        const salaFila = (tr.getAttribute('data-sala') || '').toLowerCase();
        if (SALA_BLOQUEADA && salaFila !== SALA_BLOQUEADA.toLowerCase()) {
            tr.style.display = 'none';
            return;
        }
        const estadoFila = tr.getAttribute('data-estado');
        const ponentesFila = tr.getAttribute('data-ponentes') || '';

        let ok = true;
        if (estado && estadoFila !== estado) ok = false;
        if (ponenteFiltro && !ponentesFila.includes(ponenteFiltro)) ok = false;

        if (ok) {
            for (let i = 0; i < inputsCol.length; i++) {
                const valFiltro = inputsCol[i].value.toLowerCase().trim();
                if (valFiltro) {
                    const idx = inputsCol[i].getAttribute('data-idx');
                    const valCelda = tr.getAttribute('data-c_' + idx) || '';
                    if (!valCelda.includes(valFiltro)) {
                        ok = false;
                        break;
                    }
                }
            }
        }

        tr.style.display = ok ? '' : 'none';
        if (ok) visibles++;
    });

    document.getElementById('kpiVisibles').textContent = visibles;
}

function limpiarFiltros() {
    document.getElementById('flt_estado').value = '';
    document.getElementById('flt_ponente_col').value = '';
    document.getElementById('buscadorPonente').value = '';
    document.querySelectorAll('.flt-col').forEach(i => i.value = '');
    document.getElementById('resultadosPonente').classList.add('d-none');
    aplicarFiltros();
}

function buscarPonente() {
    const q = document.getElementById('buscadorPonente').value.toLowerCase().trim();
    document.getElementById('flt_ponente_col').value = q;
    aplicarFiltros();
}

function irACarpetaPonente() {
    const q = document.getElementById('buscadorPonente').value.trim();
    if (!q) {
        mostrarAlerta('Escribe el nombre de un ponente o participante para abrir su carpeta.', false);
        return;
    }
    fetch('/api/buscar-ponente?q=' + encodeURIComponent(q))
        .then(r => r.json())
        .then(data => {
            const box = document.getElementById('resultadosPonente');
            if (!data.coincidencias || data.coincidencias.length === 0) {
                box.classList.remove('d-none');
                box.innerHTML = `<span class="text-warning"><i class="bi bi-search me-1"></i> No se encontraron carpetas para "${q}".</span>`;
                return;
            }
            if (data.coincidencias.length === 1) {
                abrirRuta(data.coincidencias[0].ruta, true);
                mostrarAlerta(`Abierta carpeta de ${data.coincidencias[0].ponente}: ${data.coincidencias[0].ruta}`);
                box.classList.add('d-none');
            } else {
                box.classList.remove('d-none');
                box.innerHTML = `<div class="small text-info fw-bold mb-2">Se encontraron ${data.coincidencias.length} carpetas para "${q}":</div>` +
                    data.coincidencias.map(c =>
                        `<span class="ponente-link" onclick='abrirRuta(${JSON.stringify(c.ruta)}, true)'><i class="bi bi-folder2-open me-1"></i>${c.ponente} — <small>${c.ruta}</small></span>`
                    ).join('');
            }
        });
}

if (SALA_BLOQUEADA) {
    aplicarFiltros();
}
</script>
</body>
</html>"""


@app.route('/')
def index():
    config, df, ruta_activa = cargar_datos_congreso()
    if config is None or df is None:
        return render_template_string(HTML_APP, config=None, puesto_activo=None, puesto_bloqueado=None, sala_restringida=None, es_admin=False)

    columnas_raw = list(df.columns)
    columnas_ordenadas = ordenar_columnas_solicitadas(columnas_raw)

    col_sala = buscar_columna(columnas_raw, ['sala'])
    col_dia = buscar_columna(columnas_raw, ['dia', 'fecha'])
    col_titulo = buscar_columna(columnas_raw, ['titulo sesion es', 'titulo sesion', 'titulo de sesion'])

    puesto_bloqueado = config.get("_puesto_bloqueado")
    puesto_activo = puesto_bloqueado or ""
    es_admin = str(puesto_activo).upper().startswith("ADM") or puesto_activo == ""
    sala_restringida = None if (not puesto_activo or es_admin) else puesto_activo

    filas = []
    con_archivo = 0
    visor_por_dia = {}

    for _, row in df.iterrows():
        ruta_sesion = construir_ruta_sesion(row, config, columnas_raw)
        ruta_abs = os.path.join(ruta_activa, ruta_sesion)
        participantes = obtener_participantes_fila(row, columnas_raw)

        part_carpetas = []
        for p in participantes:
            p_san = sanitizar_nombre_carpeta(p)
            part_carpetas.append({
                "nombre": p,
                "ruta": os.path.join(ruta_sesion, p_san)
            })

        num_archivos = contar_archivos_en_carpeta(ruta_abs)
        tiene_ppt = num_archivos > 0
        if tiene_ppt:
            con_archivo += 1

        sala_val = str(row[col_sala]).strip() if col_sala else ""
        dia_val = str(row[col_dia]).strip() if col_dia else "Sin Día"
        sesion_val = os.path.basename(ruta_sesion)
        titulo_val = str(row[col_titulo]).strip() if col_titulo else sesion_val

        filas.append({
            "datos": {c: str(row[c]) for c in columnas_ordenadas},
            "sala": sala_val,
            "participantes": participantes,
            "participantes_carpetas": part_carpetas,
            "ruta_relativa": ruta_sesion,
            "tiene_ppt": tiene_ppt,
            "num_archivos": num_archivos
        })

        if sala_restringida and sala_val.lower() != sala_restringida.lower():
            continue

        visor_por_dia.setdefault(dia_val, {}).setdefault(sesion_val, []).append({
            "titulo_corto": titulo_val[:65],
            "participantes": participantes,
            "tiene_ppt": tiene_ppt
        })

    return render_template_string(
        HTML_APP,
        config=config,
        columnas_ordenadas=columnas_ordenadas,
        filas=filas,
        con_archivo=con_archivo,
        puesto_activo=puesto_activo,
        puesto_bloqueado=puesto_bloqueado,
        sala_restringida=sala_restringida,
        es_admin=es_admin,
        visor_por_dia=visor_por_dia
    )


@app.route('/api/fijar-puesto', methods=['POST'])
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


@app.route('/api/crear-carpetas', methods=['POST'])
def api_crear_carpetas():
    config, df, ruta_activa = cargar_datos_congreso()
    if config is None:
        return jsonify({"ok": False, "mensaje": "Faltan archivos de configuración."})
    puesto = config.get("_puesto_bloqueado") or ""
    if puesto and not puesto.upper().startswith("ADM"):
        return jsonify({"ok": False, "mensaje": "Solo los puestos ADMIN pueden crear o modificar carpetas."})

    columnas_raw = list(df.columns)
    creadas = 0
    for _, row in df.iterrows():
        ruta_sesion = construir_ruta_sesion(row, config, columnas_raw)
        ruta_abs = os.path.join(ruta_activa, ruta_sesion)
        if not os.path.exists(ruta_abs):
            os.makedirs(ruta_abs, exist_ok=True)
            creadas += 1
        for p in obtener_participantes_fila(row, columnas_raw):
            p_abs = os.path.join(ruta_abs, sanitizar_nombre_carpeta(p))
            if not os.path.exists(p_abs):
                os.makedirs(p_abs, exist_ok=True)
                creadas += 1

    return jsonify({"ok": True, "mensaje": f"Carpetas de sesión (≤50 car.) y participantes listas ({creadas} nuevas)."})


@app.route('/api/abrir-ruta', methods=['POST'])
def api_abrir_ruta():
    config, df, ruta_activa = cargar_datos_congreso()
    data = request.json or {}
    ruta_rel = data.get("ruta", "PPT")
    crear = bool(data.get("crear", True))
    puesto = config.get("_puesto_bloqueado") or "" if config else ""
    es_admin = (not puesto) or puesto.upper().startswith("ADM")

    ruta_abs = os.path.join(ruta_activa, ruta_rel)
    if crear and es_admin:
        os.makedirs(ruta_abs, exist_ok=True)
        if df is not None:
            columnas_raw = list(df.columns)
            for _, row in df.iterrows():
                if construir_ruta_sesion(row, config, columnas_raw) == ruta_rel:
                    for p in obtener_participantes_fila(row, columnas_raw):
                        os.makedirs(os.path.join(ruta_abs, sanitizar_nombre_carpeta(p)), exist_ok=True)
    elif not os.path.exists(ruta_abs):
        return jsonify({"ok": False, "mensaje": "La carpeta aún no existe en disco."})

    if SISTEMA == 'Windows':
        os.startfile(ruta_abs)
    elif SISTEMA == 'Darwin':
        subprocess.Popen(["open", ruta_abs])
    else:
        subprocess.Popen(["xdg-open", ruta_abs])
    return jsonify({"ok": True})


@app.route('/api/subir-presentacion', methods=['POST'])
def api_subir_presentacion():
    config, _, ruta_activa = cargar_datos_congreso()
    puesto = config.get("_puesto_bloqueado") or "" if config else ""
    if puesto and not puesto.upper().startswith("ADM"):
        return jsonify({"ok": False, "mensaje": "Acceso denegado: las salas solo pueden ver presentaciones."})

    archivo = request.files.get('archivo')
    ruta_rel = request.form.get('ruta', 'PPT')
    if not archivo or not archivo.filename:
        return jsonify({"ok": False, "mensaje": "No se seleccionó ningún archivo."})

    carpeta_destino = os.path.join(ruta_activa, ruta_rel)
    os.makedirs(carpeta_destino, exist_ok=True)
    nombre_seguro = sanitizar_nombre_carpeta(os.path.splitext(archivo.filename)[0]) + os.path.splitext(archivo.filename)[1]
    archivo.save(os.path.join(carpeta_destino, nombre_seguro))
    return jsonify({"ok": True, "mensaje": f"Presentación '{nombre_seguro}' cargada en {ruta_rel}."})


@app.route('/api/buscar-ponente')
def api_buscar_ponente():
    q = (request.args.get("q") or "").strip().lower()
    config, df, _ = cargar_datos_congreso()
    if not q or config is None or df is None:
        return jsonify({"coincidencias": []})

    columnas_raw = list(df.columns)
    coincidencias = []
    vistas = set()

    for _, row in df.iterrows():
        ruta_sesion = construir_ruta_sesion(row, config, columnas_raw)
        for p in obtener_participantes_fila(row, columnas_raw):
            if q in p.lower():
                ruta_p = os.path.join(ruta_sesion, sanitizar_nombre_carpeta(p))
                clave = f"{p}||{ruta_p}"
                if clave not in vistas:
                    vistas.add(clave)
                    coincidencias.append({"ponente": p, "ruta": ruta_p})

    return jsonify({"coincidencias": coincidencias})


def limpiar_bloqueo_al_salir(ruta_archivo_bloqueo):
    try:
        if ruta_archivo_bloqueo and os.path.exists(ruta_archivo_bloqueo):
            os.remove(ruta_archivo_bloqueo)
    except Exception:
        pass


if __name__ == '__main__':
    cfg, _, _ = cargar_datos_congreso()
    ruta_bloq = cfg.get("_ruta_puesto_sesion") if cfg else None
    limpiar_bloqueo_al_salir(ruta_bloq)
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5051")).start()
    try:
        app.run(host='127.0.0.1', port=5051)
    finally:
        limpiar_bloqueo_al_salir(ruta_bloq)
