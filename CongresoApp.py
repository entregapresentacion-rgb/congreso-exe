import os
import re
import sys
import json
import time
import shutil
import socket
import platform
import threading
import subprocess
import webbrowser
import unicodedata
import pandas as pd
from datetime import datetime, date
from werkzeug.utils import secure_filename
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
    if texto.lower() in ("nan", "none", "null", "nat"):
        return ""
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


def detectar_todas_columnas_participantes(df_cols):
    palabras_clave = [
        "ponente", "ponentes", "participante", "participantes",
        "autor", "autores", "orador", "oradores", "speaker", "speakers",
        "moderador", "moderadores", "chair", "coordinador", "coordinadores",
        "conferenciante", "conferenciantes", "invitado", "invitados",
        "presentador", "presentadores", "disertante", "disertantes"
    ]
    encontradas = []
    for col in df_cols:
        col_norm = normalizar_clave(col)
        if any(pk in col_norm for pk in palabras_clave):
            encontradas.append(col)
    return encontradas


def dividir_nombres_personas(texto_raw):
    if not texto_raw or (isinstance(texto_raw, float) and pd.isna(texto_raw)):
        return []
    texto = str(texto_raw).strip()
    if not texto or texto.lower() in ("nan", "none", "null", "-", "n/a", "sin especificar"):
        return []

    # Separar por punto y coma, saltos de línea, barras, pipes, " / ", " y " o comas cuando separan personas
    # Primero dividimos por separadores fuertes: ; \n | /
    bloques = re.split(r'[;\n|/]+', texto)
    nombres = []
    for bloque in bloques:
        bloque = bloque.strip()
        if not bloque:
            continue
        # Si hay comas, comprobar si son listas de nombres ("Juan Pérez, María López") vs "Pérez, Juan"
        partes_coma = [p.strip() for p in bloque.split(',') if p.strip()]
        if len(partes_coma) > 1 and all(len(p.split()) >= 2 for p in partes_coma):
            sub_items = partes_coma
        else:
            sub_items = [bloque]

        for sub in sub_items:
            # Separar también por " y " o " & " si une dos nombres completos
            sub_y = re.split(r'\s+(?:y|&)\s+', sub)
            if len(sub_y) > 1 and all(len(x.split()) >= 2 for x in sub_y):
                candidatos = sub_y
            else:
                candidatos = [sub]

            for cand in candidatos:
                # Limpiar posibles prefijos de rol entre paréntesis o dos puntos (ej. "Moderador: Juan Pérez")
                cand_limpio = re.sub(r'^(?:moderador(?:a)?|ponente|autor(?:a)?|coord(?:inador)?|chair)\s*:\s*', '', cand, flags=re.IGNORECASE).strip()
                if cand_limpio and len(cand_limpio) >= 2:
                    nombres.append(cand_limpio)
    return nombres


def extraer_participantes_fila(row, df_cols):
    participantes = []
    vistas = set()

    def agregar(nombre_raw):
        for persona in dividir_nombres_personas(nombre_raw):
            limpio = sanitizar_nombre_carpeta(persona, max_len=80)
            clave = normalizar_clave(limpio)
            if limpio and limpio != "Sin_Especificar" and clave not in vistas:
                vistas.add(clave)
                participantes.append(limpio)

    # 1. Revisar si hay columnas explícitas de Nombre + Apellidos
    col_nombre = detectar_columna(df_cols, ["nombre", "first name", "nombre participante", "nombre ponente"])
    col_apellidos = detectar_columna(df_cols, ["apellidos", "apellido", "last name", "surname", "apellidos participante", "apellidos ponente"])
    if col_nombre and col_apellidos:
        nom = formatear_valor_celda(row.get(col_nombre, ""))
        ape = formatear_valor_celda(row.get(col_apellidos, ""))
        completo = f"{nom} {ape}".strip()
        if completo:
            agregar(completo)
    elif col_nombre and not detectar_todas_columnas_participantes(df_cols):
        nom = formatear_valor_celda(row.get(col_nombre, ""))
        if nom:
            agregar(nom)

    # 2. Revisar TODAS las columnas relacionadas con participantes/ponentes/autores/moderadores (ej. Ponente 1, Ponente 2, Participantes, Moderador...)
    cols_part = detectar_todas_columnas_participantes(df_cols)
    for col_p in cols_part:
        agregar(row.get(col_p, ""))

    return participantes


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


def calcular_ruta_sesion_fila(row, df_cols, cfg):
    orden_carpetas = cfg.get("orden_carpetas", [])
    sep_n3 = cfg.get("separador_nivel_3", " - ")
    max_n3 = int(cfg.get("max_chars_nivel_3", 50))

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
    return ruta_actual


def leer_programa_estructurado():
    if not os.path.exists(RUTA_EXCEL):
        return [], "No se encontró programa.xlsx"

    cfg = cargar_config()
    try:
        df = pd.read_excel(RUTA_EXCEL, engine="openpyxl")
    except Exception as e:
        return [], f"Error al leer programa.xlsx: {str(e)}"

    df_cols = list(df.columns)
    col_sala = detectar_columna(df_cols, ["sala"])
    col_dia = detectar_columna(df_cols, ["dia", "fecha"])
    col_ini = detectar_columna(df_cols, ["hora inicio", "inicio"])
    col_fin = detectar_columna(df_cols, ["hora fin", "fin"])
    col_tit = detectar_columna(df_cols, ["titulo sesion", "titulo de sesion", "sesion", "titulo"])

    # Rellenar hacia adelante celdas combinadas en Excel (merged cells de Sala, Día, Hora, Sesión)
    df_ffill = df.copy()
    cols_sesion_clave = [c for c in [col_sala, col_dia, col_ini, col_fin, col_tit] if c]
    if cols_sesion_clave:
        df_ffill[cols_sesion_clave] = df_ffill[cols_sesion_clave].ffill()

    sesiones_map = {}
    orden_rutas = []

    for idx, row in df_ffill.iterrows():
        if df.loc[idx].isna().all():
            continue

        ruta_sesion = calcular_ruta_sesion_fila(row, df_cols, cfg)
        participantes_fila = extraer_participantes_fila(df.loc[idx], df_cols)

        if ruta_sesion not in sesiones_map:
            orden_rutas.append(ruta_sesion)
            sesiones_map[ruta_sesion] = {
                "ruta_sesion": ruta_sesion,
                "nombre_carpeta_sesion": os.path.basename(ruta_sesion),
                "sala": formatear_valor_celda(row.get(col_sala, "")) if col_sala else "General",
                "dia": formatear_valor_celda(row.get(col_dia, "")) if col_dia else "",
                "hora_inicio": formatear_valor_celda(row.get(col_ini, ""), es_hora=True) if col_ini else "",
                "hora_fin": formatear_valor_celda(row.get(col_fin, ""), es_hora=True) if col_fin else "",
                "titulo": formatear_valor_celda(row.get(col_tit, "")) if col_tit else os.path.basename(ruta_sesion),
                "participantes": [],
                "_vistas": set()
            }

        sesion_obj = sesiones_map[ruta_sesion]
        for p in participantes_fila:
            clave_p = normalizar_clave(p)
            if clave_p not in sesion_obj["_vistas"]:
                sesion_obj["_vistas"].add(clave_p)
                ruta_p = os.path.join(ruta_sesion, p)
                archivos = []
                if os.path.exists(ruta_p):
                    archivos = [
                        f for f in sorted(os.listdir(ruta_p))
                        if not f.startswith('.') and os.path.isfile(os.path.join(ruta_p, f))
                    ]
                sesion_obj["participantes"].append({
                    "nombre": p,
                    "ruta": ruta_p,
                    "archivos": archivos,
                    "cargado": len(archivos) > 0
                })

    resultado = []
    for r in orden_rutas:
        s = sesiones_map[r]
        if not s["participantes"]:
            ruta_gen = os.path.join(r, "General_Sesion")
            archivos = [
                f for f in sorted(os.listdir(ruta_gen))
                if not f.startswith('.') and os.path.isfile(os.path.join(ruta_gen, f))
            ] if os.path.exists(ruta_gen) else []
            s["participantes"].append({
                "nombre": "General_Sesion",
                "ruta": ruta_gen,
                "archivos": archivos,
                "cargado": len(archivos) > 0
            })
        del s["_vistas"]
        resultado.append(s)

    return resultado, None


def generar_estructura_desde_excel():
    sesiones, err = leer_programa_estructurado()
    if err:
        return False, err

    os.makedirs(RUTA_PPT, exist_ok=True)
    total_sesiones = 0
    total_participantes = 0

    for ses in sesiones:
        os.makedirs(ses["ruta_sesion"], exist_ok=True)
        total_sesiones += 1
        for part in ses["participantes"]:
            os.makedirs(part["ruta"], exist_ok=True)
            total_participantes += 1

    return True, f"Estructura creada en {RUTA_PPT}: {total_sesiones} sesiones y {total_participantes} carpetas individuales de participantes."


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
    <title>CongresosUSB | Gestión de Presentaciones y Red</title>
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
            <span>Congresos<span style="color:#00b9ff;">USB</span> · Panel Completo</span>
        </a>
        <div class="d-flex gap-2 align-items-center">
            <span class="badge bg-dark border border-secondary px-3 py-2">PC: {{ info.hostname }} ({{ info.ip }})</span>
            <span class="badge bg-dark border border-info text-info px-3 py-2">Grupo: {{ info.workgroup }}</span>
        </div>
    </header>

    <nav class="valex-subnav">
        <button class="subnav-btn active" id="tabBtnCargar" onclick="cambiarTab('cargar')">
            <i class="bi bi-file-earmark-slides-fill"></i> Carga de Presentaciones y Participantes
        </button>
        <button class="subnav-btn" id="tabBtnRed" onclick="cambiarTab('red')">
            <i class="bi bi-hdd-network-fill"></i> Equipos en Red y Carpetas Compartidas
        </button>
    </nav>

    <div class="container-fluid px-4 pt-4" style="max-width: 1320px;">
        <div id="alertaEstado" class="alert d-none mb-4"></div>

        <!-- PESTAÑA 1: CARGA DE PRESENTACIONES Y GESTIÓN POR PARTICIPANTE -->
        <div id="panelCargar">
            <div class="valex-card mb-4">
                <div class="valex-card-header flex-wrap gap-2">
                    <div>
                        <h5 class="valex-card-title">Recepción y Carga de Presentaciones</h5>
                        <small class="text-secondary">Cada sesión contiene una carpeta individual por participante dentro de <code>{{ info.ruta_ppt }}</code></small>
                    </div>
                    <div class="d-flex flex-wrap gap-2">
                        <button class="btn btn-valex-primary btn-sm px-3 py-2" onclick="crearEstructura()">
                            <i class="bi bi-folder-plus me-1"></i> Generar / Actualizar Todas las Carpetas
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
                    <div class="text-center text-secondary py-4">Cargando sesiones y participantes...</div>
                </div>
            </div>
        </div>

        <!-- PESTAÑA 2: EQUIPOS EN RED Y CONFIGURACIÓN LOCAL -->
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
                                Comparte la carpeta <code>{{ info.ruta_ppt }}</code> en la red local para que los ordenadores de sala accedan a las presentaciones.
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

async function crearEstructura() {
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

    const salas = [...new Set(sesionesData.map(s => s.sala).filter(Boolean))].sort();
    const dias = [...new Set(sesionesData.map(s => s.dia).filter(Boolean))].sort();

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
        if (fSala && s.sala !== fSala) return false;
        if (fDia && s.dia !== fDia) return false;
        if (q) {
            const textoSesion = `${s.sala} ${s.dia} ${s.titulo} ${s.nombre_carpeta_sesion} ${s.participantes.map(p => p.nombre + ' ' + p.archivos.join(' ')).join(' ')}`.toLowerCase();
            if (!textoSesion.includes(q)) return false;
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
                        <span class="badge bg-info text-dark">${s.sala || 'Sala'}</span>
                        ${s.dia ? `<span class="badge bg-secondary">${s.dia}</span>` : ''}
                        ${(s.hora_inicio || s.hora_fin) ? `<span class="badge bg-dark border border-secondary">${s.hora_inicio} - ${s.hora_fin}</span>` : ''}
                    </div>
                    <h6 class="fw-bold mb-1 text-white">${s.titulo}</h6>
                    <small class="text-secondary font-monospace"><i class="bi bi-folder-fill text-warning me-1"></i>${s.nombre_carpeta_sesion}</small>
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
                            </div>
                            <div class="mt-1 d-flex flex-wrap gap-1">
                                ${p.archivos.length > 0
                                    ? p.archivos.map(a => `
                                        <span class="file-pill">
                                            <i class="bi bi-file-earmark-ppt-fill"></i> ${a}
                                            <button type="button" class="btn btn-link btn-sm p-0 text-info" title="Abrir presentación" onclick='ejecutarAccion("/api/abrir_ruta", {ruta: ${JSON.stringify(p.ruta + "/" + a)}})'>
                                                <i class="bi bi-box-arrow-up-right"></i>
                                            </button>
                                        </span>`).join('')
                                    : '<small class="text-secondary">Sin archivos aún en esta subcarpeta.</small>'}
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
    sesiones, err = leer_programa_estructurado()
    if err:
        return jsonify({"ok": False, "mensaje": err, "sesiones": []})
    return jsonify({"ok": True, "sesiones": sesiones})


@app.route('/api/crear_carpetas', methods=['POST'])
def api_crear_carpetas():
    ok, msg = generar_estructura_desde_excel()
    return jsonify({"ok": ok, "mensaje": msg})


@app.route('/api/subir_presentacion', methods=['POST'])
def api_subir_presentacion():
    ruta_destino = request.form.get('ruta_destino', '').strip()
    archivos = request.files.getlist('archivos')
    if not ruta_destino or not archivos:
        return jsonify({"ok": False, "mensaje": "Falta la carpeta de destino o el archivo."})

    # Asegurar que la ruta esté dentro de RUTA_PPT
    ruta_abs_ppt = os.path.abspath(RUTA_PPT)
    ruta_abs_dest = os.path.abspath(ruta_destino)
    if not ruta_abs_dest.startswith(ruta_abs_ppt):
        return jsonify({"ok": False, "mensaje": "Ruta de destino no válida."})

    os.makedirs(ruta_abs_dest, exist_ok=True)
    guardados = []
    for f in archivos:
        if f and f.filename:
            nombre_limpio = sanitizar_nombre_carpeta(os.path.splitext(f.filename)[0], max_len=100) + os.path.splitext(f.filename)[1]
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
