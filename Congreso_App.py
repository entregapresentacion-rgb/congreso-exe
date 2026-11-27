import os
import re
import sys
import json
import shutil
import socket
import traceback
import platform
import subprocess
import datetime as dt
import webview
import pandas as pd

SISTEMA = platform.system()

INVALID_CHARS = set('<>:"/\\|?*') | {chr(i) for i in range(32)}
RESERVED = (
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)
MAX_NAME_BYTES = 240

ORDEN_COLUMNAS_PRIORITARIAS = [
    ("ESTADO PPT", ["estado ppt", "estado_ppt"]),
    ("CARPETA PONENCIA", ["carpeta ponencia", "carpeta_ponencia"]),
    ("SALA", ["sala"]),
    ("DÍA", ["dia", "fecha"]),
    ("HORA INICIO", ["hora inicio", "hora_inicio", "inicio"]),
    ("HORA FIN", ["hora fin", "hora_fin", "fin"]),
    ("TÍTULO SESIÓN ES", ["titulo sesion es", "titulo sesion", "titulo de sesion", "sesion"]),
]

NOMBRE_BACKUP_LOCAL = "BACKUP_LOCAL"


def _cut_bytes(s, n=MAX_NAME_BYTES):
    b = s.encode("utf-8")
    if len(b) <= n:
        return s
    return b[:n].decode("utf-8", errors="ignore")


def sanitize_name(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return "_vacio"
    val = str(s).replace("\xa0", " ")
    clean = "".join(" " if ch in INVALID_CHARS else ch for ch in val)
    clean = re.sub(r"\s+", " ", clean).strip()
    clean = clean.rstrip(". ")
    if clean.upper() in RESERVED:
        clean += "_"
    clean = _cut_bytes(clean, MAX_NAME_BYTES).rstrip(". ").strip()
    return clean if clean else "_vacio"


def truncate50(s):
    if s is None:
        return ""
    s = str(s)
    return s if len(s) <= 50 else s[:50]


def formatear_dia_valor(valor, day_choice=0):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "_sin_fecha"
    if isinstance(valor, (dt.datetime, dt.date)):
        return f"DIA {valor.day:02d}"
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    texto = str(valor).replace("\xa0", " ").strip().strip("\"' ").strip()
    if not texto:
        return "_sin_fecha"
    if re.fullmatch(r"[0-9]{8}", texto):
        try:
            parsed = dt.datetime.strptime(texto, "%Y%m%d")
            return f"DIA {parsed.day:02d}"
        except ValueError:
            pass
    if re.fullmatch(r"[0-9]{1,2}", texto):
        d = int(texto)
        return f"DIA {d:02d}" if 1 <= d <= 31 else "DIA " + sanitize_name(texto)
    partes = [p for p in re.split(r"[/\-.\s]+", texto) if p != ""]
    if len(partes) == 3 and all(re.fullmatch(r"[0-9]{1,4}", p) for p in partes):
        idx = day_choice if 0 <= day_choice < 3 else 0
        try:
            d = int(partes[idx])
            if 1 <= d <= 31:
                return f"DIA {d:02d}"
        except ValueError:
            pass
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y%m%d", "%d %B %Y", "%B %d, %Y", "%d %b %Y"):
        try:
            parsed = dt.datetime.strptime(texto, fmt)
            return f"DIA {parsed.day:02d}"
        except ValueError:
            continue
    return "DIA " + sanitize_name(texto)


def obtener_ruta_base():
    if getattr(sys, "frozen", False):
        ruta = os.path.dirname(sys.executable)
        if SISTEMA == "Darwin" and ruta.endswith("Contents/MacOS"):
            ruta = os.path.abspath(os.path.join(ruta, "..", "..", ".."))
        return ruta
    return os.path.abspath(os.path.dirname(__file__))


def obtener_ip_local():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"


def obtener_ruta_backup_local():
    if SISTEMA == "Windows" and os.path.exists("C:\\"):
        raiz_c = "C:\\"
    else:
        raiz_c = os.path.expanduser("~")
    ruta_bk = os.path.join(raiz_c, NOMBRE_BACKUP_LOCAL)
    os.makedirs(ruta_bk, exist_ok=True)
    return ruta_bk


def crear_acceso_directo_escritorio():
    ruta_bk = obtener_ruta_backup_local()
    escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.exists(escritorio):
        escritorio = os.path.join(os.path.expanduser("~"), "Escritorio")
    if not os.path.exists(escritorio):
        return None

    nombre_base = "BACKUP_LOCAL"
    try:
        if SISTEMA == "Windows":
            url_path = os.path.join(escritorio, f"{nombre_base}.url")
            if not os.path.exists(url_path):
                uri = "file:///" + os.path.abspath(ruta_bk).replace("\\", "/")
                with open(url_path, "w", encoding="utf-8") as f:
                    f.write(f"[InternetShortcut]\nURL={uri}\n")
            return url_path
        elif SISTEMA == "Darwin":
            webloc_path = os.path.join(escritorio, f"{nombre_base}.webloc")
            if not os.path.exists(webloc_path):
                uri = "file://" + os.path.abspath(ruta_bk)
                plist = (
                    '<?xml version="1.0" encoding="UTF-8"?>\n'
                    '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" '
                    '"http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                    '<plist version="1.0"><dict><key>URL</key><string>'
                    f"{uri}</string></dict></plist>\n"
                )
                with open(webloc_path, "w", encoding="utf-8") as f:
                    f.write(plist)
            return webloc_path
        else:
            desktop_path = os.path.join(escritorio, f"{nombre_base}.desktop")
            if not os.path.exists(desktop_path):
                with open(desktop_path, "w", encoding="utf-8") as f:
                    f.write(
                        "[Desktop Entry]\n"
                        "Type=Application\n"
                        "Name=BACKUP_LOCAL\n"
                        f'Exec=xdg-open "{ruta_bk}"\n'
                        "Icon=folder\n"
                        "Terminal=false\n"
                    )
            return desktop_path
    except Exception:
        return None


def copiar_a_backup_local_sin_borrar(ruta_rel_carpeta, nombre_archivo, archivo_origen_abs):
    ruta_bk_raiz = obtener_ruta_backup_local()
    rel_limpia = os.path.normpath(ruta_rel_carpeta).lstrip("/\\")
    carpeta_bk_destino = os.path.join(ruta_bk_raiz, rel_limpia)
    os.makedirs(carpeta_bk_destino, exist_ok=True)

    base_name, ext = os.path.splitext(nombre_archivo)
    candidato = os.path.join(carpeta_bk_destino, nombre_archivo)
    contador = 2
    while os.path.exists(candidato):
        candidato = os.path.join(carpeta_bk_destino, f"{base_name}_copia{contador}{ext}")
        contador += 1

    shutil.copy2(archivo_origen_abs, candidato)
    return candidato


def sincronizar_estructura_backup_local(ruta_activa):
    ruta_bk_raiz = obtener_ruta_backup_local()
    ruta_ppt_raiz = os.path.join(ruta_activa, "PPT")
    os.makedirs(ruta_ppt_raiz, exist_ok=True)
    os.makedirs(os.path.join(ruta_bk_raiz, "PPT"), exist_ok=True)

    copiados = 0
    for dirpath, _, filenames in os.walk(ruta_ppt_raiz):
        rel_dir = os.path.relpath(dirpath, ruta_activa)
        dest_dir = os.path.join(ruta_bk_raiz, rel_dir)
        os.makedirs(dest_dir, exist_ok=True)
        for fname in filenames:
            if fname.startswith("."):
                continue
            src_file = os.path.join(dirpath, fname)
            dst_file = os.path.join(dest_dir, fname)
            if not os.path.exists(dst_file):
                shutil.copy2(src_file, dst_file)
                copiados += 1
    return copiados


def norm_col(nombre):
    s = str(nombre).replace("\xa0", " ").lower().strip()
    for a, b in [("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ñ", "n")]:
        s = s.replace(a, b)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def cargar_datos_congreso():
    ruta_activa = obtener_ruta_base()
    ruta_json = os.path.join(ruta_activa, "config_congreso.json")
    ruta_excel = os.path.join(ruta_activa, "programa.xlsx")

    if not os.path.exists(ruta_json) or not os.path.exists(ruta_excel):
        return None, None, ruta_activa

    with open(ruta_json, "r", encoding="utf-8") as f:
        config = json.load(f)

    df = pd.read_excel(ruta_excel, engine="openpyxl").fillna("")

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
        ncand = norm_col(cand)
        for c in columnas:
            if norm_col(c) == ncand:
                return c
    for cand in candidatos:
        ncand = norm_col(cand)
        for c in columnas:
            if ncand and ncand in norm_col(c):
                return c
    return None


def ordenar_columnas_tabla(columnas_raw, config=None):
    cols = [str(c) for c in columnas_raw]
    ordenadas = []
    usadas = set()

    for etiqueta_defecto, aliases in ORDEN_COLUMNAS_PRIORITARIAS:
        encontrada = buscar_columna([c for c in cols if c not in usadas], aliases)
        if encontrada:
            ordenadas.append(encontrada)
            usadas.add(encontrada)
        else:
            ordenadas.append(etiqueta_defecto)

    if config and isinstance(config.get("orden_columnas_tabla"), list):
        for c_cfg in config["orden_columnas_tabla"]:
            if c_cfg in cols and c_cfg not in usadas:
                ordenadas.append(c_cfg)
                usadas.add(c_cfg)

    for c in cols:
        if c not in usadas:
            ordenadas.append(c)
            usadas.add(c)

    return ordenadas


def resolver_columnas_nivel_3(config, columnas_raw):
    alias_defecto = [
        ["hora inicio", "hora inicio sesion", "inicio"],
        ["hora fin", "hora final", "fin"],
        ["titulo sesion es", "titulo sesion", "titulo de sesion", "sesion", "titulo"],
        ["sala"],
        ["dia", "fecha"],
    ]

    n3_cfg = config.get("nivel_3_columnas") if config else None
    if not n3_cfg and config:
        for item in config.get("orden_carpetas", []):
            if isinstance(item, list) and item:
                first_norm = norm_col(item[0])
                if not any(k in first_norm for k in ("participante", "ponente", "autor", "speaker")):
                    n3_cfg = item
                    break

    resueltas = []
    usadas = set()

    if n3_cfg:
        for etiqueta_o_col in n3_cfg:
            if not etiqueta_o_col:
                continue
            if etiqueta_o_col in columnas_raw and etiqueta_o_col not in usadas:
                resueltas.append(etiqueta_o_col)
                usadas.add(etiqueta_o_col)
                continue
            encontrada = buscar_columna([c for c in columnas_raw if c not in usadas], [etiqueta_o_col])
            if encontrada:
                resueltas.append(encontrada)
                usadas.add(encontrada)

    if not resueltas:
        for grupo in alias_defecto:
            encontrada = buscar_columna([c for c in columnas_raw if c not in usadas], grupo)
            if encontrada:
                resueltas.append(encontrada)
                usadas.add(encontrada)

    return resueltas


def _formatear_nivel_3(row, cols, sep, max_chars):
    partes_limpias = []
    for c in cols:
        if c and c in row:
            raw_val = sanitize_name(row[c])
            partes_limpias.append(truncate50(raw_val)[:max_chars])
    combinado = sep.join(partes_limpias)
    return sanitize_name(combinado)


def construir_ruta_sesion(row, config, columnas_raw):
    partes = ["PPT"]
    sep_n3 = config.get("separador_nivel_3", " - ")
    max_n3 = int(config.get("max_chars_nivel_3", 50))
    formato_dia_nn = bool(config.get("formato_dia_nn", True))

    col_sala = config.get("nivel_1") or buscar_columna(columnas_raw, ["sala"])
    col_dia = config.get("nivel_2") or buscar_columna(columnas_raw, ["dia", "fecha"])
    n3_cols = resolver_columnas_nivel_3(config, columnas_raw)

    orden_carpetas = config.get("orden_carpetas", [])
    if orden_carpetas:
        for item in orden_carpetas:
            if isinstance(item, list):
                first_norm = norm_col(item[0]) if item else ""
                if any(k in first_norm for k in ("participante", "ponente", "autor", "speaker")):
                    continue
                partes.append(_formatear_nivel_3(row, n3_cols, sep_n3, max_n3))
            elif isinstance(item, str) and item in row:
                if col_dia and item == col_dia and formato_dia_nn:
                    partes.append(sanitize_name(formatear_dia_valor(row[item])))
                else:
                    partes.append(sanitize_name(row[item]))
        if n3_cols and not any(
            isinstance(x, list) and not any(k in norm_col(x[0] if x else "") for k in ("participante", "ponente", "autor", "speaker"))
            for x in orden_carpetas
        ):
            partes.append(_formatear_nivel_3(row, n3_cols, sep_n3, max_n3))
    else:
        if col_sala and col_sala in row:
            partes.append(sanitize_name(row[col_sala]))
        if col_dia and col_dia in row:
            val_dia = formatear_dia_valor(row[col_dia]) if formato_dia_nn else row[col_dia]
            partes.append(sanitize_name(val_dia))
        if n3_cols:
            partes.append(_formatear_nivel_3(row, n3_cols, sep_n3, max_n3))

    return os.path.join(*partes)


def _separar_nombres_celda(raw_val):
    texto = str(raw_val or "").replace("\xa0", " ").strip()
    if not texto:
        return []
    candidatos = re.split(r"[;|\n\r]+", texto)
    resultado = []
    for item in candidatos:
        parte = item.strip(" \t\r\n-–—")
        if parte:
            resultado.append(parte)
    return resultado


def obtener_participantes_fila(row, columnas, config=None):
    columnas_objetivo = []
    if config and isinstance(config.get("nivel_4_columnas"), list) and config["nivel_4_columnas"]:
        for col_cfg in config["nivel_4_columnas"]:
            if col_cfg in columnas:
                columnas_objetivo.append(col_cfg)
            else:
                encontrada = buscar_columna(columnas, [col_cfg])
                if encontrada:
                    columnas_objetivo.append(encontrada)

    if not columnas_objetivo:
        for c in columnas:
            nc = norm_col(c)
            if re.fullmatch(r"participante\s*\d*", nc) or re.fullmatch(r"ponente\s*\d*", nc):
                columnas_objetivo.append(c)

    participantes = []
    claves_vistas = set()

    for c in columnas_objetivo:
        val_celda = row[c] if c in row else ""
        for nombre in _separar_nombres_celda(val_celda):
            nombre_san = sanitize_name(nombre)
            if nombre_san == "_vacio":
                continue
            clave = re.sub(r"\s+", " ", nombre_san.lower()).strip()
            if clave not in claves_vistas:
                claves_vistas.add(clave)
                participantes.append(nombre.replace("\xa0", " ").strip())

    return participantes


def contar_archivos_en_carpeta(ruta_abs):
    if not os.path.exists(ruta_abs):
        return 0
    total = 0
    for _, _, files in os.walk(ruta_abs):
        for f in files:
            if not f.startswith("."):
                total += 1
    return total


def ordenar_columnas_por_nivel_3(columnas, n3_cols, config=None):
    base_ordenada = ordenar_columnas_tabla(columnas, config)
    ordenadas = []
    usadas = set()

    for c in n3_cols:
        if c in base_ordenada and c not in usadas:
            ordenadas.append(c)
            usadas.add(c)

    for c in base_ordenada:
        if c not in usadas:
            ordenadas.append(c)
            usadas.add(c)

    return ordenadas


def _construir_info_red(config, ruta_activa):
    ip_local = obtener_ip_local()
    hostname = socket.gethostname()
    ruta_bk = obtener_ruta_backup_local()
    ruta_ppt = os.path.join(ruta_activa, "PPT")
    acceso_esc = crear_acceso_directo_escritorio()
    archivos_bk = contar_archivos_en_carpeta(ruta_bk)
    puesto_bloq = (config or {}).get("_puesto_bloqueado") or ""
    equipos = (config or {}).get("equipos", [])

    equipos_detalle = []
    for eq in equipos:
        es_adm = str(eq).upper().startswith("ADM")
        equipos_detalle.append({
            "nombre": eq,
            "es_admin": es_adm,
            "bloqueado_aqui": bool(puesto_bloq and str(puesto_bloq).upper() == str(eq).upper()),
        })

    return {
        "ok": True,
        "ip_local": ip_local,
        "hostname": hostname,
        "workgroup": (config or {}).get("workgroup", "N/A"),
        "puesto_bloqueado": puesto_bloq,
        "ruta_ppt": ruta_ppt,
        "ruta_backup": ruta_bk,
        "acceso_escritorio": acceso_esc,
        "archivos_backup": archivos_bk,
        "equipos_detalle": equipos_detalle,
    }


def _render_simple(template_str, **ctx):
    out = template_str
    for k, v in ctx.items():
        token = "{{ " + k + " }}"
        if token in out:
            out = out.replace(token, "" if v is None else str(v))
    return out


HTML_ERROR = """<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Error en Congreso_App</title></head>
<body style="background:#111726;color:#fff;font-family:sans-serif;padding:30px;">
  <h2 style="color:#f87171;">Se produjo un error al cargar Congreso_App</h2>
  <p>Ruta activa detectada: <code>{{ ruta_activa }}</code></p>
  <pre style="background:#1e293b;padding:16px;border-radius:8px;overflow:auto;color:#fca5a5;">{{ detalle_error }}</pre>
</body>
</html>"""

HTML_RED = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CongresoApp | Red, Equipos y Backup</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <style>
        body { background-color: #111726; color: #fff; font-family: 'Roboto', sans-serif; padding-bottom: 40px; }
        .valex-topbar { background-color: #1f2940; border-bottom: 1px solid rgba(255,255,255,0.08); padding: 12px 24px; }
        .valex-card { background-color: #1f2940; border: 1px solid rgba(255,255,255,0.08); border-radius: 8px; }
        .valex-card-header { padding: 16px 22px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .table-dark-custom { --bs-table-bg: #151d30; --bs-table-color: #e2e8f0; --bs-table-border-color: rgba(255,255,255,0.08); }
    </style>
</head>
<body>
    <header class="valex-topbar d-flex justify-content-between align-items-center mb-4">
        <div class="d-flex align-items-center gap-3">
            <button type="button" class="btn btn-sm btn-outline-info" onclick="volverAgenda()"><i class="bi bi-arrow-left me-1"></i> Volver a Agenda</button>
            <h5 class="m-0 fw-bold">Panel de Red, Equipos y Backup Local</h5>
        </div>
        <span class="badge bg-dark border border-secondary px-3 py-2" id="badgeIpTop">
            <i class="bi bi-pc-display text-info me-1"></i> Cargando red...
        </span>
    </header>

    <div class="container-fluid px-4" style="max-width: 1320px;">
        <div class="row g-3 mb-4">
            <div class="col-md-4">
                <div class="valex-card p-3 h-100">
                    <div class="small text-info fw-bold text-uppercase">Red Local</div>
                    <div class="fs-5 fw-bold mt-1" id="redIp">-</div>
                    <div class="small text-secondary mt-1" id="redMeta">-</div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="valex-card p-3 h-100">
                    <div class="small text-success fw-bold text-uppercase">Carpeta Backup Local</div>
                    <div class="fs-6 fw-bold mt-1 text-break"><code id="redBackup">-</code></div>
                    <div class="small text-secondary mt-1" id="redBackupCount">-</div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="valex-card p-3 h-100">
                    <div class="small text-warning fw-bold text-uppercase">Acceso Directo Escritorio</div>
                    <div class="small fw-bold mt-1 text-break"><code id="redEscritorio">-</code></div>
                    <div class="small text-secondary mt-1" id="redPpt">-</div>
                </div>
            </div>
        </div>

        <div class="valex-card">
            <div class="valex-card-header d-flex justify-content-between align-items-center">
                <div>
                    <h6 class="m-0 fw-bold text-uppercase">Estado de Puestos y Equipos Configurados</h6>
                    <small class="text-secondary">Actualización automática cada 5 segundos</small>
                </div>
                <button class="btn btn-sm btn-outline-info" onclick="cargarEstadoRed()"><i class="bi bi-arrow-clockwise me-1"></i> Actualizar</button>
            </div>
            <div class="p-3">
                <table class="table table-dark table-hover table-dark-custom align-middle mb-0">
                    <thead>
                        <tr>
                            <th>Puesto / Equipo</th>
                            <th>Rol</th>
                            <th>Estado de Sesión</th>
                            <th>IP Local</th>
                        </tr>
                    </thead>
                    <tbody id="tablaEquiposRed"></tbody>
                </table>
            </div>
        </div>
    </div>

<script>
function volverAgenda() {
    window.pywebview.api.ir_a_agenda();
}

function cargarEstadoRed() {
    window.pywebview.api.obtener_estado_red().then(d => {
        if (!d || !d.ok) return;
        document.getElementById('badgeIpTop').innerHTML = `<i class="bi bi-pc-display text-info me-1"></i> IP Local: ${d.ip_local} (${d.hostname})`;
        document.getElementById('redIp').textContent = `${d.ip_local} (${d.hostname})`;
        document.getElementById('redMeta').textContent = `Grupo: ${d.workgroup}`;
        document.getElementById('redBackup').textContent = d.ruta_backup;
        document.getElementById('redBackupCount').textContent = `Archivos en backup: ${d.archivos_backup} (nunca borra nada)`;
        document.getElementById('redEscritorio').textContent = d.acceso_escritorio || 'No disponible';
        document.getElementById('redPpt').textContent = `Carpeta PPT: ${d.ruta_ppt}`;

        const tbody = document.getElementById('tablaEquiposRed');
        tbody.innerHTML = d.equipos_detalle.map(eq => `
            <tr>
                <td class="fw-bold">${eq.nombre}</td>
                <td>${eq.es_admin
                    ? '<span class="badge bg-primary">ADMIN (Carga y gestión)</span>'
                    : '<span class="badge bg-secondary">SALA (Solo lectura)</span>'}</td>
                <td>${eq.bloqueado_aqui
                    ? '<span class="badge bg-success"><i class="bi bi-lock-fill me-1"></i> Fijado y activo en este equipo</span>'
                    : '<span class="badge bg-dark border border-secondary text-secondary">Disponible</span>'}</td>
                <td><code>${d.ip_local}</code></td>
            </tr>
        `).join('');
    });
}

window.addEventListener('pywebviewready', () => {
    cargarEstadoRed();
    setInterval(cargarEstadoRed, 5000);
});
</script>
</body>
</html>"""

HTML_APP = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CongresoApp | Agenda por Sesiones</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-body: #111726;
            --bg-card: #1f2940;
            --bg-input: #151d30;
            --border-color: rgba(255, 255, 255, 0.08);
            --text-main: #ffffff;
            --text-muted: #94a3b8;
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
            cursor: pointer;
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
            border: 1px solid rgba(56, 189, 248, 0.25);
            border-radius: 20px;
            padding: 7px 16px;
            color: #fff;
            font-size: 0.88rem;
            width: 360px;
            max-width: 100%;
        }
        .valex-search:focus { outline: none; border-color: #00b9ff; box-shadow: 0 0 0 2px rgba(0, 185, 255, 0.18); }
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
            font-size: 0.84rem;
        }
        .btn-valex-blue { background: linear-gradient(to right, #0162e8, #00b9ff); border: none; color: #fff; font-weight: 600; }
        .btn-valex-green { background: linear-gradient(to right, #029666, #48d6a8); border: none; color: #fff; font-weight: 600; }
        .btn-valex-purple { background: linear-gradient(to right, #6d28d9, #a855f7); border: none; color: #fff; font-weight: 600; }
        .btn-valex-blue:hover, .btn-valex-green:hover, .btn-valex-purple:hover { color: #fff; opacity: 0.92; }
        .sesion-bloque {
            border-radius: 10px;
            margin-bottom: 18px;
            overflow: hidden;
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.28);
        }
        .sesion-bloque.tema-0 { background-color: #192642; border: 1px solid rgba(56, 189, 248, 0.35); border-left: 5px solid #00b9ff; }
        .sesion-bloque.tema-0 .sesion-header { background: linear-gradient(90deg, rgba(1, 98, 232, 0.34), rgba(0, 185, 255, 0.12)); }
        .sesion-bloque.tema-1 { background-
