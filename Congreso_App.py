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
        ruta = os.path.dirname(os.path.abspath(sys.executable))
        if SISTEMA == "Darwin":
            partes = ruta.split(os.sep)
            for i, parte in enumerate(partes):
                if parte.endswith(".app"):
                    return os.path.abspath(os.sep.join(partes[:i])) or "/"
            if ruta.endswith("Contents/MacOS"):
                return os.path.abspath(os.path.join(ruta, "..", "..", ".."))
        return ruta
    return os.path.abspath(os.path.dirname(__file__))


def _candidatos_carpeta_datos():
    base = obtener_ruta_base()
    candidatos = [
        base,
        os.path.abspath(os.path.join(base, "..")),
        os.path.abspath(os.getcwd()),
    ]
    unicos = []
    for c in candidatos:
        if c and c not in unicos:
            unicos.append(c)
    return unicos


def cargar_html_ui(nombre_archivo):
    bases_ui = []
    if getattr(sys, "_MEIPASS", None):
        bases_ui.append(os.path.join(sys._MEIPASS, "ui"))
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        bases_ui.append(os.path.join(exe_dir, "ui"))
        bases_ui.append(os.path.join(exe_dir, "_internal", "ui"))
        bases_ui.append(os.path.abspath(os.path.join(exe_dir, "..", "Resources", "ui")))
    bases_ui.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "ui"))

    for carpeta in bases_ui:
        ruta = os.path.join(carpeta, nombre_archivo)
        if os.path.exists(ruta):
            with open(ruta, "r", encoding="utf-8") as f:
                return f.read()
    raise FileNotFoundError(f"No se encontró ui/{nombre_archivo} en ninguna ruta conocida.")


HTML_ERROR = """<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Error en Congreso_App</title></head>
<body style="background:#111726;color:#fff;font-family:sans-serif;padding:30px;">
  <h2 style="color:#f87171;">Se produjo un error al cargar Congreso_App</h2>
  <p>Ruta activa detectada: <code>{{ ruta_activa }}</code></p>
  <pre style="background:#1e293b;padding:16px;border-radius:8px;overflow:auto;color:#fca5a5;">{{ detalle_error }}</pre>
</body>
</html>"""

HTML_APP = cargar_html_ui("app.html")
HTML_RED = cargar_html_ui("red.html")


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
    candidatos = _candidatos_carpeta_datos()
    ruta_activa = candidatos[0] if candidatos else obtener_ruta_base()

    for r in candidatos:
        ruta_json = os.path.join(r, "config_congreso.json")
        ruta_excel = os.path.join(r, "programa.xlsx")
        if os.path.exists(ruta_json) and os.path.exists(ruta_excel):
            ruta_activa = r
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

    return None, None, ruta_activa


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


class ApiBridge:
    def __init__(self, window):
        self._window = window

    def ir_a_red(self):
        self._window.load_html(HTML_RED)

    def ir_a_agenda(self):
        self._window.load_html(HTML_APP)

    def obtener_estado_red(self):
        config, _, ruta_activa = cargar_datos_congreso()
        return _construir_info_red(config, ruta_activa)

    def obtener_datos_agenda(self):
        ruta_activa = obtener_ruta_base()
        try:
            config, df, ruta_activa = cargar_datos_congreso()
            ruta_backup = obtener_ruta_backup_local()
            crear_acceso_directo_escritorio()

            if config is None or df is None:
                return {
                    "ok": True,
                    "config_cargada": False,
                    "ruta_activa": ruta_activa,
                    "ruta_backup": ruta_backup,
                }

            columnas_raw = list(df.columns)
            n3_cols_resueltas = resolver_columnas_nivel_3(config, columnas_raw)
            columnas_ordenadas = ordenar_columnas_por_nivel_3(columnas_raw, n3_cols_resueltas, config)

            if n3_cols_resueltas:
                df = df.sort_values(
                    by=n3_cols_resueltas,
                    key=lambda s: s.astype(str).str.replace("\xa0", " ", regex=False).str.lower().str.strip(),
                    kind="mergesort",
                )

            col_sala = config.get("nivel_1") or buscar_columna(columnas_raw, ["sala"])
            col_dia = config.get("nivel_2") or buscar_columna(columnas_raw, ["dia", "fecha"])
            col_titulo_ponencia = buscar_columna(columnas_raw, ["titulo ponencia", "ponencia", "comunicacion", "titulo charla"])
            formato_dia_nn = bool(config.get("formato_dia_nn", True))

            puesto_bloqueado = config.get("_puesto_bloqueado")
            puesto_activo = puesto_bloqueado or ""
            es_admin = str(puesto_activo).upper().startswith("ADM") or puesto_activo == ""
            sala_restringida = None if (not puesto_activo or es_admin) else sanitize_name(puesto_activo)

            sesiones_map = {}
            sesiones_agrupadas = []
            dias_unicos = []
            salas_unicas = []

            for _, row in df.iterrows():
                ruta_sesion = construir_ruta_sesion(row, config, columnas_raw)
                sala_val = sanitize_name(row[col_sala]) if (col_sala and col_sala in row and str(row[col_sala]).strip()) else ""
                if sala_val == "_vacio":
                    sala_val = ""
                raw_dia = row[col_dia] if col_dia and col_dia in row else ""
                dia_val = formatear_dia_valor(raw_dia) if (col_dia and formato_dia_nn) else (str(raw_dia).strip() or "Sin Día")

                if dia_val and dia_val not in dias_unicos:
                    dias_unicos.append(dia_val)
                if sala_val and sala_val not in salas_unicas:
                    salas_unicas.append(sala_val)

                if ruta_sesion not in sesiones_map:
                    campos_n3 = []
                    for c in (n3_cols_resueltas or columnas_ordenadas[:5]):
                        if c in row:
                            val_c = str(row[c]).replace("\xa0", " ").strip()
                            if val_c:
                                campos_n3.append({"etiqueta": c, "valor": val_c})

                    meta_textos = [f"{m['etiqueta']} {m['valor']}" for m in campos_n3]
                    texto_sesion = " ".join([ruta_sesion, sala_val, dia_val] + meta_textos)
                    sesion_obj = {
                        "ruta_sesion": ruta_sesion,
                        "sala": sala_val,
                        "dia": dia_val,
                        "campos_nivel_3": campos_n3,
                        "texto_busqueda": texto_sesion,
                        "ponentes": [],
                        "_ponentes_vistos": set(),
                    }
                    sesiones_map[ruta_sesion] = sesion_obj
                    sesiones_agrupadas.append(sesion_obj)
                else:
                    sesion_obj = sesiones_map[ruta_sesion]

                participantes = obtener_participantes_fila(row, columnas_raw, config) or ["Sin ponente"]
                detalle_extra = str(row[col_titulo_ponencia]).replace("\xa0", " ").strip() if (col_titulo_ponencia and col_titulo_ponencia in row) else ""
                datos_fila_txt = " ".join(str(row[c]).replace("\xa0", " ") for c in columnas_raw if c in row)

                for p in participantes:
                    p_san = sanitize_name(p)
                    clave_p = re.sub(r"\s+", " ", p_san.lower()).strip()
                    if clave_p in sesion_obj["_ponentes_vistos"]:
                        continue
                    sesion_obj["_ponentes_vistos"].add(clave_p)

                    ruta_carpeta = os.path.join(ruta_sesion, p_san) if p_san != "_vacio" else ruta_sesion
                    ruta_abs = os.path.join(ruta_activa, ruta_carpeta)
                    num_archivos = contar_archivos_en_carpeta(ruta_abs)
                    tiene_ppt = num_archivos > 0

                    sesion_obj["ponentes"].append({
                        "ponente": p,
                        "detalle_extra": detalle_extra,
                        "ruta_carpeta": ruta_carpeta,
                        "tiene_ppt": tiene_ppt,
                        "num_archivos": num_archivos,
                        "texto_busqueda": f"{p} {detalle_extra} {datos_fila_txt}",
                    })

            for s in sesiones_agrupadas:
                s.pop("_ponentes_vistos", None)

            total_ponentes = sum(len(s["ponentes"]) for s in sesiones_agrupadas)
            con_archivo = sum(1 for s in sesiones_agrupadas for p in s["ponentes"] if p["tiene_ppt"])
            archivos_backup = contar_archivos_en_carpeta(ruta_backup)

            return {
                "ok": True,
                "config_cargada": True,
                "workgroup": config.get("workgroup", "N/A"),
                "equipos": config.get("equipos", []),
                "sesiones_agrupadas": sesiones_agrupadas,
                "total_ponentes": total_ponentes,
                "con_archivo": con_archivo,
                "dias_unicos": dias_unicos,
                "salas_unicas": salas_unicas,
                "puesto_activo": puesto_activo,
                "puesto_bloqueado": puesto_bloqueado,
                "sala_restringida": sala_restringida,
                "es_admin": es_admin,
                "archivos_backup": archivos_backup,
                "ruta_backup": ruta_backup,
            }
        except Exception:
            self._window.load_html(
                _render_simple(HTML_ERROR, ruta_activa=ruta_activa, detalle_error=traceback.format_exc())
            )
            return {"ok": False}

    def fijar_puesto(self, puesto):
        config, _, _ = cargar_datos_congreso()
        if config is None:
            return {"ok": False, "mensaje": "No hay configuración cargada."}
        puesto = (puesto or "").strip()
        if puesto not in config.get("equipos", []):
            return {"ok": False, "mensaje": "Puesto no válido."}

        with open(config["_ruta_puesto_sesion"], "w", encoding="utf-8") as fp:
            json.dump({"puesto": puesto}, fp)

        if not puesto.upper().startswith("ADM"):
            return {"ok": True, "mensaje": f"Sala '{puesto}' fijada y bloqueada (solo lectura)."}
        return {"ok": True, "mensaje": f"Puesto ADMIN '{puesto}' activado."}

    def crear_carpetas(self):
        config, df, ruta_activa = cargar_datos_congreso()
        if config is None:
            return {"ok": False, "mensaje": "Faltan archivos de configuración."}
        puesto = config.get("_puesto_bloqueado") or ""
        if puesto and not puesto.upper().startswith("ADM"):
            return {"ok": False, "mensaje": "Solo los puestos ADMIN pueden crear o modificar carpetas."}

        ruta_bk = obtener_ruta_backup_local()
        crear_acceso_directo_escritorio()

        columnas_raw = list(df.columns)
        creadas = 0
        rutas_creadas = set()

        for _, row in df.iterrows():
            ruta_sesion = construir_ruta_sesion(row, config, columnas_raw)
            ruta_abs = os.path.join(ruta_activa, ruta_sesion)
            if ruta_abs not in rutas_creadas and not os.path.exists(ruta_abs):
                os.makedirs(ruta_abs, exist_ok=True)
                rutas_creadas.add(ruta_abs)
                creadas += 1

            for p in obtener_participantes_fila(row, columnas_raw, config):
                p_san = sanitize_name(p)
                if p_san == "_vacio":
                    continue
                p_abs = os.path.join(ruta_abs, p_san)
                if p_abs not in rutas_creadas and not os.path.exists(p_abs):
                    os.makedirs(p_abs, exist_ok=True)
                    rutas_creadas.add(p_abs)
                    creadas += 1

        copiados_bk = sincronizar_estructura_backup_local(ruta_activa)
        return {
            "ok": True,
            "mensaje": f"Carpetas listas ({creadas} nuevas). {ruta_bk} sincronizado ({copiados_bk} archivos copiados, sin borrar nada).",
        }

    def abrir_ruta(self, ruta_rel, crear=True):
        config, _, ruta_activa = cargar_datos_congreso()
        ruta_rel = ruta_rel or "PPT"
        puesto = config.get("_puesto_bloqueado") or "" if config else ""
        es_admin = (not puesto) or puesto.upper().startswith("ADM")

        ruta_abs = os.path.join(ruta_activa, ruta_rel)
        if crear and es_admin:
            os.makedirs(ruta_abs, exist_ok=True)
        elif not os.path.exists(ruta_abs):
            return {"ok": False, "mensaje": "La carpeta aún no existe en disco."}

        if SISTEMA == "Windows":
            os.startfile(ruta_abs)
        elif SISTEMA == "Darwin":
            subprocess.Popen(["open", ruta_abs])
        else:
            subprocess.Popen(["xdg-open", ruta_abs])
        return {"ok": True}

    def seleccionar_y_subir_presentacion(self, ruta_rel):
        config, _, ruta_activa = cargar_datos_congreso()
        puesto = config.get("_puesto_bloqueado") or "" if config else ""
        if puesto and not puesto.upper().startswith("ADM"):
            return {"ok": False, "mensaje": "Acceso denegado: las salas solo pueden ver presentaciones."}

        resultado = self._window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False)
        if not resultado:
            return {"ok": False, "cancelado": True}

        archivo_origen = resultado[0] if isinstance(resultado, (list, tuple)) else resultado
        if not archivo_origen or not os.path.exists(archivo_origen):
            return {"ok": False, "mensaje": "No se pudo leer el archivo seleccionado."}

        ruta_rel = ruta_rel or "PPT"
        carpeta_destino = os.path.join(ruta_activa, ruta_rel)
        os.makedirs(carpeta_destino, exist_ok=True)
        nombre_base, ext = os.path.splitext(os.path.basename(archivo_origen))
        nombre_seguro = sanitize_name(nombre_base) + sanitize_name(ext).replace("_vacio", "")
        destino_ppt_abs = os.path.join(carpeta_destino, nombre_seguro)

        shutil.copy2(archivo_origen, destino_ppt_abs)

        ruta_bk_copia = copiar_a_backup_local_sin_borrar(ruta_rel, nombre_seguro, destino_ppt_abs)
        crear_acceso_directo_escritorio()

        return {
            "ok": True,
            "mensaje": f"Presentación '{nombre_seguro}' cargada en {ruta_rel} y copiada a {ruta_bk_copia} (sin borrar nada).",
        }


def limpiar_bloqueo_al_salir(ruta_archivo_bloqueo):
    try:
        if ruta_archivo_bloqueo and os.path.exists(ruta_archivo_bloqueo):
            os.remove(ruta_archivo_bloqueo)
    except Exception:
        pass


if __name__ == "__main__":
    cfg, _, _ = cargar_datos_congreso()
    try:
        obtener_ruta_backup_local()
        crear_acceso_directo_escritorio()
    except Exception:
        pass

    ruta_bloq = cfg.get("_ruta_puesto_sesion") if cfg else None
    limpiar_bloqueo_al_salir(ruta_bloq)

    window = webview.create_window("CongresoApp | Agenda por Sesiones", html=HTML_APP, width=1360, height=860)
    api = ApiBridge(window)
    window.expose(
        api.ir_a_red,
        api.ir_a_agenda,
        api.obtener_estado_red,
        api.obtener_datos_agenda,
        api.fijar_puesto,
        api.crear_carpetas,
        api.abrir_ruta,
        api.seleccionar_y_subir_presentacion,
    )
    try:
        webview.start()
    finally:
        limpiar_bloqueo_al_salir(ruta_bloq)
