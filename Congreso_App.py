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
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
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

CONFIG_POR_DEFECTO = {
    "workgroup": "CONGRESO_RED",
    "equipos": ["ADM1", "ADM2", "SALA_1", "SALA_2", "AUDITORIUM"],
    "nivel_1": "",
    "nivel_2": "",
    "nivel_3_columnas": [],
    "nivel_4_columnas": [],
    "separador_nivel_3": " - ",
    "max_chars_nivel_3": 50,
    "formato_dia_nn": True,
    "orden_carpetas": [],
}

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
                    f'{uri}</string></dict></plist>\n'
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


def localizar_carpeta_datos():
    return obtener_ruta_base()


def cargar_datos_congreso():
    ruta_activa = obtener_ruta_base()
    ruta_json = os.path.join(ruta_activa, "config_congreso.json")
    ruta_excel = os.path.join(ruta_activa, "programa.xlsx")

    if not os.path.exists(ruta_excel):
        return None, None, ruta_activa

    if os.path.exists(ruta_json):
        with open(ruta_json, "r", encoding="utf-8") as f:
            config = json.load(f)
    else:
        config = dict(CONFIG_POR_DEFECTO)

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
    config["_ruta_json"] = ruta_json
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


def abrir_en_explorador(ruta_abs):
    if SISTEMA == "Windows":
        os.startfile(ruta_abs)
    elif SISTEMA == "Darwin":
        subprocess.Popen(["open", ruta_abs])
    else:
        subprocess.Popen(["xdg-open", ruta_abs])


def limpiar_bloqueo_al_salir(ruta_archivo_bloqueo):
    try:
        if ruta_archivo_bloqueo and os.path.exists(ruta_archivo_bloqueo):
            os.remove(ruta_archivo_bloqueo)
    except Exception:
        pass


class CongresoAppEscritorio(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CongresoApp — Escritorio Nativo")
        self.geometry("1360x840")
        self.minsize(1050, 680)
        self.configure(bg="#111726")

        self.config_data = None
        self.df = None
        self.ruta_activa = obtener_ruta_base()
        self.sesiones_agrupadas = []
        self.filas_mostradas = []

        self._configurar_estilos()
        self._construir_interfaz()
        self.recargar_todo()
        self.protocol("WM_DELETE_WINDOW", self._al_cerrar)

    def _configurar_estilos(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook", background="#111726", borderwidth=0)
        style.configure("TNotebook.Tab", background="#1f2940", foreground="#cbd5e1", padding=[16, 8], font=("Roboto", 11, "bold"))
        style.map("TNotebook.Tab", background=[("selected", "#0162e8")], foreground=[("selected", "#ffffff")])
        style.configure("Treeview", background="#151d30", foreground="#f8fafc", fieldbackground="#151d30", rowheight=32, font=("Roboto", 11))
        style.configure("Treeview.Heading", background="#1f2940", foreground="#38bdf8", font=("Roboto", 11, "bold"))
        style.map("Treeview", background=[("selected", "#0284c7")])

    def _construir_interfaz(self):
        top = tk.Frame(self, bg="#1f2940", padx=18, pady=10)
        top.pack(fill="x")

        tk.Label(top, text="CongresoApp", bg="#1f2940", fg="#38bdf8", font=("Roboto", 16, "bold")).pack(side="left")
        self.lbl_estado_top = tk.Label(top, text="", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 11))
        self.lbl_estado_top.pack(side="right")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=14, pady=12)

        self.tab_agenda = tk.Frame(self.notebook, bg="#111726")
        self.tab_config = tk.Frame(self.notebook, bg="#111726")
        self.tab_red = tk.Frame(self.notebook, bg="#111726")

        self.notebook.add(self.tab_agenda, text="  Agenda y Ponentes  ")
        self.notebook.add(self.tab_config, text="  Configurador Congreso  ")
        self.notebook.add(self.tab_red, text="  Red y Equipos  ")

        self._construir_tab_agenda()
        self._construir_tab_configurador()
        self._construir_tab_red()

    def _construir_tab_agenda(self):
        bar_puesto = tk.Frame(self.tab_agenda, bg="#1f2940", padx=14, pady=10)
        bar_puesto.pack(fill="x", padx=6, pady=(4, 8))

        tk.Label(bar_puesto, text="Puesto:", bg="#1f2940", fg="#94a3b8", font=("Roboto", 11, "bold")).pack(side="left")
        self.combo_puesto = ttk.Combobox(bar_puesto, state="readonly", width=24, font=("Roboto", 11))
        self.combo_puesto.pack(side="left", padx=8)
        self.btn_fijar_puesto = tk.Button(
            bar_puesto, text="Confirmar y Bloquear", bg="#029666", fg="#fff",
            font=("Roboto", 10, "bold"), relief="flat", padx=12, pady=5, command=self.fijar_puesto
        )
        self.btn_fijar_puesto.pack(side="left", padx=4)

        self.btn_generar_carpetas = tk.Button(
            bar_puesto, text="Generar Carpetas", bg="#029666", fg="#fff",
            font=("Roboto", 10, "bold"), relief="flat", padx=12, pady=5, command=self.generar_carpetas
        )
        self.btn_generar_carpetas.pack(side="right", padx=4)

        self.btn_abrir_ppt = tk.Button(
            bar_puesto, text="Carpeta PPT", bg="#0162e8", fg="#fff",
            font=("Roboto", 10, "bold"), relief="flat", padx=12, pady=5, command=lambda: self.abrir_carpeta_rel("PPT", crear=True)
        )
        self.btn_abrir_ppt.pack(side="right", padx=4)

        bar_filtros = tk.Frame(self.tab_agenda, bg="#192238", padx=14, pady=8)
        bar_filtros.pack(fill="x", padx=6, pady=(0, 8))

        tk.Label(bar_filtros, text="Buscar:", bg="#192238", fg="#e2e8f0", font=("Roboto", 11)).pack(side="left")
        self.var_busqueda = tk.StringVar()
        self.var_busqueda.trace_add("write", lambda *_: self.aplicar_filtros())
        tk.Entry(bar_filtros, textvariable=self.var_busqueda, bg="#151d30", fg="#fff", insertbackground="#fff", font=("Roboto", 11), width=28).pack(side="left", padx=6)

        tk.Label(bar_filtros, text="Día:", bg="#192238", fg="#e2e8f0", font=("Roboto", 11)).pack(side="left", padx=(8, 2))
        self.combo_dia = ttk.Combobox(bar_filtros, state="readonly", width=14, font=("Roboto", 10))
        self.combo_dia.pack(side="left", padx=4)
        self.combo_dia.bind("<<ComboboxSelected>>", lambda *_: self.aplicar_filtros())

        tk.Label(bar_filtros, text="Sala:", bg="#192238", fg="#e2e8f0", font=("Roboto", 11)).pack(side="left", padx=(8, 2))
        self.combo_sala = ttk.Combobox(bar_filtros, state="readonly", width=22, font=("Roboto", 10))
        self.combo_sala.pack(side="left", padx=4)
        self.combo_sala.bind("<<ComboboxSelected>>", lambda *_: self.aplicar_filtros())

        tk.Button(bar_filtros, text="Limpiar", bg="#334155", fg="#fff", font=("Roboto", 10, "bold"), relief="flat", padx=10, pady=4, command=self.limpiar_filtros).pack(side="left", padx=8)

        panel_sesion = tk.Frame(self.tab_agenda, bg="#192642", padx=16, pady=12, highlightbackground="#00b9ff", highlightthickness=1)
        panel_sesion.pack(fill="x", padx=6, pady=(0, 8))
        self.lbl_sesion_grande = tk.Label(
            panel_sesion,
            text="Selecciona una sesión o ponente para ver su detalle ampliado",
            bg="#192642",
            fg="#ffffff",
            font=("Roboto", 14, "bold"),
            justify="left",
            anchor="w",
            wraplength=1280,
        )
        self.lbl_sesion_grande.pack(fill="x")
        self.lbl_ruta_grande = tk.Label(
            panel_sesion,
            text="",
            bg="#192642",
            fg="#fbbf24",
            font=("Consolas", 12),
            justify="left",
            anchor="w",
        )
        self.lbl_ruta_grande.pack(fill="x", pady=(4, 0))

        cols = ("sesion", "ponente", "estado", "archivos", "carpeta")
        self.tree = ttk.Treeview(self.tab_agenda, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("sesion", text="Sesión (Hora / Título / Sala / Día)")
        self.tree.heading("ponente", text="Participante / Ponente")
        self.tree.heading("estado", text="Estado PPT")
        self.tree.heading("archivos", text="Archivos")
        self.tree.heading("carpeta", text="Subcarpeta")

        self.tree.column("sesion", width=460, anchor="w")
        self.tree.column("ponente", width=240, anchor="w")
        self.tree.column("estado", width=150, anchor="center")
        self.tree.column("archivos", width=90, anchor="center")
        self.tree.column("carpeta", width=360, anchor="w")

        self.tree.pack(fill="both", expand=True, padx=6, pady=(0, 8))
        self.tree.bind("<<TreeviewSelect>>", self._on_seleccion_tree)

        bar_acc = tk.Frame(self.tab_agenda, bg="#1f2940", padx=14, pady=10)
        bar_acc.pack(fill="x", padx=6, pady=(0, 4))

        tk.Button(bar_acc, text="Abrir Carpeta Ponencia", bg="#334155", fg="#fff", font=("Roboto", 10, "bold"), relief="flat", padx=12, pady=6, command=self.abrir_carpeta_sesion_actual).pack(side="left", padx=4)
        tk.Button(bar_acc, text="Abrir Subcarpeta Participante", bg="#0162e8", fg="#fff", font=("Roboto", 10, "bold"), relief="flat", padx=12, pady=6, command=self.abrir_subcarpeta_actual).pack(side="left", padx=4)
        self.btn_subir = tk.Button(bar_acc, text="Subir Presentación (+ Backup)", bg="#029666", fg="#fff", font=("Roboto", 10, "bold"), relief="flat", padx=14, pady=6, command=self.subir_presentacion_actual)
        self.btn_subir.pack(side="left", padx=4)
        self.lbl_resumen_kpi = tk.Label(bar_acc, text="", bg="#1f2940", fg="#6ee7b7", font=("Roboto", 11, "bold"))
        self.lbl_resumen_kpi.pack(side="right")

    def _construir_tab_configurador(self):
        marco = tk.Frame(self.tab_config, bg="#1f2940", padx=20, pady=18)
        marco.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(marco, text="Configuración del Congreso (config_congreso.json)", bg="#1f2940", fg="#38bdf8", font=("Roboto", 14, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        tk.Label(marco, text="Workgroup / Red:", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 11)).grid(row=1, column=0, sticky="w", pady=6)
        self.ent_workgroup = tk.Entry(marco, bg="#151d30", fg="#fff", font=("Roboto", 11), width=45)
        self.ent_workgroup.grid(row=1, column=1, sticky="w", pady=6)

        tk.Label(marco, text="Equipos (separados por coma):", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 11)).grid(row=2, column=0, sticky="w", pady=6)
        self.ent_equipos = tk.Entry(marco, bg="#151d30", fg="#fff", font=("Roboto", 11), width=65)
        self.ent_equipos.grid(row=2, column=1, sticky="w", pady=6)

        tk.Label(marco, text="Columna Nivel 1 (Sala):", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 11)).grid(row=3, column=0, sticky="w", pady=6)
        self.combo_n1 = ttk.Combobox(marco, width=42, font=("Roboto", 11))
        self.combo_n1.grid(row=3, column=1, sticky="w", pady=6)

        tk.Label(marco, text="Columna Nivel 2 (Día):", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 11)).grid(row=4, column=0, sticky="w", pady=6)
        self.combo_n2 = ttk.Combobox(marco, width=42, font=("Roboto", 11))
        self.combo_n2.grid(row=4, column=1, sticky="w", pady=6)

        tk.Label(marco, text="Columnas Nivel 3 (separadas por coma):", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 11)).grid(row=5, column=0, sticky="w", pady=6)
        self.ent_n3 = tk.Entry(marco, bg="#151d30", fg="#fff", font=("Roboto", 11), width=65)
        self.ent_n3.grid(row=5, column=1, sticky="w", pady=6)

        tk.Label(marco, text="Columnas Nivel 4 Ponentes (separadas por coma):", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 11)).grid(row=6, column=0, sticky="w", pady=6)
        self.ent_n4 = tk.Entry(marco, bg="#151d30", fg="#fff", font=("Roboto", 11), width=65)
        self.ent_n4.grid(row=6, column=1, sticky="w", pady=6)

        tk.Label(marco, text="Separador Nivel 3:", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 11)).grid(row=7, column=0, sticky="w", pady=6)
        self.ent_sep = tk.Entry(marco, bg="#151d30", fg="#fff", font=("Roboto", 11), width=20)
        self.ent_sep.grid(row=7, column=1, sticky="w", pady=6)

        self.var_formato_dia = tk.BooleanVar(value=True)
        tk.Checkbutton(marco, text="Formatear días como DIA NN", variable=self.var_formato_dia, bg="#1f2940", fg="#e2e8f0", selectcolor="#151d30", font=("Roboto", 11)).grid(row=8, column=1, sticky="w", pady=8)

        tk.Button(marco, text="Guardar Configuración", bg="#029666", fg="#fff", font=("Roboto", 11, "bold"), relief="flat", padx=18, pady=8, command=self.guardar_configuracion).grid(row=9, column=1, sticky="w", pady=14)

    def _construir_tab_red(self):
        marco = tk.Frame(self.tab_red, bg="#1f2940", padx=20, pady=18)
        marco.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(marco, text="Estado de Red, Equipos y Backup Local", bg="#1f2940", fg="#38bdf8", font=("Roboto", 14, "bold")).pack(anchor="w", pady=(0, 10))
        self.txt_red = tk.Text(marco, bg="#151d30", fg="#e2e8f0", font=("Consolas", 11), height=22, relief="flat", padx=14, pady=12)
        self.txt_red.pack(fill="both", expand=True)
        tk.Button(marco, text="Refrescar Estado de Red y Backup", bg="#0162e8", fg="#fff", font=("Roboto", 10, "bold"), relief="flat", padx=14, pady=6, command=self.refrescar_panel_red).pack(anchor="w", pady=(10, 0))

    def _es_admin(self):
        puesto = (self.config_data or {}).get("_puesto_bloqueado") or ""
        return (not puesto) or str(puesto).upper().startswith("ADM")

    def recargar_todo(self):
        try:
            obtener_ruta_backup_local()
            crear_acceso_directo_escritorio()
            config, df, ruta_activa = cargar_datos_congreso()
            self.config_data = config
            self.df = df
            self.ruta_activa = ruta_activa

            if config is None or df is None:
                self.lbl_estado_top.config(text=f"No se encontró programa.xlsx en {ruta_activa}")
                self.refrescar_panel_red()
                return

            columnas_raw = list(df.columns)
            n3_cols = resolver_columnas_nivel_3(config, columnas_raw)
            if n3_cols:
                df = df.sort_values(
                    by=n3_cols,
                    key=lambda s: s.astype(str).str.replace("\xa0", " ", regex=False).str.lower().str.strip(),
                    kind="mergesort",
                )

            col_sala = config.get("nivel_1") or buscar_columna(columnas_raw, ["sala"])
            col_dia = config.get("nivel_2") or buscar_columna(columnas_raw, ["dia", "fecha"])
            formato_dia_nn = bool(config.get("formato_dia_nn", True))

            puesto_bloq = config.get("_puesto_bloqueado") or ""
            es_adm = self._es_admin()
            sala_restringida = None if (not puesto_bloq or es_adm) else sanitize_name(puesto_bloq)

            equipos = config.get("equipos", [])
            self.combo_puesto["values"] = equipos
            if puesto_bloq:
                self.combo_puesto.set(puesto_bloq)
                self.combo_puesto.config(state="disabled")
                self.btn_fijar_puesto.config(state="disabled")
            else:
                self.combo_puesto.config(state="readonly")
                self.btn_fijar_puesto.config(state="normal")

            self.btn_generar_carpetas.config(state="normal" if es_adm else "disabled")
            self.btn_subir.config(state="normal" if es_adm else "disabled")

            sesiones_map = {}
            sesiones_lista = []
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
                    resumen_partes = []
                    for c in n3_cols:
                        if c in row and str(row[c]).strip():
                            resumen_partes.append(f"{c}: {str(row[c]).strip()}")
                    if sala_val:
                        resumen_partes.append(f"Sala: {sala_val}")
                    if dia_val:
                        resumen_partes.append(f"Día: {dia_val}")

                    sesion_obj = {
                        "ruta_sesion": ruta_sesion,
                        "sala": sala_val,
                        "dia": dia_val,
                        "titulo_grande": "  |  ".join(resumen_partes) if resumen_partes else ruta_sesion,
                        "ponentes": [],
                        "_vistos": set(),
                    }
                    sesiones_map[ruta_sesion] = sesion_obj
                    sesiones_lista.append(sesion_obj)
                else:
                    sesion_obj = sesiones_map[ruta_sesion]

                participantes = obtener_participantes_fila(row, columnas_raw, config) or ["Sin ponente"]
                for p in participantes:
                    p_san = sanitize_name(p)
                    clave_p = re.sub(r"\s+", " ", p_san.lower()).strip()
                    if clave_p in sesion_obj["_vistos"]:
                        continue
                    sesion_obj["_vistos"].add(clave_p)
                    ruta_carpeta = os.path.join(ruta_sesion, p_san) if p_san != "_vacio" else ruta_sesion
                    num_arch = contar_archivos_en_carpeta(os.path.join(ruta_activa, ruta_carpeta))
                    sesion_obj["ponentes"].append({
                        "ponente": p,
                        "ruta_sesion": ruta_sesion,
                        "ruta_carpeta": ruta_carpeta,
                        "sala": sala_val,
                        "dia": dia_val,
                        "titulo_grande": sesion_obj["titulo_grande"],
                        "tiene_ppt": num_arch > 0,
                        "num_archivos": num_arch,
                    })

            self.sesiones_agrupadas = sesiones_lista

            self.combo_dia["values"] = [""] + dias_unicos
            self.combo_dia.set("")
            if sala_restringida:
                self.combo_sala["values"] = [sala_restringida]
                self.combo_sala.set(sala_restringida)
                self.combo_sala.config(state="disabled")
            else:
                self.combo_sala["values"] = [""] + salas_unicas
                self.combo_sala.set("")
                self.combo_sala.config(state="readonly")

            self.lbl_estado_top.config(
                text=f"IP: {obtener_ip_local()}  |  Puesto: {puesto_bloq or 'Sin fijar'} ({'ADMIN' if es_adm else 'SALA'})  |  Backup: {obtener_ruta_backup_local()}"
            )

            self.ent_workgroup.delete(0, "end")
            self.ent_workgroup.insert(0, config.get("workgroup", "CONGRESO_RED"))
            self.ent_equipos.delete(0, "end")
            self.ent_equipos.insert(0, ", ".join(config.get("equipos", [])))
            self.combo_n1["values"] = [""] + columnas_raw
            self.combo_n1.set(config.get("nivel_1", ""))
            self.combo_n2["values"] = [""] + columnas_raw
            self.combo_n2.set(config.get("nivel_2", ""))
            self.ent_n3.delete(0, "end")
            self.ent_n3.insert(0, ", ".join(config.get("nivel_3_columnas", [])))
            self.ent_n4.delete(0, "end")
            self.ent_n4.insert(0, ", ".join(config.get("nivel_4_columnas", [])))
            self.ent_sep.delete(0, "end")
            self.ent_sep.insert(0, config.get("separador_nivel_3", " - "))
            self.var_formato_dia.set(bool(config.get("formato_dia_nn", True)))

            self.aplicar_filtros()
            self.refrescar_panel_red()
        except Exception:
            messagebox.showerror("Error", traceback.format_exc())

    def aplicar_filtros(self):
        q = self.var_busqueda.get().lower().strip()
        f_dia = self.combo_dia.get().lower().strip()
        f_sala = self.combo_sala.get().lower().strip()

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.filas_mostradas = []

        cargados = 0
        total = 0

        for sesion in self.sesiones_agrupadas:
            if f_dia and sesion["dia"].lower() != f_dia:
                continue
            if f_sala and sesion["sala"].lower() != f_sala:
                continue

            for p in sesion["ponentes"]:
                total += 1
                if p["tiene_ppt"]:
                    cargados += 1
                texto = f"{sesion['titulo_grande']} {sesion['ruta_sesion']} {p['ponente']} {p['ruta_carpeta']}".lower()
                if q and q not in texto:
                    continue
                idx = len(self.filas_mostradas)
                self.filas_mostradas.append(p)
                estado_txt = f"CARGADA ({p['num_archivos']})" if p["tiene_ppt"] else "PENDIENTE"
                self.tree.insert("", "end", iid=str(idx), values=(sesion["titulo_grande"], p["ponente"], estado_txt, p["num_archivos"], p["ruta_carpeta"]))

        arch_bk = contar_archivos_en_carpeta(obtener_ruta_backup_local())
        self.lbl_resumen_kpi.config(text=f"Cargadas: {cargados}/{total}  |  Backup Local: {arch_bk} archivos")

        if self.filas_mostradas:
            self.tree.selection_set("0")
            self._mostrar_detalle_fila(self.filas_mostradas[0])
        else:
            self.lbl_sesion_grande.config(text="Sin resultados con el filtro actual")
            self.lbl_ruta_grande.config(text="")

    def _on_seleccion_tree(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.filas_mostradas):
            self._mostrar_detalle_fila(self.filas_mostradas[idx])

    def _mostrar_detalle_fila(self, fila):
        self.lbl_sesion_grande.config(text=fila["titulo_grande"])
        self.lbl_ruta_grande.config(text=f"Carpeta ponencia: {fila['ruta_sesion']}   |   Subcarpeta: {fila['ruta_carpeta']}   |   Ponente: {fila['ponente']}")

    def _fila_seleccionada(self):
        sel = self.tree.selection()
        if not sel:
            return None
        idx = int(sel[0])
        return self.filas_mostradas[idx] if 0 <= idx < len(self.filas_mostradas) else None

    def limpiar_filtros(self):
        self.var_busqueda.set("")
        self.combo_dia.set("")
        if str(self.combo_sala["state"]) != "disabled":
            self.combo_sala.set("")
        self.aplicar_filtros()

    def fijar_puesto(self):
        puesto = self.combo_puesto.get().strip()
        if not puesto or not self.config_data:
            return
        with open(self.config_data["_ruta_puesto_sesion"], "w", encoding="utf-8") as fp:
            json.dump({"puesto": puesto}, fp)
        messagebox.showinfo("Puesto fijado", f"Puesto '{puesto}' confirmado y bloqueado para esta sesión.")
        self.recargar_todo()

    def generar_carpetas(self):
        if not self._es_admin():
            messagebox.showwarning("Permisos", "Solo los puestos ADMIN pueden generar carpetas.")
            return
        if self.config_data is None or self.df is None:
            return

        obtener_ruta_backup_local()
        crear_acceso_directo_escritorio()
        columnas_raw = list(self.df.columns)
        creadas = 0
        vistas = set()

        for _, row in self.df.iterrows():
            ruta_sesion = construir_ruta_sesion(row, self.config_data, columnas_raw)
            ruta_abs = os.path.join(self.ruta_activa, ruta_sesion)
            if ruta_abs not in vistas and not os.path.exists(ruta_abs):
                os.makedirs(ruta_abs, exist_ok=True)
                vistas.add(ruta_abs)
                creadas += 1
            for p in obtener_participantes_fila(row, columnas_raw, self.config_data):
                p_san = sanitize_name(p)
                if p_san == "_vacio":
                    continue
                p_abs = os.path.join(ruta_abs, p_san)
                if p_abs not in vistas and not os.path.exists(p_abs):
                    os.makedirs(p_abs, exist_ok=True)
                    vistas.add(p_abs)
                    creadas += 1

        copiados = sincronizar_estructura_backup_local(self.ruta_activa)
        messagebox.showinfo("Carpetas listas", f"Carpetas nuevas: {creadas}. Backup sincronizado ({copiados} archivos nuevos, sin borrar nada).")
        self.recargar_todo()

    def abrir_carpeta_rel(self, ruta_rel, crear=True):
        ruta_abs = os.path.join(self.ruta_activa, ruta_rel)
        if crear and self._es_admin():
            os.makedirs(ruta_abs, exist_ok=True)
        elif not os.path.exists(ruta_abs):
            messagebox.showwarning("Carpeta no existe", f"No existe: {ruta_abs}")
            return
        abrir_en_explorador(ruta_abs)

    def abrir_carpeta_sesion_actual(self):
        fila = self._fila_seleccionada()
        if fila:
            self.abrir_carpeta_rel(fila["ruta_sesion"], crear=True)

    def abrir_subcarpeta_actual(self):
        fila = self._fila_seleccionada()
        if fila:
            self.abrir_carpeta_rel(fila["ruta_carpeta"], crear=True)

    def subir_presentacion_actual(self):
        if not self._es_admin():
            messagebox.showwarning("Permisos", "Solo los puestos ADMIN pueden subir presentaciones.")
            return
        fila = self._fila_seleccionada()
        if not fila:
            messagebox.showwarning("Selección", "Selecciona primero un participante en la agenda.")
            return

        origen = filedialog.askopenfilename(title=f"Seleccionar presentación para {fila['ponente']}")
        if not origen:
            return

        carpeta_destino = os.path.join(self.ruta_activa, fila["ruta_carpeta"])
        os.makedirs(carpeta_destino, exist_ok=True)
        nombre_base, ext = os.path.splitext(os.path.basename(origen))
        nombre_seguro = sanitize_name(nombre_base) + sanitize_name(ext).replace("_vacio", "")
        destino_ppt = os.path.join(carpeta_destino, nombre_seguro)
        shutil.copy2(origen, destino_ppt)

        ruta_bk_copia = copiar_a_backup_local_sin_borrar(fila["ruta_carpeta"], nombre_seguro, destino_ppt)
        crear_acceso_directo_escritorio()

        messagebox.showinfo("Presentación cargada", f"Guardado en:\n{destino_ppt}\n\nCopia de respaldo en:\n{ruta_bk_copia}")
        self.recargar_todo()

    def guardar_configuracion(self):
        base_cfg = dict(CONFIG_POR_DEFECTO)
        if self.config_data:
            for k, v in self.config_data.items():
                if not str(k).startswith("_"):
                    base_cfg[k] = v

        base_cfg["workgroup"] = self.ent_workgroup.get().strip() or "CONGRESO_RED"
        base_cfg["equipos"] = [x.strip() for x in self.ent_equipos.get().split(",") if x.strip()]
        base_cfg["nivel_1"] = self.combo_n1.get().strip()
        base_cfg["nivel_2"] = self.combo_n2.get().strip()
        base_cfg["nivel_3_columnas"] = [x.strip() for x in self.ent_n3.get().split(",") if x.strip()]
        base_cfg["nivel_4_columnas"] = [x.strip() for x in self.ent_n4.get().split(",") if x.strip()]
        base_cfg["separador_nivel_3"] = self.ent_sep.get()
        base_cfg["formato_dia_nn"] = bool(self.var_formato_dia.get())

        os.makedirs(self.ruta_activa, exist_ok=True)
        ruta_json = os.path.join(self.ruta_activa, "config_congreso.json")
        with open(ruta_json, "w", encoding="utf-8") as f:
            json.dump(base_cfg, f, ensure_ascii=False, indent=2)

        messagebox.showinfo("Configuración guardada", f"Guardado en {ruta_json}")
        self.recargar_todo()

    def refrescar_panel_red(self):
        ip = obtener_ip_local()
        host = socket.gethostname()
        ruta_bk = obtener_ruta_backup_local()
        acceso_esc = crear_acceso_directo_escritorio()
        arch_bk = contar_archivos_en_carpeta(ruta_bk)
        puesto_bloq = (self.config_data or {}).get("_puesto_bloqueado") or "Ninguno"
        equipos = (self.config_data or {}).get("equipos", [])

        lineas = [
            f"IP Local:            {ip}",
            f"Nombre Equipo:       {host}",
            f"Workgroup:           {(self.config_data or {}).get('workgroup', 'N/A')}",
            f"Carpeta Ejecutable:  {self.ruta_activa}",
            f"Excel Esperado:      {os.path.join(self.ruta_activa, 'programa.xlsx')}",
            f"Carpeta PPT:         {os.path.join(self.ruta_activa, 'PPT')}",
            f"Backup Local Raíz:   {ruta_bk} ({arch_bk} archivos — nunca borra nada)",
            f"Acceso Escritorio:   {acceso_esc or 'No disponible'}",
            f"Puesto Fijado Aquí:  {puesto_bloq}",
            "",
            "--- ESTADO DE PUESTOS CONFIGURADOS ---",
        ]
        for eq in equipos:
            rol = "ADMIN (Carga y carpetas)" if str(eq).upper().startswith("ADM") else "SALA (Solo lectura)"
            estado = "ACTIVO Y BLOQUEADO EN ESTE EQUIPO" if str(eq).upper() == str(puesto_bloq).upper() else "Disponible"
            lineas.append(f"  * {eq:<16} | {rol:<26} | {estado}")

        self.txt_red.delete("1.0", "end")
        self.txt_red.insert("1.0", "\n".join(lineas))

    def _al_cerrar(self):
        ruta_bloq = (self.config_data or {}).get("_ruta_puesto_sesion")
        limpiar_bloqueo_al_salir(ruta_bloq)
        self.destroy()


if __name__ == "__main__":
    app = CongresoAppEscritorio()
    app.mainloop()
