import os
import re
import sys
import json
import shutil
import socket
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


def abrir_carpeta_en_sistema(ruta_abs):
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


class CongresoAppEscritorio:
    def __init__(self, root):
        self.root = root
        self.root.title("CongresoApp | Agenda por Sesiones")
        self.root.geometry("1220x780")
        self.root.configure(bg="#111726")

        self.config, self.df, self.ruta_activa = cargar_datos_congreso()
        self.ruta_backup = obtener_ruta_backup_local()
        crear_acceso_directo_escritorio()

        self.ruta_bloqueo = self.config.get("_ruta_puesto_sesion") if self.config else None
        limpiar_bloqueo_al_salir(self.ruta_bloqueo)
        if self.config is not None:
            self.config["_puesto_bloqueado"] = None

        self.filas_agenda = []
        self._configurar_estilos()
        self._construir_interfaz()
        self._cargar_agenda()
        self.root.protocol("WM_DELETE_WINDOW", self._al_cerrar)

    def _configurar_estilos(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TNotebook", background="#1f2940", borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            background="#151d30",
            foreground="#cbd5e1",
            padding=[14, 8],
            font=("Roboto", 10, "bold"),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#0162e8")],
            foreground=[("selected", "#ffffff")],
        )
        style.configure(
            "Treeview",
            background="#151d30",
            foreground="#f8fafc",
            fieldbackground="#151d30",
            rowheight=30,
            font=("Roboto", 10),
        )
        style.configure(
            "Treeview.Heading",
            background="#1f2940",
            foreground="#38bdf8",
            font=("Roboto", 10, "bold"),
        )
        style.map("Treeview", background=[("selected", "#0162e8")], foreground=[("selected", "#ffffff")])

    def _puesto_actual(self):
        return (self.config or {}).get("_puesto_bloqueado") or ""

    def _es_admin(self):
        puesto = self._puesto_actual()
        return (not puesto) or str(puesto).upper().startswith("ADM")

    def _sala_restringida(self):
        puesto = self._puesto_actual()
        if not puesto or self._es_admin():
            return None
        return sanitize_name(puesto).lower()

    def _construir_interfaz(self):
        topbar = tk.Frame(self.root, bg="#1f2940", padx=18, pady=10)
        topbar.pack(fill="x")

        tk.Label(
            topbar,
            text="CongresoApp",
            bg="#1f2940",
            fg="#ffffff",
            font=("Roboto", 16, "bold"),
        ).pack(side="left")

        ip_local = obtener_ip_local()
        wg = (self.config or {}).get("workgroup", "N/A")
        self.lbl_estado_top = tk.Label(
            topbar,
            text=f"IP: {ip_local}  |  Grupo: {wg}  |  Backup: {self.ruta_backup}",
            bg="#1f2940",
            fg="#94a3b8",
            font=("Roboto", 9),
        )
        self.lbl_estado_top.pack(side="right")

        if self.config is None or self.df is None:
            aviso = tk.Frame(self.root, bg="#1f2940", padx=24, pady=24)
            aviso.pack(fill="both", expand=True, padx=20, pady=20)
            tk.Label(
                aviso,
                text="No se encontraron programa.xlsx y config_congreso.json",
                bg="#1f2940",
                fg="#fbbf24",
                font=("Roboto", 14, "bold"),
            ).pack(pady=(0, 8))
            tk.Label(
                aviso,
                text=f"Colócalos en la misma carpeta del ejecutable:\n{self.ruta_activa}",
                bg="#1f2940",
                fg="#e2e8f0",
                font=("Roboto", 11),
            ).pack(pady=(0, 16))
            tk.Button(
                aviso,
                text="Reintentar",
                bg="#0162e8",
                fg="#ffffff",
                relief="flat",
                padx=16,
                pady=8,
                command=self._recargar_todo,
            ).pack()
            return

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=14, pady=12)

        self.tab_agenda = tk.Frame(notebook, bg="#111726")
        self.tab_red = tk.Frame(notebook, bg="#111726")
        notebook.add(self.tab_agenda, text="Agenda y Ponentes")
        notebook.add(self.tab_red, text="Red, Puestos y Backup")

        self._construir_tab_agenda()
        self._construir_tab_red()

    def _construir_tab_agenda(self):
        puesto_box = tk.Frame(self.tab_agenda, bg="#1f2940", padx=14, pady=10)
        puesto_box.pack(fill="x", padx=6, pady=(6, 10))

        tk.Label(
            puesto_box,
            text="Puesto de trabajo:",
            bg="#1f2940",
            fg="#ffffff",
            font=("Roboto", 10, "bold"),
        ).pack(side="left", padx=(0, 8))

        equipos = self.config.get("equipos", [])
        self.combo_puesto = ttk.Combobox(puesto_box, values=equipos, state="readonly", width=22)
        if equipos:
            self.combo_puesto.current(0)
        self.combo_puesto.pack(side="left", padx=(0, 8))

        self.btn_fijar_puesto = tk.Button(
            puesto_box,
            text="Confirmar y Bloquear",
            bg="#029666",
            fg="#ffffff",
            relief="flat",
            padx=12,
            pady=4,
            command=self._fijar_puesto,
        )
        self.btn_fijar_puesto.pack(side="left", padx=(0, 14))

        self.lbl_rol_puesto = tk.Label(
            puesto_box,
            text="Sin fijar (ADMIN)",
            bg="#1f2940",
            fg="#38bdf8",
            font=("Roboto", 10, "bold"),
        )
        self.lbl_rol_puesto.pack(side="left")

        kpi_frame = tk.Frame(self.tab_agenda, bg="#111726")
        kpi_frame.pack(fill="x", padx=6, pady=(0, 10))

        self.lbl_kpi_sesiones = tk.Label(
            kpi_frame, text="Sesiones: 0", bg="#0162e8", fg="#ffffff", font=("Roboto", 11, "bold"), padx=14, pady=8
        )
        self.lbl_kpi_sesiones.pack(side="left", padx=(0, 10))

        self.lbl_kpi_ponentes = tk.Label(
            kpi_frame, text="Ponentes: 0", bg="#d9264a", fg="#ffffff", font=("Roboto", 11, "bold"), padx=14, pady=8
        )
        self.lbl_kpi_ponentes.pack(side="left", padx=(0, 10))

        self.lbl_kpi_cargadas = tk.Label(
            kpi_frame, text="Cargadas: 0 / 0", bg="#029666", fg="#ffffff", font=("Roboto", 11, "bold"), padx=14, pady=8
        )
        self.lbl_kpi_cargadas.pack(side="left", padx=(0, 10))

        self.lbl_kpi_backup = tk.Label(
            kpi_frame, text="Backup: 0 arch.", bg="#d95f18", fg="#ffffff", font=("Roboto", 11, "bold"), padx=14, pady=8
        )
        self.lbl_kpi_backup.pack(side="left")

        filtros = tk.Frame(self.tab_agenda, bg="#1f2940", padx=14, pady=10)
        filtros.pack(fill="x", padx=6, pady=(0, 10))

        tk.Label(filtros, text="Buscar:", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 9, "bold")).pack(side="left", padx=(0, 4))
        self.var_busqueda = tk.StringVar()
        self.var_busqueda.trace_add("write", lambda *_: self._refrescar_tabla())
        tk.Entry(filtros, textvariable=self.var_busqueda, bg="#151d30", fg="#ffffff", insertbackground="#ffffff", width=26).pack(
            side="left", padx=(0, 10)
        )

        tk.Label(filtros, text="Día:", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 9, "bold")).pack(side="left", padx=(0, 4))
        self.var_dia = tk.StringVar(value="Todos")
        self.combo_dia = ttk.Combobox(filtros, textvariable=self.var_dia, state="readonly", width=14)
        self.combo_dia.bind("<<ComboboxSelected>>", lambda *_: self._refrescar_tabla())
        self.combo_dia.pack(side="left", padx=(0, 10))

        tk.Label(filtros, text="Sala:", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 9, "bold")).pack(side="left", padx=(0, 4))
        self.var_sala = tk.StringVar(value="Todas")
        self.combo_sala = ttk.Combobox(filtros, textvariable=self.var_sala, state="readonly", width=22)
        self.combo_sala.bind("<<ComboboxSelected>>", lambda *_: self._refrescar_tabla())
        self.combo_sala.pack(side="left", padx=(0, 10))

        tk.Label(filtros, text="Estado:", bg="#1f2940", fg="#e2e8f0", font=("Roboto", 9, "bold")).pack(side="left", padx=(0, 4))
        self.var_estado = tk.StringVar(value="Todos")
        self.combo_estado = ttk.Combobox(
            filtros, textvariable=self.var_estado, values=["Todos", "CARGADA", "PENDIENTE"], state="readonly", width=13
        )
        self.combo_estado.bind("<<ComboboxSelected>>", lambda *_: self._refrescar_tabla())
        self.combo_estado.pack(side="left", padx=(0, 10))

        tk.Button(
            filtros, text="Limpiar", bg="#334155", fg="#ffffff", relief="flat", padx=10, pady=3, command=self._limpiar_filtros
        ).pack(side="left")

        tabla_frame = tk.Frame(self.tab_agenda, bg="#111726")
        tabla_frame.pack(fill="both", expand=True, padx=6, pady=(0, 10))

        columnas = ("estado", "participante", "sesion", "sala", "dia", "carpeta")
        self.tree = ttk.Treeview(tabla_frame, columns=columnas, show="headings", selectmode="browse")
        self.tree.heading("estado", text="ESTADO")
        self.tree.heading("participante", text="PARTICIPANTE / PONENTE")
        self.tree.heading("sesion", text="SESIÓN")
        self.tree.heading("sala", text="SALA")
        self.tree.heading("dia", text="DÍA")
        self.tree.heading("carpeta", text="SUBCARPETA PPT")

        self.tree.column("estado", width=120, anchor="center")
        self.tree.column("participante", width=230)
        self.tree.column("sesion", width=340)
        self.tree.column("sala", width=170)
        self.tree.column("dia", width=100, anchor="center")
        self.tree.column("carpeta", width=240)

        scroll_y = ttk.Scrollbar(tabla_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        acciones = tk.Frame(self.tab_agenda, bg="#1f2940", padx=14, pady=10)
        acciones.pack(fill="x", padx=6, pady=(0, 6))

        self.btn_generar = tk.Button(
            acciones,
            text="Generar Carpetas",
            bg="#029666",
            fg="#ffffff",
            relief="flat",
            padx=12,
            pady=6,
            command=self._generar_carpetas,
        )
        self.btn_generar.pack(side="left", padx=(0, 8))

        tk.Button(
            acciones,
            text="Abrir Carpeta PPT",
            bg="#0162e8",
            fg="#ffffff",
            relief="flat",
            padx=12,
            pady=6,
            command=lambda: self._abrir_carpeta_rel("PPT"),
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            acciones,
            text="Abrir Subcarpeta Seleccionada",
            bg="#2563eb",
            fg="#ffffff",
            relief="flat",
            padx=12,
            pady=6,
            command=self._abrir_subcarpeta_seleccionada,
        ).pack(side="left", padx=(0, 8))

        self.btn_subir = tk.Button(
            acciones,
            text="Subir Presentación...",
            bg="#059669",
            fg="#ffffff",
            relief="flat",
            padx=12,
            pady=6,
            command=self._subir_presentacion,
        )
        self.btn_subir.pack(side="left", padx=(0, 8))

        tk.Button(
            acciones,
            text="Actualizar Estado",
            bg="#475569",
            fg="#ffffff",
            relief="flat",
            padx=12,
            pady=6,
            command=self._cargar_agenda,
        ).pack(side="right")

    def _construir_tab_red(self):
        panel = tk.Frame(self.tab_red, bg="#1f2940", padx=20, pady=20)
        panel.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(
            panel,
            text="Estado de Equipo, Rutas y Backup Local",
            bg="#1f2940",
            fg="#ffffff",
            font=("Roboto", 13, "bold"),
        ).pack(anchor="w", pady=(0, 12))

        self.txt_red = tk.Text(
            panel,
            bg="#151d30",
            fg="#e2e8f0",
            font=("Consolas", 10),
            height=22,
            relief="flat",
            padx=14,
            pady=14,
        )
        self.txt_red.pack(fill="both", expand=True, pady=(0, 12))

        tk.Button(
            panel,
            text="Abrir Carpeta BACKUP_LOCAL",
            bg="#0162e8",
            fg="#ffffff",
            relief="flat",
            padx=14,
            pady=6,
            command=lambda: abrir_carpeta_en_sistema(self.ruta_backup),
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            panel,
            text="Refrescar Información",
            bg="#475569",
            fg="#ffffff",
            relief="flat",
            padx=14,
            pady=6,
            command=self._actualizar_info_red,
        ).pack(side="left")

    def _actualizar_permisos_ui(self):
        es_adm = self._es_admin()
        puesto = self._puesto_actual()
        if puesto:
            self.lbl_rol_puesto.config(
                text=f"Bloqueado en {puesto} ({'ADMIN' if es_adm else 'SALA - SOLO LECTURA'})",
                fg="#48d6a8" if es_adm else "#fbbf24",
            )
            self.btn_fijar_puesto.config(state="disabled")
            self.combo_puesto.config(state="disabled")
        else:
            self.lbl_rol_puesto.config(text="Sin fijar (ADMIN)", fg="#38bdf8")

        self.btn_generar.config(state="normal" if es_adm else "disabled")
        self.btn_subir.config(state="normal" if es_adm else "disabled")

        sala_bloq = self._sala_restringida()
        if sala_bloq:
            for idx, val in enumerate(self.combo_sala["values"]):
                if val.lower() == sala_bloq:
                    self.combo_sala.current(idx)
                    break
            self.combo_sala.config(state="disabled")
        else:
            self.combo_sala.config(state="readonly")

    def _fijar_puesto(self):
        puesto = self.combo_puesto.get().strip()
        if not puesto:
            return
        if self.ruta_bloqueo:
            with open(self.ruta_bloqueo, "w", encoding="utf-8") as fp:
                json.dump({"puesto": puesto}, fp)
        self.config["_puesto_bloqueado"] = puesto
        self._actualizar_permisos_ui()
        self._refrescar_tabla()
        messagebox.showinfo("Puesto fijado", f"Puesto '{puesto}' fijado para esta sesión.")

    def _cargar_agenda(self):
        if self.config is None or self.df is None:
            return

        columnas_raw = list(self.df.columns)
        n3_cols = resolver_columnas_nivel_3(self.config, columnas_raw)
        df_ordenado = self.df
        if n3_cols:
            df_ordenado = self.df.sort_values(
                by=n3_cols,
                key=lambda s: s.astype(str).str.replace("\xa0", " ", regex=False).str.lower().str.strip(),
                kind="mergesort",
            )

        col_sala = self.config.get("nivel_1") or buscar_columna(columnas_raw, ["sala"])
        col_dia = self.config.get("nivel_2") or buscar_columna(columnas_raw, ["dia", "fecha"])
        formato_dia_nn = bool(self.config.get("formato_dia_nn", True))

        sesiones_map = {}
        dias_unicos = ["Todos"]
        salas_unicas = ["Todas"]
        filas = []

        for _, row in df_ordenado.iterrows():
            ruta_sesion = construir_ruta_sesion(row, self.config, columnas_raw)
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
                campos_txt = []
                for c in n3_cols:
                    if c in row and str(row[c]).strip():
                        campos_txt.append(f"{c}: {str(row[c]).replace(chr(160), ' ').strip()}")
                sesiones_map[ruta_sesion] = {
                    "titulo": " | ".join(campos_txt) if campos_txt else ruta_sesion,
                    "vistos": set(),
                }

            sesion_info = sesiones_map[ruta_sesion]
            participantes = obtener_participantes_fila(row, columnas_raw, self.config) or ["Sin ponente"]
            datos_fila = " ".join(str(row[c]).replace("\xa0", " ") for c in columnas_raw if c in row)

            for p in participantes:
                p_san = sanitize_name(p)
                clave_p = re.sub(r"\s+", " ", p_san.lower()).strip()
                if clave_p in sesion_info["vistos"]:
                    continue
                sesion_info["vistos"].add(clave_p)

                ruta_carpeta = os.path.join(ruta_sesion, p_san) if p_san != "_vacio" else ruta_sesion
                ruta_abs = os.path.join(self.ruta_activa, ruta_carpeta)
                num_archivos = contar_archivos_en_carpeta(ruta_abs)
                estado = f"CARGADA ({num_archivos})" if num_archivos > 0 else "PENDIENTE"

                filas.append({
                    "estado": estado,
                    "es_cargada": num_archivos > 0,
                    "participante": p,
                    "sesion": sesion_info["titulo"],
                    "ruta_sesion": ruta_sesion,
                    "sala": sala_val,
                    "dia": dia_val,
                    "carpeta": ruta_carpeta,
                    "busqueda": f"{p} {sesion_info['titulo']} {sala_val} {dia_val} {ruta_carpeta} {datos_fila}".lower(),
                })

        self.filas_agenda = filas
        self.combo_dia["values"] = dias_unicos
        if self.var_dia.get() not in dias_unicos:
            self.var_dia.set("Todos")

        self.combo_sala["values"] = salas_unicas
        if self.var_sala.get() not in salas_unicas:
            self.var_sala.set("Todas")

        self._actualizar_permisos_ui()
        self._refrescar_tabla()
        self._actualizar_info_red()

    def _refrescar_tabla(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        q = self.var_busqueda.get().lower().strip()
        dia_sel = self.var_dia.get()
        sala_sel = self.var_sala.get()
        estado_sel = self.var_estado.get()
        sala_bloq = self._sala_restringida()

        sesiones_visibles = set()
        ponentes_visibles = 0
        cargadas_visibles = 0

        for f in self.filas_agenda:
            if sala_bloq and f["sala"].lower() != sala_bloq:
                continue
            if sala_sel and sala_sel != "Todas" and f["sala"] != sala_sel:
                continue
            if dia_sel and dia_sel != "Todos" and f["dia"] != dia_sel:
                continue
            if estado_sel == "CARGADA" and not f["es_cargada"]:
                continue
            if estado_sel == "PENDIENTE" and f["es_cargada"]:
                continue
            if q and q not in f["busqueda"]:
                continue

            self.tree.insert(
                "",
                "end",
                values=(f["estado"], f["participante"], f["sesion"], f["sala"], f["dia"], f["carpeta"]),
            )
            sesiones_visibles.add(f["ruta_sesion"])
            ponentes_visibles += 1
            if f["es_cargada"]:
                cargadas_visibles += 1

        archivos_bk = contar_archivos_en_carpeta(self.ruta_backup)
        self.lbl_kpi_sesiones.config(text=f"Sesiones: {len(sesiones_visibles)}")
        self.lbl_kpi_ponentes.config(text=f"Ponentes: {ponentes_visibles}")
        self.lbl_kpi_cargadas.config(text=f"Cargadas: {cargadas_visibles} / {ponentes_visibles}")
        self.lbl_kpi_backup.config(text=f"Backup: {archivos_bk} arch.")

    def _limpiar_filtros(self):
        self.var_busqueda.set("")
        self.var_dia.set("Todos")
        if not self._sala_restringida():
            self.var_sala.set("Todas")
        self.var_estado.set("Todos")
        self._refrescar_tabla()

    def _generar_carpetas(self):
        if not self._es_admin():
            messagebox.showwarning("Solo lectura", "Solo los puestos ADMIN pueden generar carpetas.")
            return

        columnas_raw = list(self.df.columns)
        creadas = 0
        rutas_creadas = set()

        for _, row in self.df.iterrows():
            ruta_sesion = construir_ruta_sesion(row, self.config, columnas_raw)
            ruta_abs = os.path.join(self.ruta_activa, ruta_sesion)
            if ruta_abs not in rutas_creadas and not os.path.exists(ruta_abs):
                os.makedirs(ruta_abs, exist_ok=True)
                rutas_creadas.add(ruta_abs)
                creadas += 1

            for p in obtener_participantes_fila(row, columnas_raw, self.config):
                p_san = sanitize_name(p)
                if p_san == "_vacio":
                    continue
                p_abs = os.path.join(ruta_abs, p_san)
                if p_abs not in rutas_creadas and not os.path.exists(p_abs):
                    os.makedirs(p_abs, exist_ok=True)
                    rutas_creadas.add(p_abs)
                    creadas += 1

        copiados_bk = sincronizar_estructura_backup_local(self.ruta_activa)
        self._cargar_agenda()
        messagebox.showinfo(
            "Carpetas generadas",
            f"Carpetas listas ({creadas} nuevas). Backup sincronizado ({copiados_bk} archivos copiados).",
        )

    def _abrir_carpeta_rel(self, ruta_rel):
        ruta_abs = os.path.join(self.ruta_activa, ruta_rel)
        if self._es_admin():
            os.makedirs(ruta_abs, exist_ok=True)
        elif not os.path.exists(ruta_abs):
            messagebox.showwarning("Carpeta no existe", "La carpeta aún no existe en disco.")
            return
        abrir_carpeta_en_sistema(ruta_abs)

    def _abrir_subcarpeta_seleccionada(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Selección", "Selecciona primero una fila de la agenda.")
            return
        valores = self.tree.item(sel[0], "values")
        if len(valores) >= 6:
            self._abrir_carpeta_rel(valores[5])

    def _subir_presentacion(self):
        if not self._es_admin():
            messagebox.showwarning("Solo lectura", "Solo los puestos ADMIN pueden subir presentaciones.")
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Selección", "Selecciona primero una fila de la agenda.")
            return
        valores = self.tree.item(sel[0], "values")
        ruta_rel = valores[5]
        archivo_origen = filedialog.askopenfilename(title="Seleccionar presentación")
        if not archivo_origen:
            return

        carpeta_destino = os.path.join(self.ruta_activa, ruta_rel)
        os.makedirs(carpeta_destino, exist_ok=True)
        nombre_base, ext = os.path.splitext(os.path.basename(archivo_origen))
        nombre_seguro = sanitize_name(nombre_base) + sanitize_name(ext).replace("_vacio", "")
        destino_ppt_abs = os.path.join(carpeta_destino, nombre_seguro)
        shutil.copy2(archivo_origen, destino_ppt_abs)

        copia_bk = copiar_a_backup_local_sin_borrar(ruta_rel, nombre_seguro, destino_ppt_abs)
        crear_acceso_directo_escritorio()
        self._cargar_agenda()
        messagebox.showinfo(
            "Presentación cargada",
            f"Guardada en:\n{destino_ppt_abs}\n\nCopia de seguridad en:\n{copia_bk}",
        )

    def _actualizar_info_red(self):
        if not hasattr(self, "txt_red"):
            return
        ip_local = obtener_ip_local()
        hostname = socket.gethostname()
        wg = (self.config or {}).get("workgroup", "N/A")
        puesto = self._puesto_actual() or "Sin fijar"
        ruta_ppt = os.path.join(self.ruta_activa, "PPT")
        acceso_esc = crear_acceso_directo_escritorio() or "No disponible"
        archivos_bk = contar_archivos_en_carpeta(self.ruta_backup)

        lineas = [
            f"IP Local:            {ip_local}",
            f"Nombre Equipo:       {hostname}",
            f"Workgroup:           {wg}",
            f"Puesto Activo:       {puesto}",
            f"Carpeta Ejecutable:  {self.ruta_activa}",
            f"Carpeta PPT:         {ruta_ppt}",
            f"Backup Local Raíz:   {self.ruta_backup} ({archivos_bk} archivos — nunca borra nada)",
            f"Acceso Escritorio:   {acceso_esc}",
            "",
            "--- PUESTOS CONFIGURADOS ---",
        ]
        for eq in (self.config or {}).get("equipos", []):
            rol = "ADMIN (Carga y carpetas)" if str(eq).upper().startswith("ADM") else "SALA (Solo lectura)"
            estado = "Bloqueado aquí" if puesto.upper() == str(eq).upper() else "Disponible"
            lineas.append(f"  * {eq:<24} | {rol:<24} | {estado}")

        self.txt_red.delete("1.0", "end")
        self.txt_red.insert("1.0", "\n".join(lineas))

    def _recargar_todo(self):
        self.config, self.df, self.ruta_activa = cargar_datos_congreso()
        for w in self.root.winfo_children():
            w.destroy()
        self._construir_interfaz()
        self._cargar_agenda()

    def _al_cerrar(self):
        limpiar_bloqueo_al_salir(self.ruta_bloqueo)
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = CongresoAppEscritorio(root)
    root.mainloop()
