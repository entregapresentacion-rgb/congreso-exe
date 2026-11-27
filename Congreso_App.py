import os
import re
import sys
import json
import base64
import traceback
import platform
import threading
import subprocess
import webbrowser
import datetime as dt
import pandas as pd
from flask import Flask, render_template_string, request, jsonify

SISTEMA = platform.system()

INVALID_CHARS = set('<>:"/\\|?*') | {chr(i) for i in range(32)}
RESERVED = (
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)
MAX_NAME_BYTES = 240


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


def norm_col(nombre):
    s = str(nombre).replace("\xa0", " ").lower().strip()
    for a, b in [("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ñ", "n")]:
        s = s.replace(a, b)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


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


def ordenar_columnas_por_nivel_3(columnas, n3_cols):
    ordenadas = []
    usadas = set()

    for c in n3_cols:
        if c in columnas and c not in usadas:
            ordenadas.append(c)
            usadas.add(c)

    col_id = buscar_columna([c for c in columnas if c not in usadas], ["id"])
    if col_id:
        ordenadas.append(col_id)
        usadas.add(col_id)

    for c in columnas:
        if c not in usadas:
            ordenadas.append(c)
            usadas.add(c)

    return ordenadas


app = Flask(__name__)

HTML_ERROR = base64.b64decode("PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVzIj4KPGhlYWQ+PG1ldGEgY2hhcnNldD0iVVRGLTgiPjx0aXRsZT5FcnJvciBlbiBDb25ncmVzb19BcHA8L3RpdGxlPjwvaGVhZD4KPGJvZHkgc3R5bGU9ImJhY2tncm91bmQ6IzExMTcyNjtjb2xvcjojZmZmO2ZvbnQtZmFtaWx5OnNhbnMtc2VyaWY7cGFkZGluZzozMHB4OyI+CiAgPGgyIHN0eWxlPSJjb2xvcjojZjg3MTcxOyI+U2UgcHJvZHVqbyB1biBlcnJvciBhbCBjYXJnYXIgQ29uZ3Jlc29fQXBwPC9oMj4KICA8cD5SdXRhIGFjdGl2YSBkZXRlY3RhZGE6IDxjb2RlPnt7IHJ1dGFfYWN0aXZhIH19PC9jb2RlPjwvcD4KICA8cHJlIHN0eWxlPSJiYWNrZ3JvdW5kOiMxZTI5M2I7cGFkZGluZzoxNnB4O2JvcmRlci1yYWRpdXM6OHB4O292ZXJmbG93OmF1dG87Y29sb3I6I2ZjYTVhNTsiPnt7IGRldGFsbGVfZXJyb3IgfX08L3ByZT4KPC9ib2R5Pgo8L2h0bWw+").decode("utf-8")
HTML_APP = base64.b64decode("PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVzIj4KPGhlYWQ+CiAgICA8bWV0YSBjaGFyc2V0PSJVVEYtOCI+CiAgICA8bWV0YSBuYW1lPSJ2aWV3cG9ydCIgY29udGVudD0id2lkdGg9ZGV2aWNlLXdpZHRoLCBpbml0aWFsLXNjYWxlPTEuMCI+CiAgICA8dGl0bGU+Q29uZ3Jlc29BcHAgfCBBZ2VuZGEgcG9yIFNlc2lvbmVzPC90aXRsZT4KICAgIDxsaW5rIGhyZWY9Imh0dHBzOi8vY2RuLmpzZGVsaXZyLm5ldC9ucG0vYm9vdHN0cmFwQDUuMy4yL2Rpc3QvY3NzL2Jvb3RzdHJhcC5taW4uY3NzIiByZWw9InN0eWxlc2hlZXQiPgogICAgPGxpbmsgaHJlZj0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9ib290c3RyYXAtaWNvbnNAMS4xMS4zL2ZvbnQvYm9vdHN0cmFwLWljb25zLm1pbi5jc3MiIHJlbD0ic3R5bGVzaGVldCI+CiAgICA8bGluayBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tL2NzczI/ZmFtaWx5PVJvYm90bzp3Z2h0QDMwMDs0MDA7NTAwOzcwMCZkaXNwbGF5PXN3YXAiIHJlbD0ic3R5bGVzaGVldCI+CiAgICA8c3R5bGU+CiAgICAgICAgOnJvb3QgewogICAgICAgICAgICAtLWJnLWJvZHk6ICMxMTE3MjY7CiAgICAgICAgICAgIC0tYmctY2FyZDogIzFmMjk0MDsKICAgICAgICAgICAgLS1iZy1pbnB1dDogIzE1MWQzMDsKICAgICAgICAgICAgLS1ib3JkZXItY29sb3I6IHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4wOCk7CiAgICAgICAgICAgIC0tdGV4dC1tYWluOiAjZmZmZmZmOwogICAgICAgICAgICAtLXRleHQtbXV0ZWQ6ICM5NGEzYjg7CiAgICAgICAgfQogICAgICAgIGJvZHkgewogICAgICAgICAgICBiYWNrZ3JvdW5kLWNvbG9yOiB2YXIoLS1iZy1ib2R5KTsKICAgICAgICAgICAgY29sb3I6IHZhcigtLXRleHQtbWFpbik7CiAgICAgICAgICAgIGZvbnQtZmFtaWx5OiAnUm9ib3RvJywgLWFwcGxlLXN5c3RlbSwgQmxpbmtNYWNTeXN0ZW1Gb250LCBzYW5zLXNlcmlmOwogICAgICAgICAgICBtYXJnaW46IDA7CiAgICAgICAgICAgIHBhZGRpbmctYm90dG9tOiA1MHB4OwogICAgICAgIH0KICAgICAgICAudmFsZXgtdG9wYmFyIHsKICAgICAgICAgICAgYmFja2dyb3VuZC1jb2xvcjogdmFyKC0tYmctY2FyZCk7CiAgICAgICAgICAgIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXItY29sb3IpOwogICAgICAgICAgICBwYWRkaW5nOiAxMnB4IDI0cHg7CiAgICAgICAgfQogICAgICAgIC5icmFuZC1sb2dvIHsKICAgICAgICAgICAgZGlzcGxheTogZmxleDsKICAgICAgICAgICAgYWxpZ24taXRlbXM6IGNlbnRlcjsKICAgICAgICAgICAgZ2FwOiAxMHB4OwogICAgICAgICAgICBmb250LXdlaWdodDogNzAwOwogICAgICAgICAgICBmb250LXNpemU6IDEuMjVyZW07CiAgICAgICAgICAgIGNvbG9yOiAjZmZmOwogICAgICAgICAgICB0ZXh0LWRlY29yYXRpb246IG5vbmU7CiAgICAgICAgfQogICAgICAgIC5icmFuZC1pY29uIHsKICAgICAgICAgICAgd2lkdGg6IDMycHg7CiAgICAgICAgICAgIGhlaWdodDogMzJweDsKICAgICAgICAgICAgYm9yZGVyLXJhZGl1czogOHB4OwogICAgICAgICAgICBiYWNrZ3JvdW5kOiBsaW5lYXItZ3JhZGllbnQoMTM1ZGVnLCAjMDE2MmU4LCAjMDBiOWZmKTsKICAgICAgICAgICAgZGlzcGxheTogZmxleDsKICAgICAgICAgICAgYWxpZ24taXRlbXM6IGNlbnRlcjsKICAgICAgICAgICAganVzdGlmeS1jb250ZW50OiBjZW50ZXI7CiAgICAgICAgfQogICAgICAgIC52YWxleC1zZWFyY2ggewogICAgICAgICAgICBiYWNrZ3JvdW5kLWNvbG9yOiB2YXIoLS1iZy1pbnB1dCk7CiAgICAgICAgICAgIGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoNTYsIDE4OSwgMjQ4LCAwLjI1KTsKICAgICAgICAgICAgYm9yZGVyLXJhZGl1czogMjBweDsKICAgICAgICAgICAgcGFkZGluZzogN3B4IDE2cHg7CiAgICAgICAgICAgIGNvbG9yOiAjZmZmOwogICAgICAgICAgICBmb250LXNpemU6IDAuODhyZW07CiAgICAgICAgICAgIHdpZHRoOiAzNjBweDsKICAgICAgICB9CiAgICAgICAgLnZhbGV4LXNlYXJjaDpmb2N1cyB7IG91dGxpbmU6IG5vbmU7IGJvcmRlci1jb2xvcjogIzAwYjlmZjsgYm94LXNoYWRvdzogMCAwIDAgMnB4IHJnYmEoMCwgMTg1LCAyNTUsIDAuMTgpOyB9CiAgICAgICAgLmtwaS1jYXJkIHsKICAgICAgICAgICAgYm9yZGVyLXJhZGl1czogOHB4OwogICAgICAgICAgICBwYWRkaW5nOiAxOHB4IDIwcHggMzZweCAyMHB4OwogICAgICAgICAgICBwb3NpdGlvbjogcmVsYXRpdmU7CiAgICAgICAgICAgIG92ZXJmbG93OiBoaWRkZW47CiAgICAgICAgICAgIGNvbG9yOiAjZmZmOwogICAgICAgICAgICBib3gtc2hhZG93OiAwIDRweCAxNXB4IHJnYmEoMCwgMCwgMCwgMC4yNSk7CiAgICAgICAgICAgIGhlaWdodDogMTAwJTsKICAgICAgICB9CiAgICAgICAgLmtwaS1ibHVlIHsgYmFja2dyb3VuZDogbGluZWFyLWdyYWRpZW50KHRvIHJpZ2h0LCAjMDE2MmU4IDAlLCAjMDBiOWZmIDEwMCUpOyB9CiAgICAgICAgLmtwaS1waW5rIHsgYmFja2dyb3VuZDogbGluZWFyLWdyYWRpZW50KHRvIHJpZ2h0LCAjZjkzYTVhIDAlLCAjZjc3NzhjIDEwMCUpOyB9CiAgICAgICAgLmtwaS1ncmVlbiB7IGJhY2tncm91bmQ6IGxpbmVhci1ncmFkaWVudCh0byByaWdodCwgIzAyOTY2NiAwJSwgIzQ4ZDZhOCAxMDAlKTsgfQogICAgICAgIC5rcGktb3JhbmdlIHsgYmFja2dyb3VuZDogbGluZWFyLWdyYWRpZW50KHRvIHJpZ2h0LCAjZjc2YTJkIDAlLCAjZWZhNjVmIDEwMCUpOyB9CiAgICAgICAgLmtwaS10aXRsZSB7IGZvbnQtc2l6ZTogMC43MnJlbTsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IDAuOHB4OyBmb250LXdlaWdodDogNzAwOyBvcGFjaXR5OiAwLjk7IG1hcmdpbi1ib3R0b206IDZweDsgfQogICAgICAgIC5rcGktdmFsdWUgeyBmb250LXNpemU6IDEuNTVyZW07IGZvbnQtd2VpZ2h0OiA3MDA7IGxpbmUtaGVpZ2h0OiAxLjI7IH0KICAgICAgICAua3BpLXN1YiB7IGZvbnQtc2l6ZTogMC43OHJlbTsgb3BhY2l0eTogMC44NTsgbWFyZ2luLXRvcDogNHB4OyB9CiAgICAgICAgLmtwaS13YXZlIHsgcG9zaXRpb246IGFic29sdXRlOyBib3R0b206IDA7IGxlZnQ6IDA7IHdpZHRoOiAxMDAlOyBoZWlnaHQ6IDMycHg7IG9wYWNpdHk6IDAuMzU7IH0KICAgICAgICAudmFsZXgtY2FyZCB7CiAgICAgICAgICAgIGJhY2tncm91bmQtY29sb3I6IHZhcigtLWJnLWNhcmQpOwogICAgICAgICAgICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXItY29sb3IpOwogICAgICAgICAgICBib3JkZXItcmFkaXVzOiA4cHg7CiAgICAgICAgICAgIGJveC1zaGFkb3c6IDAgNHB4IDIwcHggcmdiYSgwLCAwLCAwLCAwLjIpOwogICAgICAgIH0KICAgICAgICAudmFsZXgtY2FyZC1oZWFkZXIgewogICAgICAgICAgICBwYWRkaW5nOiAxNnB4IDIycHg7CiAgICAgICAgICAgIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXItY29sb3IpOwogICAgICAgICAgICBkaXNwbGF5OiBmbGV4OwogICAgICAgICAgICBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47CiAgICAgICAgICAgIGFsaWduLWl0ZW1zOiBjZW50ZXI7CiAgICAgICAgICAgIGZsZXgtd3JhcDogd3JhcDsKICAgICAgICAgICAgZ2FwOiAxMHB4OwogICAgICAgIH0KICAgICAgICAudmFsZXgtY2FyZC10aXRsZSB7IGZvbnQtc2l6ZTogMC45NXJlbTsgZm9udC13ZWlnaHQ6IDcwMDsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IDAuNXB4OyBtYXJnaW46IDA7IH0KICAgICAgICAudmFsZXgtY2FyZC1zdWJ0aXRsZSB7IGZvbnQtc2l6ZTogMC44cmVtOyBjb2xvcjogdmFyKC0tdGV4dC1tdXRlZCk7IG1hcmdpbi10b3A6IDJweDsgfQogICAgICAgIC5mb3JtLXNlbGVjdCwgLmZvcm0tY29udHJvbCB7CiAgICAgICAgICAgIGJhY2tncm91bmQtY29sb3I6IHZhcigtLWJnLWlucHV0KSAhaW1wb3J0YW50OwogICAgICAgICAgICBib3JkZXI6IDFweCBzb2xpZCByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMTQpICFpbXBvcnRhbnQ7CiAgICAgICAgICAgIGNvbG9yOiAjZmZmICFpbXBvcnRhbnQ7CiAgICAgICAgICAgIGZvbnQtc2l6ZTogMC44MnJlbTsKICAgICAgICB9CiAgICAgICAgLmJ0bi12YWxleC1ibHVlIHsgYmFja2dyb3VuZDogbGluZWFyLWdyYWRpZW50KHRvIHJpZ2h0LCAjMDE2MmU4LCAjMDBiOWZmKTsgYm9yZGVyOiBub25lOyBjb2xvcjogI2ZmZjsgZm9udC13ZWlnaHQ6IDYwMDsgfQogICAgICAgIC5idG4tdmFsZXgtZ3JlZW4geyBiYWNrZ3JvdW5kOiBsaW5lYXItZ3JhZGllbnQodG8gcmlnaHQsICMwMjk2NjYsICM0OGQ2YTgpOyBib3JkZXI6IG5vbmU7IGNvbG9yOiAjZmZmOyBmb250LXdlaWdodDogNjAwOyB9CiAgICAgICAgLmJ0bi12YWxleC1ibHVlOmhvdmVyLCAuYnRuLXZhbGV4LWdyZWVuOmhvdmVyIHsgY29sb3I6ICNmZmY7IG9wYWNpdHk6IDAuOTI7IH0KCiAgICAgICAgLnNlc2lvbi1ibG9xdWUgewogICAgICAgICAgICBib3JkZXItcmFkaXVzOiAxMHB4OwogICAgICAgICAgICBtYXJnaW4tYm90dG9tOiAxOHB4OwogICAgICAgICAgICBvdmVyZmxvdzogaGlkZGVuOwogICAgICAgICAgICBib3gtc2hhZG93OiAwIDZweCAxOHB4IHJnYmEoMCwgMCwgMCwgMC4yOCk7CiAgICAgICAgfQogICAgICAgIC5zZXNpb24tYmxvcXVlLnRlbWEtMCB7CiAgICAgICAgICAgIGJhY2tncm91bmQtY29sb3I6ICMxOTI2NDI7CiAgICAgICAgICAgIGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoNTYsIDE4OSwgMjQ4LCAwLjM1KTsKICAgICAgICAgICAgYm9yZGVyLWxlZnQ6IDVweCBzb2xpZCAjMDBiOWZmOwogICAgICAgIH0KICAgICAgICAuc2VzaW9uLWJsb3F1ZS50ZW1hLTAgLnNlc2lvbi1oZWFkZXIgewogICAgICAgICAgICBiYWNrZ3JvdW5kOiBsaW5lYXItZ3JhZGllbnQoOTBkZWcsIHJnYmEoMSwgOTgsIDIzMiwgMC4zNCksIHJnYmEoMCwgMTg1LCAyNTUsIDAuMTIpKTsKICAgICAgICB9CiAgICAgICAgLnNlc2lvbi1ibG9xdWUudGVtYS0xIHsKICAgICAgICAgICAgYmFja2dyb3VuZC1jb2xvcjogIzIxMWUzZDsKICAgICAgICAgICAgYm9yZGVyOiAxcHggc29saWQgcmdiYSgxNjgsIDg1LCAyNDcsIDAuMzUpOwogICAgICAgICAgICBib3JkZXItbGVmdDogNXB4IHNvbGlkICNhODU1Zjc7CiAgICAgICAgfQogICAgICAgIC5zZXNpb24tYmxvcXVlLnRlbWEtMSAuc2VzaW9uLWhlYWRlciB7CiAgICAgICAgICAgIGJhY2tncm91bmQ6IGxpbmVhci1ncmFkaWVudCg5MGRlZywgcmdiYSgxMjYsIDM0LCAyMDYsIDAuMzQpLCByZ2JhKDE2OCwgODUsIDI0NywgMC4xMikpOwogICAgICAgIH0KICAgICAgICAuc2VzaW9uLWJsb3F1ZS50ZW1hLTIgewogICAgICAgICAgICBiYWNrZ3JvdW5kLWNvbG9yOiAjMTYyZTM2OwogICAgICAgICAgICBib3JkZXI6IDFweCBzb2xpZCByZ2JhKDQ1LCAyMTIsIDE5MSwgMC4zNSk7CiAgICAgICAgICAgIGJvcmRlci1sZWZ0OiA1cHggc29saWQgIzJkZDRiZjsKICAgICAgICB9CiAgICAgICAgLnNlc2lvbi1ibG9xdWUudGVtYS0yIC5zZXNpb24taGVhZGVyIHsKICAgICAgICAgICAgYmFja2dyb3VuZDogbGluZWFyLWdyYWRpZW50KDkwZGVnLCByZ2JhKDEzLCAxNDgsIDEzNiwgMC4zNCksIHJnYmEoNDUsIDIxMiwgMTkxLCAwLjEyKSk7CiAgICAgICAgfQogICAgICAgIC5zZXNpb24tYmxvcXVlLnRlbWEtMyB7CiAgICAgICAgICAgIGJhY2tncm91bmQtY29sb3I6ICMyZDIzMzY7CiAgICAgICAgICAgIGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoMjQ0LCAxMTQsIDE4MiwgMC4zNSk7CiAgICAgICAgICAgIGJvcmRlci1sZWZ0OiA1cHggc29saWQgI2Y0NzJiNjsKICAgICAgICB9CiAgICAgICAgLnNlc2lvbi1ibG9xdWUudGVtYS0zIC5zZXNpb24taGVhZGVyIHsKICAgICAgICAgICAgYmFja2dyb3VuZDogbGluZWFyLWdyYWRpZW50KDkwZGVnLCByZ2JhKDE5MCwgMjQsIDkzLCAwLjMyKSwgcmdiYSgyNDQsIDExNCwgMTgyLCAwLjEyKSk7CiAgICAgICAgfQogICAgICAgIC5zZXNpb24taGVhZGVyIHsKICAgICAgICAgICAgcGFkZGluZzogMTRweCAyMHB4OwogICAgICAgICAgICBib3JkZXItYm90dG9tOiAxcHggc29saWQgcmdiYSgyNTUsIDI1NSwgMjU1LCAwLjA4KTsKICAgICAgICAgICAgZGlzcGxheTogZmxleDsKICAgICAgICAgICAganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOwogICAgICAgICAgICBhbGlnbi1pdGVtczogY2VudGVyOwogICAgICAgICAgICBmbGV4LXdyYXA6IHdyYXA7CiAgICAgICAgICAgIGdhcDogMTJweDsKICAgICAgICB9CiAgICAgICAgLnNlc2lvbi1tZXRhLWNoaXAgewogICAgICAgICAgICBkaXNwbGF5OiBpbmxpbmUtZmxleDsKICAgICAgICAgICAgYWxpZ24taXRlbXM6IGNlbnRlcjsKICAgICAgICAgICAgZ2FwOiA1cHg7CiAgICAgICAgICAgIGJhY2tncm91bmQ6IHJnYmEoMTUsIDIzLCA0MiwgMC41NSk7CiAgICAgICAgICAgIGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoMjU1LCAyNTUsIDI1NSwgMC4xMik7CiAgICAgICAgICAgIGJvcmRlci1yYWRpdXM6IDZweDsKICAgICAgICAgICAgcGFkZGluZzogM3B4IDlweDsKICAgICAgICAgICAgZm9udC1zaXplOiAwLjc4cmVtOwogICAgICAgICAgICBjb2xvcjogI2UyZThmMDsKICAgICAgICB9CiAgICAgICAgLnNlc2lvbi1ib2R5IHsKICAgICAgICAgICAgcGFkZGluZzogMTRweCAyMHB4OwogICAgICAgICAgICBkaXNwbGF5OiBmbGV4OwogICAgICAgICAgICBmbGV4LWRpcmVjdGlvbjogY29sdW1uOwogICAgICAgICAgICBnYXA6IDEwcHg7CiAgICAgICAgfQogICAgICAgIC5wb25lbnRlLWNhcmQgewogICAgICAgICAgICBiYWNrZ3JvdW5kLWNvbG9yOiByZ2JhKDE1LCAyMywgNDIsIDAuNTIpOwogICAgICAgICAgICBib3JkZXI6IDFweCBzb2xpZCByZ2JhKDI1NSwgMjU1LCAyNTUsIDAuMDkpOwogICAgICAgICAgICBib3JkZXItcmFkaXVzOiA4cHg7CiAgICAgICAgICAgIHBhZGRpbmc6IDExcHggMTZweDsKICAgICAgICAgICAgZGlzcGxheTogZmxleDsKICAgICAgICAgICAganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOwogICAgICAgICAgICBhbGlnbi1pdGVtczogY2VudGVyOwogICAgICAgICAgICBmbGV4LXdyYXA6IHdyYXA7CiAgICAgICAgICAgIGdhcDogMTJweDsKICAgICAgICB9CiAgICAgICAgLnBvbmVudGUtY2FyZDpob3ZlciB7CiAgICAgICAgICAgIGJhY2tncm91bmQtY29sb3I6IHJnYmEoMTUsIDIzLCA0MiwgMC43OCk7CiAgICAgICAgICAgIGJvcmRlci1jb2xvcjogcmdiYSg1NiwgMTg5LCAyNDgsIDAuMzUpOwogICAgICAgIH0KICAgICAgICAuY2FydGVsLWVzdGFkbyB7CiAgICAgICAgICAgIGZvbnQtc2l6ZTogMC43OHJlbTsKICAgICAgICAgICAgZm9udC13ZWlnaHQ6IDcwMDsKICAgICAgICAgICAgbGV0dGVyLXNwYWNpbmc6IDAuNXB4OwogICAgICAgICAgICB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOwogICAgICAgICAgICBwYWRkaW5nOiA1cHggMTJweDsKICAgICAgICAgICAgYm9yZGVyLXJhZGl1czogNnB4OwogICAgICAgICAgICBkaXNwbGF5OiBpbmxpbmUtZmxleDsKICAgICAgICAgICAgYWxpZ24taXRlbXM6IGNlbnRlcjsKICAgICAgICAgICAgZ2FwOiA2cHg7CiAgICAgICAgfQogICAgICAgIC5jYXJ0ZWwtY2FyZ2FkYSB7CiAgICAgICAgICAgIGJhY2tncm91bmQ6IHJnYmEoMTYsIDE4NSwgMTI5LCAwLjIyKTsKICAgICAgICAgICAgY29sb3I6ICM2ZWU3Yjc7CiAgICAgICAgICAgIGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoNTIsIDIxMSwgMTUzLCAwLjUpOwogICAgICAgIH0KICAgICAgICAuY2FydGVsLXBlbmRpZW50ZSB7CiAgICAgICAgICAgIGJhY2tncm91bmQ6IHJnYmEoMjQ1LCAxNTgsIDExLCAwLjIyKTsKICAgICAgICAgICAgY29sb3I6ICNmY2QzNGQ7CiAgICAgICAgICAgIGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoMjUxLCAxOTEsIDM2LCAwLjUpOwogICAgICAgIH0KICAgIDwvc3R5bGU+CjwvaGVhZD4KPGJvZHk+CgogICAgPGhlYWRlciBjbGFzcz0idmFsZXgtdG9wYmFyIGQtZmxleCBmbGV4LXdyYXAganVzdGlmeS1jb250ZW50LWJldHdlZW4gYWxpZ24taXRlbXMtY2VudGVyIGdhcC0zIj4KICAgICAgICA8ZGl2IGNsYXNzPSJkLWZsZXggYWxpZ24taXRlbXMtY2VudGVyIGdhcC0zIj4KICAgICAgICAgICAgPGEgaHJlZj0iLyIgY2xhc3M9ImJyYW5kLWxvZ28iPgogICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0iYnJhbmQtaWNvbiI+PGkgY2xhc3M9ImJpIGJpLWVhc2VsMi1maWxsIj48L2k+PC9kaXY+CiAgICAgICAgICAgICAgICA8c3Bhbj5Db25ncmVzbzxzcGFuIHN0eWxlPSJjb2xvcjojMDBiOWZmOyI+QXBwPC9zcGFuPjwvc3Bhbj4KICAgICAgICAgICAgPC9hPgogICAgICAgICAgICA8ZGl2IGNsYXNzPSJkLWZsZXggYWxpZ24taXRlbXMtY2VudGVyIGdhcC0yIj4KICAgICAgICAgICAgICAgIDxpbnB1dCB0eXBlPSJ0ZXh0IiBpZD0iYnVzY2Fkb3JHbG9iYWwiIGNsYXNzPSJ2YWxleC1zZWFyY2giIHBsYWNlaG9sZGVyPSJGaWx0cmFyIHBvciBwb25lbnRlLCBzZXNpw7NuLCBzYWxhIG8gZMOtYS4uLiIgb25rZXl1cD0iZmlsdHJhckFnZW5kYSgpIiBvbmlucHV0PSJmaWx0cmFyQWdlbmRhKCkiPgogICAgICAgICAgICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGNsYXNzPSJidG4gYnRuLXNtIGJ0bi1vdXRsaW5lLWluZm8gcHgtMiBweS0xIiBvbmNsaWNrPSJsaW1waWFyRmlsdHJvcygpIiB0aXRsZT0iTGltcGlhciBmaWx0cm8iPgogICAgICAgICAgICAgICAgICAgIDxpIGNsYXNzPSJiaSBiaS14LWxnIj48L2k+CiAgICAgICAgICAgICAgICA8L2J1dHRvbj4KICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0iZC1mbGV4IGFsaWduLWl0ZW1zLWNlbnRlciBnYXAtMiI+CiAgICAgICAgICAgIDxzcGFuIGNsYXNzPSJiYWRnZSBiZy1kYXJrIGJvcmRlciBib3JkZXItc2Vjb25kYXJ5IHB4LTMgcHktMiI+CiAgICAgICAgICAgICAgICA8aSBjbGFzcz0iYmkgYmktaGRkLW5ldHdvcmsgdGV4dC1pbmZvIG1lLTEiPjwvaT4ge3sgY29uZmlnLndvcmtncm91cCBpZiBjb25maWcgZWxzZSAnTi9BJyB9fQogICAgICAgICAgICA8L3NwYW4+CiAgICAgICAgICAgIDxzcGFuIGNsYXNzPSJiYWRnZSBiZy1kYXJrIGJvcmRlciBib3JkZXItc2Vjb25kYXJ5IHB4LTMgcHktMiI+CiAgICAgICAgICAgICAgICA8aSBjbGFzcz0iYmkgYmktc2hpZWxkLWxvY2sgdGV4dC13YXJuaW5nIG1lLTEiPjwvaT4KICAgICAgICAgICAgICAgIHslIGlmIHB1ZXN0b19hY3Rpdm8gJX0KICAgICAgICAgICAgICAgICAgICB7eyBwdWVzdG9fYWN0aXZvIH19ICh7eyAnQURNSU4nIGlmIGVzX2FkbWluIGVsc2UgJ1NBTEEgLSBTT0xPIExFQ1RVUkEnIH19KQogICAgICAgICAgICAgICAgeyUgZWxzZSAlfQogICAgICAgICAgICAgICAgICAgIFNpbiBwdWVzdG8gZmlqYWRvCiAgICAgICAgICAgICAgICB7JSBlbmRpZiAlfQogICAgICAgICAgICA8L3NwYW4+CiAgICAgICAgPC9kaXY+CiAgICA8L2hlYWRlcj4KCiAgICA8ZGl2IGNsYXNzPSJjb250YWluZXItZmx1aWQgcHgtNCBwdC00IiBzdHlsZT0ibWF4LXdpZHRoOiAxNDQwcHg7Ij4KCiAgICAgICAgeyUgaWYgbm90IGNvbmZpZyAlfQogICAgICAgICAgICA8ZGl2IGNsYXNzPSJ2YWxleC1jYXJkIHAtNSB0ZXh0LWNlbnRlciBteS01Ij4KICAgICAgICAgICAgICAgIDxpIGNsYXNzPSJiaSBiaS1leGNsYW1hdGlvbi1kaWFtb25kIHRleHQtd2FybmluZyBkaXNwbGF5LTQgbWItMyI+PC9pPgogICAgICAgICAgICAgICAgPGg0IGNsYXNzPSJmdy1ib2xkIj5ObyBzZSBlbmNvbnRyYXJvbiBsb3MgYXJjaGl2b3MgZGUgY29uZmlndXJhY2nDs248L2g0PgogICAgICAgICAgICAgICAgPHAgc3R5bGU9ImNvbG9yOiB2YXIoLS10ZXh0LW11dGVkKTsgbWF4LXdpZHRoOiA1NjBweDsgbWFyZ2luOiAxMHB4IGF1dG8gMjBweDsiPgogICAgICAgICAgICAgICAgICAgIENvbG9jYSA8Y29kZT5wcm9ncmFtYS54bHN4PC9jb2RlPiB5IDxjb2RlPmNvbmZpZ19jb25ncmVzby5qc29uPC9jb2RlPiBlbiBsYSBtaXNtYSBjYXJwZXRhIHF1ZSA8c3Ryb25nPkNvbmdyZXNvX0FwcDwvc3Ryb25nPi4KICAgICAgICAgICAgICAgIDwvcD4KICAgICAgICAgICAgICAgIDxhIGhyZWY9Ii8iIGNsYXNzPSJidG4gYnRuLXZhbGV4LWJsdWUgcHgtNCI+UmVpbnRlbnRhcjwvYT4KICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgeyUgZWxzZSAlfQoKICAgICAgICA8ZGl2IGNsYXNzPSJ2YWxleC1jYXJkIG1iLTQiPgogICAgICAgICAgICA8ZGl2IGNsYXNzPSJ2YWxleC1jYXJkLWhlYWRlciI+CiAgICAgICAgICAgICAgICA8ZGl2PgogICAgICAgICAgICAgICAgICAgIDxoNSBjbGFzcz0idmFsZXgtY2FyZC10aXRsZSI+UHVlc3RvIGRlIFRyYWJham8gZW4gZWwgQ29uZ3Jlc288L2g1PgogICAgICAgICAgICAgICAgICAgIDxkaXYgY2xhc3M9InZhbGV4LWNhcmQtc3VidGl0bGUiPgogICAgICAgICAgICAgICAgICAgICAgICB7JSBpZiBwdWVzdG9fYmxvcXVlYWRvICV9CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBQdWVzdG8gZmlqYWRvIHkgYmxvcXVlYWRvIGVuIDxzdHJvbmc+e3sgcHVlc3RvX2Jsb3F1ZWFkbyB9fTwvc3Ryb25nPi4gUGFyYSBjYW1iaWFybG8sIHZ1ZWx2ZSBhIGFicmlyIGVsIHByb2dyYW1hLgogICAgICAgICAgICAgICAgICAgICAgICB7JSBlbHNlICV9CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBTZWxlY2Npb25hIDxzdHJvbmc+QURNMS9BRE0yPC9zdHJvbmc+IChwdWVkZSBjYXJnYXIgcHJlc2VudGFjaW9uZXMpIG8gdW5hIDxzdHJvbmc+U0FMQTwvc3Ryb25nPiAoc29sbyBsZWN0dXJhIHkgcXVlZGEgYmxvcXVlYWRhKS4KICAgICAgICAgICAgICAgICAgICAgICAgeyUgZW5kaWYgJX0KICAgICAgICAgICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0iZC1mbGV4IGFsaWduLWl0ZW1zLWNlbnRlciBnYXAtMiI+CiAgICAgICAgICAgICAgICAgICAgPHNlbGVjdCBpZD0ic2VsZWN0UHVlc3RvIiBjbGFzcz0iZm9ybS1zZWxlY3QgZm9ybS1zZWxlY3Qtc20iIHN0eWxlPSJ3aWR0aDogMjYwcHg7IiB7JSBpZiBwdWVzdG9fYmxvcXVlYWRvICV9ZGlzYWJsZWR7JSBlbmRpZiAlfT4KICAgICAgICAgICAgICAgICAgICAgICAgPG9wdGlvbiB2YWx1ZT0iIj4tLSBTZWxlY2Npb25hciBwdWVzdG8gLS08L29wdGlvbj4KICAgICAgICAgICAgICAgICAgICAgICAgeyUgZm9yIGVxIGluIGNvbmZpZy5lcXVpcG9zICV9CiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8b3B0aW9uIHZhbHVlPSJ7eyBlcSB9fSIgeyUgaWYgZXEgPT0gcHVlc3RvX2FjdGl2byAlfXNlbGVjdGVkeyUgZW5kaWYgJX0+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAge3sgZXEgfX0ge3sgJyhBRE1JTiAtIENhcmdhKScgaWYgZXEuc3RhcnRzd2l0aCgnQURNJykgZWxzZSAnKFNBTEEgLSBTb2xvIHZlciknIH19CiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8L29wdGlvbj4KICAgICAgICAgICAgICAgICAgICAgICAgeyUgZW5kZm9yICV9CiAgICAgICAgICAgICAgICAgICAgPC9zZWxlY3Q+CiAgICAgICAgICAgICAgICAgICAgeyUgaWYgbm90IHB1ZXN0b19ibG9xdWVhZG8gJX0KICAgICAgICAgICAgICAgICAgICAgICAgPGJ1dHRvbiBjbGFzcz0iYnRuIGJ0bi1zbSBidG4tdmFsZXgtZ3JlZW4gcHgtMyIgb25jbGljaz0iZmlqYXJQdWVzdG8oKSI+Q29uZmlybWFyIHkgQmxvcXVlYXI8L2J1dHRvbj4KICAgICAgICAgICAgICAgICAgICB7JSBlbHNlICV9CiAgICAgICAgICAgICAgICAgICAgICAgIDxzcGFuIGNsYXNzPSJiYWRnZSBiZy1zZWNvbmRhcnkgcHgtMyBweS0yIj48aSBjbGFzcz0iYmkgYmktbG9jay1maWxsIG1lLTEiPjwvaT4gQmxvcXVlYWRvPC9zcGFuPgogICAgICAgICAgICAgICAgICAgIHslIGVuZGlmICV9CiAgICAgICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgPC9kaXY+CgogICAgICAgIDxkaXYgY2xhc3M9InJvdyBnLTMgbWItNCI+CiAgICAgICAgICAgIDxkaXYgY2xhc3M9ImNvbC1zbS02IGNvbC14bC0zIj4KICAgICAgICAgICAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGtwaS1ibHVlIj4KICAgICAgICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJrcGktdGl0bGUiPlNlc2lvbmVzIFZpc2libGVzPC9kaXY+CiAgICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ia3BpU2VzaW9uZXMiPnt7IHNlc2lvbmVzX2FncnVwYWRhc3xsZW5ndGggfX08L2Rpdj4KICAgICAgICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJrcGktc3ViIj5Ub3RhbCBzZXNpb25lczoge3sgc2VzaW9uZXNfYWdydXBhZGFzfGxlbmd0aCB9fTwvZGl2PgogICAgICAgICAgICAgICAgICAgIDxzdmcgY2xhc3M9ImtwaS13YXZlIiB2aWV3Qm94PSIwIDAgMjAwIDQwIiBwcmVzZXJ2ZUFzcGVjdFJhdGlvPSJub25lIj48cGF0aCBkPSJNMCwzMCBMMzAsMjIgTDYwLDI4IEw5NSwxMiBMMTMwLDI0IEwxNjUsMTAgTDIwMCwxOCBMMjAwLDQwIEwwLDQwIFoiIGZpbGw9IiNmZmZmZmYiLz48L3N2Zz4KICAgICAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPGRpdiBjbGFzcz0iY29sLXNtLTYgY29sLXhsLTMiPgogICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQga3BpLXBpbmsiPgogICAgICAgICAgICAgICAgICAgIDxkaXYgY2xhc3M9ImtwaS10aXRsZSI+UG9uZW50ZXMgVmlzaWJsZXM8L2Rpdj4KICAgICAgICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJrcGlQb25lbnRlcyI+e3sgdG90YWxfcG9uZW50ZXMgfX08L2Rpdj4KICAgICAgICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJrcGktc3ViIj5Ub3RhbCBlbiBwcm9ncmFtYToge3sgdG90YWxfcG9uZW50ZXMgfX08L2Rpdj4KICAgICAgICAgICAgICAgICAgICA8c3ZnIGNsYXNzPSJrcGktd2F2ZSIgdmlld0JveD0iMCAwIDIwMCA0MCIgcHJlc2VydmVBc3BlY3RSYXRpbz0ibm9uZSI+PHBhdGggZD0iTTAsMjUgTDQwLDMyIEw4MCwxNSBMMTIwLDI1IEwxNjAsMTIgTDIwMCwyMiBMMjAwLDQwIEwwLDQwIFoiIGZpbGw9IiNmZmZmZmYiLz48L3N2Zz4KICAgICAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPGRpdiBjbGFzcz0iY29sLXNtLTYgY29sLXhsLTMiPgogICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQga3BpLWdyZWVuIj4KICAgICAgICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJrcGktdGl0bGUiPlByZXNlbnRhY2lvbmVzIENhcmdhZGFzPC9kaXY+CiAgICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXZhbHVlIj57eyBjb25fYXJjaGl2byB9fSAvIHt7IHRvdGFsX3BvbmVudGVzIH19PC9kaXY+CiAgICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXN1YiI+UG9yIHBvbmVudGUgZW4gY2FkYSBzZXNpw7NuPC9kaXY+CiAgICAgICAgICAgICAgICAgICAgPHN2ZyBjbGFzcz0ia3BpLXdhdmUiIHZpZXdCb3g9IjAgMCAyMDAgNDAiIHByZXNlcnZlQXNwZWN0UmF0aW89Im5vbmUiPjxwYXRoIGQ9Ik0wLDMyIEwzNSwxOCBMNzUsMjYgTDExNSwxMCBMMTU1LDIwIEwyMDAsOCBMMjAwLDQwIEwwLDQwIFoiIGZpbGw9IiNmZmZmZmYiLz48L3N2Zz4KICAgICAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPGRpdiBjbGFzcz0iY29sLXNtLTYgY29sLXhsLTMiPgogICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQga3BpLW9yYW5nZSI+CiAgICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXRpdGxlIj5PcmRlbiBOaXZlbCAzPC9kaXY+CiAgICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXZhbHVlIj57eyBuM19jb2xzX3Jlc3VlbHRhc3xsZW5ndGggfX0gQ2FtcG9zPC9kaXY+CiAgICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXN1YiI+e3sgbjNfY29sc19yZXN1ZWx0YXN8am9pbignIMK3ICcpIGlmIG4zX2NvbHNfcmVzdWVsdGFzIGVsc2UgJ1BvciBkZWZlY3RvJyB9fTwvZGl2PgogICAgICAgICAgICAgICAgICAgIDxzdmcgY2xhc3M9ImtwaS13YXZlIiB2aWV3Qm94PSIwIDAgMjAwIDQwIiBwcmVzZXJ2ZUFzcGVjdFJhdGlvPSJub25lIj48cGF0aCBkPSJNMCwyMCBMNDUsMjggTDg1LDE0IEwxMzAsMjYgTDE3MCwxNSBMMjAwLDI0IEwyMDAsNDAgTDAsNDAgWiIgZmlsbD0iI2ZmZmZmZiIvPjwvc3ZnPgogICAgICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgICAgIDwvZGl2PgogICAgICAgIDwvZGl2PgoKICAgICAgICA8ZGl2IGlkPSJhbGVydGFBY2Npb24iIGNsYXNzPSJhbGVydCBkLW5vbmUgbWItNCI+PC9kaXY+CgogICAgICAgIDxkaXYgY2xhc3M9InZhbGV4LWNhcmQgbWItNCI+CiAgICAgICAgICAgIDxkaXYgY2xhc3M9InZhbGV4LWNhcmQtaGVhZGVyIj4KICAgICAgICAgICAgICAgIDxkaXY+CiAgICAgICAgICAgICAgICAgICAgPGg1IGNsYXNzPSJ2YWxleC1jYXJkLXRpdGxlIj5BZ2VuZGEgRW5nbG9iYWRhIHBvciBTZXNpb25lczwvaDU+CiAgICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0idmFsZXgtY2FyZC1zdWJ0aXRsZSI+CiAgICAgICAgICAgICAgICAgICAgICAgIE9yZGVuIE5pdmVsIDM6IDxzdHJvbmc+e3sgbjNfY29sc19yZXN1ZWx0YXN8am9pbignIOKGkiAnKSBpZiBuM19jb2xzX3Jlc3VlbHRhcyBlbHNlICdFc3TDoW5kYXInIH19PC9zdHJvbmc+IMK3IFN1YmNhcnBldGFzIHBvciBQYXJ0aWNpcGFudGUgMS4uMTAKICAgICAgICAgICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0iZC1mbGV4IGZsZXgtd3JhcCBhbGlnbi1pdGVtcy1jZW50ZXIgZ2FwLTIiPgogICAgICAgICAgICAgICAgICAgIDxzZWxlY3QgaWQ9ImZsdF9kaWEiIGNsYXNzPSJmb3JtLXNlbGVjdCBmb3JtLXNlbGVjdC1zbSIgc3R5bGU9IndpZHRoOiAxNTBweDsiIG9uY2hhbmdlPSJmaWx0cmFyQWdlbmRhKCkiPgogICAgICAgICAgICAgICAgICAgICAgICA8b3B0aW9uIHZhbHVlPSIiPlRvZG9zIGxvcyBkw61hczwvb3B0aW9uPgogICAgICAgICAgICAgICAgICAgICAgICB7JSBmb3IgZCBpbiBkaWFzX3VuaWNvcyAlfQogICAgICAgICAgICAgICAgICAgICAgICAgICAgPG9wdGlvbiB2YWx1ZT0ie3sgZHxsb3dlciB9fSI+e3sgZCB9fTwvb3B0aW9uPgogICAgICAgICAgICAgICAgICAgICAgICB7JSBlbmRmb3IgJX0KICAgICAgICAgICAgICAgICAgICA8L3NlbGVjdD4KICAgICAgICAgICAgICAgICAgICA8c2VsZWN0IGlkPSJmbHRfc2FsYSIgY2xhc3M9ImZvcm0tc2VsZWN0IGZvcm0tc2VsZWN0LXNtIiBzdHlsZT0id2lkdGg6IDIxMHB4OyIgb25jaGFuZ2U9ImZpbHRyYXJBZ2VuZGEoKSIgeyUgaWYgc2FsYV9yZXN0cmluZ2lkYSAlfWRpc2FibGVkeyUgZW5kaWYgJX0+CiAgICAgICAgICAgICAgICAgICAgICAgIDxvcHRpb24gdmFsdWU9IiI+VG9kYXMgbGFzIHNhbGFzPC9vcHRpb24+CiAgICAgICAgICAgICAgICAgICAgICAgIHslIGZvciBzIGluIHNhbGFzX3VuaWNhcyAlfQogICAgICAgICAgICAgICAgICAgICAgICAgICAgPG9wdGlvbiB2YWx1ZT0ie3sgc3xsb3dlciB9fSIgeyUgaWYgc2FsYV9yZXN0cmluZ2lkYSBhbmQgc3xsb3dlciA9PSBzYWxhX3Jlc3RyaW5naWRhfGxvd2VyICV9c2VsZWN0ZWR7JSBlbmRpZiAlfT57eyBzIH19PC9vcHRpb24+CiAgICAgICAgICAgICAgICAgICAgICAgIHslIGVuZGZvciAlfQogICAgICAgICAgICAgICAgICAgIDwvc2VsZWN0PgogICAgICAgICAgICAgICAgICAgIDxzZWxlY3QgaWQ9ImZsdF9lc3RhZG8iIGNsYXNzPSJmb3JtLXNlbGVjdCBmb3JtLXNlbGVjdC1zbSIgc3R5bGU9IndpZHRoOiAxNTBweDsiIG9uY2hhbmdlPSJmaWx0cmFyQWdlbmRhKCkiPgogICAgICAgICAgICAgICAgICAgICAgICA8b3B0aW9uIHZhbHVlPSIiPkVzdGFkbzogVG9kb3M8L29wdGlvbj4KICAgICAgICAgICAgICAgICAgICAgICAgPG9wdGlvbiB2YWx1ZT0iY2FyZ2FkbyI+Q2FyZ2FkYTwvb3B0aW9uPgogICAgICAgICAgICAgICAgICAgICAgICA8b3B0aW9uIHZhbHVlPSJwZW5kaWVudGUiPlBlbmRpZW50ZTwvb3B0aW9uPgogICAgICAgICAgICAgICAgICAgIDwvc2VsZWN0PgogICAgICAgICAgICAgICAgICAgIHslIGlmIGVzX2FkbWluICV9CiAgICAgICAgICAgICAgICAgICAgICAgIDxidXR0b24gY2xhc3M9ImJ0biBidG4tc20gYnRuLXZhbGV4LWdyZWVuIHB4LTMiIG9uY2xpY2s9ImNyZWFyQXJib2xDYXJwZXRhcygpIj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgIDxpIGNsYXNzPSJiaSBiaS1mb2xkZXItcGx1cyBtZS0xIj48L2k+IEdlbmVyYXIgQ2FycGV0YXMKICAgICAgICAgICAgICAgICAgICAgICAgPC9idXR0b24+CiAgICAgICAgICAgICAgICAgICAgeyUgZW5kaWYgJX0KICAgICAgICAgICAgICAgICAgICA8YnV0dG9uIGNsYXNzPSJidG4gYnRuLXNtIGJ0bi12YWxleC1ibHVlIHB4LTMiIG9uY2xpY2s9ImxpbXBpYXJGaWx0cm9zKCkiPgogICAgICAgICAgICAgICAgICAgICAgICA8aSBjbGFzcz0iYmkgYmkteC1jaXJjbGUgbWUtMSI+PC9pPiBMaW1waWFyCiAgICAgICAgICAgICAgICAgICAgPC9idXR0b24+CiAgICAgICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgPC9kaXY+CgogICAgICAgIDxkaXYgaWQ9ImNvbnRlbmVkb3JTZXNpb25lcyI+CiAgICAgICAgICAgIHslIGZvciBzZXNpb24gaW4gc2VzaW9uZXNfYWdydXBhZGFzICV9CiAgICAgICAgICAgIDxkaXYgY2xhc3M9InNlc2lvbi1ibG9xdWUgdGVtYS17eyBsb29wLmluZGV4MCAlIDQgfX0iCiAgICAgICAgICAgICAgICAgZGF0YS1zYWxhPSJ7eyBzZXNpb24uc2FsYXxsb3dlciB9fSIKICAgICAgICAgICAgICAgICBkYXRhLWRpYT0ie3sgc2VzaW9uLmRpYXxsb3dlciB9fSIKICAgICAgICAgICAgICAgICBkYXRhLXNlc2lvbi10ZXh0bz0ie3sgc2VzaW9uLnRleHRvX2J1c3F1ZWRhfGxvd2VyIH19Ij4KCiAgICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJzZXNpb24taGVhZGVyIj4KICAgICAgICAgICAgICAgICAgICA8ZGl2PgogICAgICAgICAgICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJkLWZsZXggZmxleC13cmFwIGFsaWduLWl0ZW1zLWNlbnRlcSBnYXAtMiBtYi0yIj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgIHslIGZvciBtZXRhIGluIHNlc2lvbi5jYW1wb3Nfbml2ZWxfMyAlfQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIDxzcGFuIGNsYXNzPSJzZXNpb24tbWV0YS1jaGlwIj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgPHN0cm9uZyBjbGFzcz0idGV4dC1pbmZvIj57eyBtZXRhLmV0aXF1ZXRhIH19Ojwvc3Ryb25nPiB7eyBtZXRhLnZhbG9yIH19CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgPC9zcGFuPgogICAgICAgICAgICAgICAgICAgICAgICAgICAgeyUgZW5kZm9yICV9CiAgICAgICAgICAgICAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJzbWFsbCIgc3R5bGU9ImNvbG9yOiAjY2JkNWUxOyI+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8aSBjbGFzcz0iYmkgYmktZm9sZGVyMiB0ZXh0LXdhcm5pbmcgbWUtMSI+PC9pPgogICAgICAgICAgICAgICAgICAgICAgICAgICAgPGNvZGU+e3sgc2VzaW9uLnJ1dGFfc2VzaW9uIH19PC9jb2RlPgogICAgICAgICAgICAgICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgICAgICAgICA8L2Rpdj4KCiAgICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0iZC1mbGV4IGFsaWduLWl0ZW1zLWNlbnRlcSBnYXAtMiI+CiAgICAgICAgICAgICAgICAgICAgICAgIDxzcGFuIGNsYXNzPSJiYWRnZSBiZy1kYXJrIGJvcmRlciBib3JkZXItc2Vjb25kYXJ5IHB4LTIgcHktMiI+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8aSBjbGFzcz0iYmkgYmktcGVvcGxlLWZpbGwgdGV4dC1pbmZvIG1lLTEiPjwvaT4ge3sgc2VzaW9uLnBvbmVudGVzfGxlbmd0aCB9fSBwb25lbnRlKHMpCiAgICAgICAgICAgICAgICAgICAgICAgIDwvc3Bhbj4KICAgICAgICAgICAgICAgICAgICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGNsYXNzPSJidG4gYnRuLXNtIGJ0bi1vdXRsaW5lLWxpZ2h0IiBvbmNsaWNrPSdhYnJpclJ1dGEoe3sgc2VzaW9uLnJ1dGFfc2VzaW9uIHwgdG9qc29uIH19LCB7eyAidHJ1ZSIgaWYgZXNfYWRtaW4gZWxzZSAiZmFsc2UiIH19KScgdGl0bGU9IkFicmlyIGNhcnBldGEgcmHDrXogZGUgbGEgc2VzacOzbiI+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8aSBjbGFzcz0iYmkgYmktZm9sZGVyLXN5bWxpbmsgbWUtMSI+PC9pPiBDYXJwZXRhIFNlc2nDs24KICAgICAgICAgICAgICAgICAgICAgICAgPC9idXR0b24+CiAgICAgICAgICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgICAgICAgICA8L2Rpdj4KCiAgICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJzZXNpb24tYm9keSI+CiAgICAgICAgICAgICAgICAgICAgeyUgZm9yIHAgaW4gc2VzaW9uLnBvbmVudGVzICV9CiAgICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0icG9uZW50ZS1jYXJkIgogICAgICAgICAgICAgICAgICAgICAgICAgZGF0YS1lc3RhZG89Int7ICdjYXJnYWRvJyBpZiBwLnRpZW5lX3BwdCBlbHNlICdwZW5kaWVudGUnIH19IgogICAgICAgICAgICAgICAgICAgICAgICAgZGF0YS1wb25lbnRlLXRleHRvPSJ7eyBwLnRleHRvX2J1c3F1ZWRhfGxvd2VyIH19Ij4KCiAgICAgICAgICAgICAgICAgICAgICAgIDxkaXYgY2xhc3M9ImQtZmxleCBhbGlnbi1pdGVtcy1jZW50ZXIgZ2FwLTMiPgogICAgICAgICAgICAgICAgICAgICAgICAgICAgPGRpdiBzdHlsZT0id2lkdGg6IDM2cHg7IGhlaWdodDogMzZweDsgYm9yZGVyLXJhZGl1czogNTAlOyBiYWNrZ3JvdW5kOiByZ2JhKDU2LCAxODksIDI0OCwgMC4xNik7IGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoNTYsIDE4OSwgMjQ4LCAwLjM1KTsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7Ij4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICA8aSBjbGFzcz0iYmkgYmktcGVyc29uLWZpbGwgdGV4dC1pbmZvIj48L2k+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgIDxkaXY+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0iZnctYm9sZCB0ZXh0LXdoaXRlIiBzdHlsZT0iZm9udC1zaXplOiAwLjk2cmVtOyI+e3sgcC5wb25lbnRlIH19PC9kaXY+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgeyUgaWYgcC5kZXRhbGxlX2V4dHJhICV9CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIDxkaXYgY2xhc3M9InNtYWxsIiBzdHlsZT0iY29sb3I6IHZhcigtLXRleHQtbXV0ZWQpOyI+e3sgcC5kZXRhbGxlX2V4dHJhIH19PC9kaXY+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgeyUgZW5kZm9yICV9CiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgICAgICAgICAgICAgPC9kaXY+CgogICAgICAgICAgICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJkLWZsZXggYWxpZ24taXRlbXMtY2VudGVyIGZsZXgtd3JhcCBnYXAtMiI+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICB7JSBpZiBwLnRpZW5lX3BwdCAlfQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIDxzcGFuIGNsYXNzPSJjYXJ0ZWwtZXN0YWRvIGNhcnRlbC1jYXJnYWRhIj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgPGkgY2xhc3M9ImJpIGJpLWNoZWNrLWNpcmNsZS1maWxsIj48L2k+IENBUkdBREEgKHt7IHAudW1fYXJjaGl2b3MgfX0pCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgPC9zcGFuPgogICAgICAgICAgICAgICAgICAgICAgICAgICAgeyUgZWxzZSAlfQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIDxzcGFuIGNsYXNzPSJjYXJ0ZWwtZXN0YWRvIGNhcnRlbC1wZW5kaWVudGUiPgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICA8aSBjbGFzcz0iYmkgYmktb3VybGdsYXNzLXNwbGl0Ij48L2k+IFBFTkRJRU5URQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIDwvc3Bhbj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgIHslIGVuZGlmICV9CgogICAgICAgICAgICAgICAgICAgICAgICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGNsYXNzPSJidG4gYnRuLXNtIGJ0bi12YWxleC1ibHVlIHB4LTMiIG9uY2xpY2s9J2FicmlyUnV0YSh7eyBwLnJ1dGFfY2FycGV0YSB8IHRvanNvbiB9fSwge3sgInRydWUiIGlmIGVzX2FkbWluIGVsc2UgImZhbHNlIiB9fSknPgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIDxpIGNsYXNzPSJiaSBiaS1mb2xkZXIyLW9wZW4gbWUtMSI+PC9pPiBBYnJpciBjYXJwZXRhCiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8L2J1dHRvbj4KCiAgICAgICAgICAgICAgICAgICAgICAgICAgICB7JSBpZiBlc19hZG1pbiAlfQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIDxsYWJlbCBjbGFzcz0iYnRuIGJ0bi1zbSBidG4tdmFsZXgtZ3JlZW4gbWItMCBweC0zIiB0aXRsZT0iU3ViaXIgcHJlc2VudGFjacOzbiBhIGxhIGNhcnBldGEgZGUge3sgcC5wb25lbnRlIH19Ij4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgPGkgY2xhc3M9ImJpIGJpLXVwbG9hZCBtZS0xIj48L2k+IFN1YmlyCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIDxpbnB1dCB0eXBlPSJmaWxlIiBjbGFzcz0iZC1ub25lIiBvbmNoYW5nZT0nc3ViaXJQcmVzZW50YWNpb24odGhpcywge3sgcC5ydXRhX2NhcnBldGEgfCB0b2pzb24gfX0pJz4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICA8L2xhYmVsPgogICAgICAgICAgICAgICAgICAgICAgICAgICAgeyUgZW5kaWYgJX0KICAgICAgICAgICAgICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgICAgICAgICAgICAgeyUgZW5kZm9yICV9CiAgICAgICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgICAgIHslIGVuZGZvciAlfQogICAgICAgIDwvZGl2PgoKICAgICAgICB7JSBlbmRpZiAlfQogICAgPC9kaXY+Cgo8c2NyaXB0Pgpjb25zdCBTQUxBX0JMT1FVRUFEQSA9IHt7IHNhbGFfcmVzdHJpbmdpZGEgfCB0b2pzb24gfX07CgpmdW5jdGlvbiBtb3N0cmFyQWxlcnRhKHRleHRvLCBvayA9IHRydWUpIHsKICAgIGNvbnN0IGVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2FsZXJ0YUFjY2lvbicpOwogICAgZWwuY2xhc3NOYW1lID0gJ2FsZXJ0IG1iLTQnOwogICAgZWwuc3R5bGUuYmFja2dyb3VuZCA9IG9rID8gJ3JnYmEoMiwxNTAsMTAyLDAuMiknIDogJ3JnYmEoMjQ5LDU4LDkwLDAuMiknOwogICAgZWwuc3R5bGUuYm9yZGVyID0gb2sgPyAnMXB4IHNvbGlkICM0OGQ2YTgnIDogJzFweCBzb2xpZCAjZjc3NzhjJzsKICAgIGVsLnN0eWxlLmNvbG9yID0gb2sgPyAnIzZlZTdiNycgOiAnI2ZkYTRhZic7CiAgICBlbC5pbm5lckhUTUwgPSBgPGkgY2xhc3M9ImJpICR7b2sgPyAnYmktY2hlY2stY2lyY2xlLWZpbGwnIDogJ2JpLWV4Y2xhbWF0aW9uLXRyaWFuZ2xlLWZpbGwnfSBtZS0yIj48L2k+JHt0ZXh0b31gOwp9CgpmdW5jdGlvbiBmaWphclB1ZXN0bygpIHsKICAgIGNvbnN0IHB1ZXN0byA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzZWxlY3RQdWVzdG8nKS52YWx1ZTsKICAgIGlmICghcHVlc3RvKSByZXR1cm47CiAgICBmZXRjaCgnL2FwaS9maWphci1wdWVzdG8nLCB7CiAgICAgICAgbWV0aG9kOiAnUE9TVCcsCiAgICAgICAgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sCiAgICAgICAgYm9keTogSlNPTi5zdHJpbmdpZnkoeyBwdWVzdG8gfSkKICAgIH0pCiAgICAudGhlbihyID0+IHIuanNvbigpKQogICAgLnRoZW4oZGF0YSA9PiB7CiAgICAgICAgbW9zdHJhckFsZXJ0YShkYXRhLm1lbnNhamUsIGRhdGEub2spOwogICAgICAgIGlmIChkYXRhLm9rKSBzZXRUaW1lb3V0KCgpID0+IGxvY2F0aW9uLnJlbG9hZCgpLCA3MDApOwogICAgfSk7Cn0KCmZ1bmN0aW9uIGNyZWFyQXJib2xDYXJwZXRhcygpIHsKICAgIGZldGNoKCcvYXBpL2NyZWFyLWNhcnBldGFzJywgeyBtZXRob2Q6ICdQT1NUJyB9KQogICAgICAgIC50aGVuKHIgPT4gci5qc29uKCkpCiAgICAgICAgLnRoZW4oZGF0YSA9PiB7CiAgICAgICAgICAgIG1vc3RyYXJBbGVydGEoZGF0YS5tZW5zYWplLCBkYXRhLm9rKTsKICAgICAgICAgICAgaWYgKGRhdGEub2spIHNldFRpbWVvdXQoKCkgPT4gbG9jYXRpb24ucmVsb2FkKCksIDkwMCk7CiAgICAgICAgfSk7Cn0KCmZ1bmN0aW9uIGFicmlyUnV0YShydXRhUmVsYXRpdmEsIGNyZWFyU2lOb0V4aXN0ZSA9IHRydWUpIHsKICAgIGZldGNoKCcvYXBpL2FicmlyLXJ1dGEnLCB7CiAgICAgICAgbWV0aG9kOiAnUE9TVCcsCiAgICAgICAgaGVhZGVyczogeyAnQ29udGVudC1UeXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nIH0sCiAgICAgICAgYm9keTogSlNPTi5zdHJpbmdpZnkoeyBydXRhOiBydXRhUmVsYXRpdmEsIGNyZWFyOiBjcmVhclNpTm9FeGlzdGUgfSkKICAgIH0pCiAgICAudGhlbihyID0+IHIuanNvbigpKQogICAgLnRoZW4oZGF0YSA9PiB7CiAgICAgICAgaWYgKCFkYXRhLm9rKSBtb3N0cmFyQWxlcnRhKGRhdGEubWVuc2FqZSwgZmFsc2UpOwogICAgfSk7Cn0KCmZ1bmN0aW9uIHN1YmlyUHJlc2VudGFjaW9uKGlucHV0LCBydXRhUmVsYXRpdmEpIHsKICAgIGlmICghaW5wdXQuZmlsZXMgfHwgIWlucHV0LmZpbGVzWzBdKSByZXR1cm47CiAgICBjb25zdCBmZCA9IG5ldyBGb3JtRGF0YSgpOwogICAgZmQuYXBwZW5kKCdhcmNoaXZvJywgaW5wdXQuZmlsZXNbMF0pOwogICAgZmQuYXBwZW5kKCdydXRhJywgcnV0YVJlbGF0aXZhKTsKICAgIGZldGNoKCcvYXBpL3N1YmlyLXByZXNlbnRhY2lvbicsIHsgbWV0aG9kOiAnUE9TVCcsIGJvZHk6IGZkIH0pCiAgICAgICAgLnRoZW4ociA9PiByLmpzb24oKSkKICAgICAgICAudGhlbihkYXRhID0+IHsKICAgICAgICAgICAgbW9zdHJhckFsZXJ0YShkYXRhLm1lbnNhamUsIGRhdGEub2spOwogICAgICAgICAgICBpZiAoZGF0YS5vaykgc2V0VGltZW91dCgoKSA9PiBsb2NhdGlvbi5yZWxvYWQoKSwgOTAwKTsKICAgICAgICB9KTsKfQoKZnVuY3Rpb24gZmlsdHJhckFnZW5kYSgpIHsKICAgIGNvbnN0IHEgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYnVzY2Fkb3JHbG9iYWwnKS52YWx1ZS50b0xvd2VyQ2FzZSgpLnRyaW0oKTsKICAgIGNvbnN0IGZsdERpYSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmbHRfZGlhJykudmFsdWUudG9Mb3dlckNhc2UoKS50cmltKCk7CiAgICBjb25zdCBmbHRTYWxhID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZsdF9zYWxhJykudmFsdWUudG9Mb3dlckNhc2UoKS50cmltKCk7CiAgICBjb25zdCBmbHRFc3RhZG8gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmx0X2VzdGFkbycpLnZhbHVlLnRvTG93ZXJDYXNlKCkudHJpbSgpOwoKICAgIGxldCBzZXNpb25lc1Zpc2libGVzID0gMDsKICAgIGxldCBwb25lbnRlc1Zpc2libGVzID0gMDsKCiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcuc2VzaW9uLWJsb3F1ZScpLmZvckVhY2goYmxvcXVlID0+IHsKICAgICAgICBjb25zdCBzYWxhQmxvcXVlID0gYmxvcXVlLmdldEF0dHJpYnV0ZSgnZGF0YS1zYWxhJykgfHwgJyc7CiAgICAgICAgY29uc3QgZGlhQmxvcXVlID0gYmxvcXVlLmdldEF0dHJpYnV0ZSgnZGF0YS1kaWEnKSB8fCAnJzsKICAgICAgICBjb25zdCB0ZXh0b1Nlc2lvbiA9IGJsb3F1ZS5nZXRBdHRyaWJ1dGUoJ2RhdGEtc2VzaW9uLXRleHRvJykgfHwgJyc7CgogICAgICAgIGlmIChTQUxBX0JMT1FVRUFEQSAmJiBzYWxhQmxvcXVlICE9PSBTQUxBX0JMT1FVRUFEQS50b0xvd2VyQ2FzZSgpKSB7CiAgICAgICAgICAgIGJsb3F1ZS5zdHlsZS5kaXNwbGF5ID0gJ25vbmUnOwogICAgICAgICAgICByZXR1cm47CiAgICAgICAgfQogICAgICAgIGlmIChmbHRTYWxhICYmIHNhbGFCbG9xdWUgIT09IGZsdFNhbGEpIHsKICAgICAgICAgICAgYmxvcXVlLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CiAgICAgICAgICAgIHJldHVybjsKICAgICAgICB9CiAgICAgICAgaWYgKGZsdERpYSAmJiBkaWFCbG9xdWUgIT09IGZsdERpYSkgewogICAgICAgICAgICBibG9xdWUuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKICAgICAgICAgICAgcmV0dXJuOwogICAgICAgIH0KCiAgICAgICAgY29uc3QgY29pbmNpZGVDb25TZXNpb24gPSAhcSB8fCB0ZXh0b1Nlc2lvbi5pbmNsdWRlcyhxKTsKICAgICAgICBsZXQgcG9uZW50ZXNFbkJsb3F1ZSA9IDA7CgogICAgICAgIGJsb3F1ZS5xdWVyeVNlbGVjdG9yQWxsKCcucG9uZW50ZS1jYXJkJykuZm9yRWFjaChjYXJkID0+IHsKICAgICAgICAgICAgY29uc3QgZXN0YWRvQ2FyZCA9IGNhcmQuZ2V0QXR0cmlidXRlKCdkYXRhLWVzdGFkbycpIHx8ICcnOwogICAgICAgICAgICBjb25zdCB0ZXh0b1BvbmVudGUgPSBjYXJkLmdldEF0dHJpYnV0ZSgnZGF0YS1wb25lbnRlLXRleHRvJykgfHwgJyc7CgogICAgICAgICAgICBjb25zdCBva0VzdGFkbyA9ICFmbHRFc3RhZG8gfHwgZXN0YWRvQ2FyZCA9PT0gZmx0RXN0YWRvOwogICAgICAgICAgICBjb25zdCBva1RleHRvID0gY29pbmNpZGVDb25TZXNpb24gfHwgdGV4dG9Qb25lbnRlLmluY2x1ZGVzKHEpOwoKICAgICAgICAgICAgaWYgKG9rRXN0YWRvICYmIG9rVGV4dG8pIHsKICAgICAgICAgICAgICAgIGNhcmQuc3R5bGUuZGlzcGxheSA9ICcnOwogICAgICAgICAgICAgICAgcG9uZW50ZXNFbkJsb3F1ZSsrOwogICAgICAgICAgICAgICAgcG9uZW50ZXNWaXNpYmxlcysrOwogICAgICAgICAgICB9IGVsc2UgewogICAgICAgICAgICAgICAgY2FyZC5zdHlsZS5kaXNwbGF5ID0gJ25vbmUnOwogICAgICAgICAgICB9CiAgICAgICAgfSk7CgogICAgICAgIGlmIChwb25lbnRlc0VuQmxvcXVlID4gMCkgewogICAgICAgICAgICBibG9xdWUuc3R5bGUuZGlzcGxheSA9ICcnOwogICAgICAgICAgICBzZXNpb25lc1Zpc2libGVzKys7CiAgICAgICAgfSBlbHNlIHsKICAgICAgICAgICAgYmxvcXVlLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CiAgICAgICAgfQogICAgfSk7CgogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2twaVNlc2lvbmVzJykudGV4dENvbnRlbnQgPSBzZXNpb25lc1Zpc2libGVzOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2twaVBvbmVudGVzJykudGV4dENvbnRlbnQgPSBwb25lbnRlc1Zpc2libGVzOwp9CgpmdW5jdGlvbiBsaW1waWFyRmlsdHJvcygpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdidXNjYWRvckdsb2JhbCcpLnZhbHVlID0gJyc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmx0X2RpYScpLnZhbHVlID0gJyc7CiAgICBpZiAoIVNBTEFfQkxPUVVFQURBKSB7CiAgICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZsdF9zYWxhJykudmFsdWUgPSAnJzsKICAgIH0KICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdmbHRfZXN0YWRvJykudmFsdWUgPSAnJzsKICAgIGZpbHRyYXJBZ2VuZGEoKTsKfQoKZmlsdHJhckFnZW5kYSgpOwo8L3NjcmlwdD4KPC9ib2R5Pgo8L2h0bWw+").decode("utf-8")


@app.route("/")
def index():
    ruta_activa = obtener_ruta_base()
    try:
        config, df, ruta_activa = cargar_datos_congreso()
        if config is None or df is None:
            return render_template_string(
                HTML_APP,
                config=None,
                puesto_activo=None,
                puesto_bloqueado=None,
                sala_restringida=None,
                es_admin=False,
            )

        columnas_raw = list(df.columns)
        n3_cols_resueltas = resolver_columnas_nivel_3(config, columnas_raw)
        columnas_ordenadas = ordenar_columnas_por_nivel_3(columnas_raw, n3_cols_resueltas)

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
            datos_fila_txt = " ".join(str(row[c]).replace("\xa0", " ") for c in columnas_ordenadas)

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

        total_ponentes = sum(len(s["ponentes"]) for s in sesiones_agrupadas)
        con_archivo = sum(1 for s in sesiones_agrupadas for p in s["ponentes"] if p["tiene_ppt"])

        return render_template_string(
            HTML_APP,
            config=config,
            n3_cols_resueltas=n3_cols_resueltas,
            sesiones_agrupadas=sesiones_agrupadas,
            total_ponentes=total_ponentes,
            con_archivo=con_archivo,
            dias_unicos=dias_unicos,
            salas_unicas=salas_unicas,
            puesto_activo=puesto_activo,
            puesto_bloqueado=puesto_bloqueado,
            sala_restringida=sala_restringida,
            es_admin=es_admin,
        )
    except Exception:
        return render_template_string(
            HTML_ERROR,
            ruta_activa=ruta_activa,
            detalle_error=traceback.format_exc(),
        ), 500


@app.route("/api/fijar-puesto", methods=["POST"])
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


@app.route("/api/crear-carpetas", methods=["POST"])
def api_crear_carpetas():
    config, df, ruta_activa = cargar_datos_congreso()
    if config is None:
        return jsonify({"ok": False, "mensaje": "Faltan archivos de configuración."})
    puesto = config.get("_puesto_bloqueado") or ""
    if puesto and not puesto.upper().startswith("ADM"):
        return jsonify({"ok": False, "mensaje": "Solo los puestos ADMIN pueden crear o modificar carpetas."})

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

    return jsonify({"ok": True, "mensaje": f"Carpetas de sesión y subcarpetas por ponente listas ({creadas} nuevas)."})


@app.route("/api/abrir-ruta", methods=["POST"])
def api_abrir_ruta():
    config, _, ruta_activa = cargar_datos_congreso()
    data = request.json or {}
    ruta_rel = data.get("ruta", "PPT")
    crear = bool(data.get("crear", True))
    puesto = config.get("_puesto_bloqueado") or "" if config else ""
    es_admin = (not puesto) or puesto.upper().startswith("ADM")

    ruta_abs = os.path.join(ruta_activa, ruta_rel)
    if crear and es_admin:
        os.makedirs(ruta_abs, exist_ok=True)
    elif not os.path.exists(ruta_abs):
        return jsonify({"ok": False, "mensaje": "La carpeta aún no existe en disco."})

    if SISTEMA == "Windows":
        os.startfile(ruta_abs)
    elif SISTEMA == "Darwin":
        subprocess.Popen(["open", ruta_abs])
    else:
        subprocess.Popen(["xdg-open", ruta_abs])
    return jsonify({"ok": True})


@app.route("/api/subir-presentacion", methods=["POST"])
def api_subir_presentacion():
    config, _, ruta_activa = cargar_datos_congreso()
    puesto = config.get("_puesto_bloqueado") or "" if config else ""
    if puesto and not puesto.upper().startswith("ADM"):
        return jsonify({"ok": False, "mensaje": "Acceso denegado: las salas solo pueden ver presentaciones."})

    archivo = request.files.get("archivo")
    ruta_rel = request.form.get("ruta", "PPT")
    if not archivo or not archivo.filename:
        return jsonify({"ok": False, "mensaje": "No se seleccionó ningún archivo."})

    carpeta_destino = os.path.join(ruta_activa, ruta_rel)
    os.makedirs(carpeta_destino, exist_ok=True)
    nombre_base, ext = os.path.splitext(archivo.filename)
    nombre_seguro = sanitize_name(nombre_base) + sanitize_name(ext).replace("_vacio", "")
    archivo.save(os.path.join(carpeta_destino, nombre_seguro))
    return jsonify({"ok": True, "mensaje": f"Presentación '{nombre_seguro}' cargada en {ruta_rel}."})


def limpiar_bloqueo_al_salir(ruta_archivo_bloqueo):
    try:
        if ruta_archivo_bloqueo and os.path.exists(ruta_archivo_bloqueo):
            os.remove(ruta_archivo_bloqueo)
    except Exception:
        pass


if __name__ == "__main__":
    cfg, _, _ = cargar_datos_congreso()
    ruta_bloq = cfg.get("_ruta_puesto_sesion") if cfg else None
    limpiar_bloqueo_al_salir(ruta_bloq)
    try:
        threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:5051")).start()
    except Exception:
        pass
    try:
        app.run(host="127.0.0.1", port=5051, debug=False, use_reloader=False)
    finally:
        limpiar_bloqueo_al_salir(ruta_bloq)
