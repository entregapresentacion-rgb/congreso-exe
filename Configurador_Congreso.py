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

        <form method="POST
