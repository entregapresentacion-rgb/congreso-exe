import os
import sys
import json
import shutil
import platform
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd

SISTEMA = platform.system()

def obtener_ruta_base():
    if getattr(sys, 'frozen', False):
        ruta = os.path.dirname(sys.executable)
        if SISTEMA == 'Darwin' and ruta.endswith('Contents/MacOS'):
            return os.path.abspath(os.path.join(ruta, '..', '..', '..'))
        return ruta
    return os.path.abspath(os.path.dirname(__file__))

class ConfiguradorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Configurador de Paquete USB - Congreso (Win + Mac)")
        self.root.geometry("720x640")
        self.root.configure(bg="#f0f2f5")

        self.ruta_excel = tk.StringVar()
        self.workgroup = tk.StringVar(value="CONGRESO")
        self.num_admins = tk.IntVar(value=2)
        self.columnas_disponibles = []
        self.orden_carpetas = []

        tk.Label(root, text="Configurador de Congreso (Generador de Paquete USB)", font=("Helvetica", 14, "bold"), bg="#111827", fg="white", pady=14).pack(fill="x")

        main = tk.Frame(root, bg="#f0f2f5", padx=20, pady=12)
        main.pack(fill="both", expand=True)

        tk.Label(main, text="1. Selecciona el archivo Excel del Congreso:", font=("Helvetica", 9, "bold"), bg="#f0f2f5").pack(anchor="w")
        f1 = tk.Frame(main, bg="#f0f2f5")
        f1.pack(fill="x", pady=(4, 10))
        tk.Entry(f1, textvariable=self.ruta_excel, state="readonly", font=("Helvetica", 10)).pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(f1, text="Examinar Excel...", command=self.cargar_excel, bg="#2563eb", fg="white", font=("Helvetica", 9, "bold"), padx=12).pack(side="right", padx=(8, 0))

        f2 = tk.Frame(main, bg="#f0f2f5")
        f2.pack(fill="x", pady=(0, 10))
        
        col_izq = tk.Frame(f2, bg="#f0f2f5")
        col_izq.pack(side="left", fill="x", expand=True)
        tk.Label(col_izq, text="2. Grupo de Trabajo (Red):", font=("Helvetica", 9, "bold"), bg="#f0f2f5").pack(anchor="w")
        tk.Entry(col_izq, textvariable=self.workgroup, font=("Helvetica", 10), width=22).pack(anchor="w", pady=3)

        col_der = tk.Frame(f2, bg="#f0f2f5")
        col_der.pack(side="left", fill="x", expand=True)
        tk.Label(col_der, text="Cantidad de ordenadores Admin (ADM1, ADM2...):", font=("Helvetica", 9, "bold"), bg="#f0f2f5").pack(anchor="w")
        tk.Spinbox(col_der, from_=1, to=20, textvariable=self.num_admins, width=8, font=("Helvetica", 10), command=self.actualizar_lista_equipos).pack(anchor="w", pady=3)

        tk.Label(main, text="3. Orden de subcarpetas dentro de PPT (Haz clic en las columnas en orden):", font=("Helvetica", 9, "bold"), bg="#f0f2f5").pack(anchor="w")
        self.frame_cols = tk.Frame(main, bg="#f0f2f5")
        self.frame_cols.pack(fill="x", pady=4)
        self.lbl_orden = tk.Label(main, text="Estructura actual: PPT/Sala/Sesion/", font=("Courier", 9, "bold"), fg="#0d6efd", bg="#f0f2f5")
        self.lbl_orden.pack(anchor="w", pady=(0, 10))

        tk.Label(main, text="4. Nombres exactos de equipos (Puedes editar o añadir líneas):", font=("Helvetica", 9, "bold"), bg="#f0f2f5").pack(anchor="w")
        self.txt_equipos = scrolledtext.ScrolledText(main, height=7, font=("Courier", 10))
        self.txt_equipos.pack(fill="x", pady=(4, 14))

        tk.Button(main, text="📦 CREAR CARPETA LISTA PARA PENDRIVE", command=self.generar_paquete, bg="#10b981", fg="white", font=("Helvetica", 11, "bold"), pady=10).pack(fill="x")

    def cargar_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if not path:
            return
        self.ruta_excel.set(path)
        df = pd.read_excel(path, engine='openpyxl')
        self.columnas_disponibles = [str(c) for c in df.columns]
        
        for w in self.frame_cols.winfo_children():
            w.destroy()
        self.orden_carpetas = []
        for col in self.columnas_disponibles:
            tk.Button(self.frame_cols, text=f"+ {col}", command=lambda c=col: self.agregar_nivel_carpeta(c), bg="white", relief="groove", font=("Helvetica", 8)).pack(side="left", padx=2)
        tk.Button(self.frame_cols, text="Limpiar", command=self.limpiar_orden, bg="#fee2e2", fg="#991b1b", font=("Helvetica", 8)).pack(side="right")

        if "Sala" in self.columnas_disponibles: self.agregar_nivel_carpeta("Sala")
        if "Sesion" in self.columnas_disponibles: self.agregar_nivel_carpeta("Sesion")

        self.actualizar_lista_equipos()

    def agregar_nivel_carpeta(self, col):
        if col not in self.orden_carpetas:
            self.orden_carpetas.append(col)
            self.lbl_orden.config(text="Estructura actual: PPT/" + "/".join(self.orden_carpetas) + "/")

    def limpiar_orden(self):
        self.orden_carpetas = []
        self.lbl_orden.config(text="Estructura actual: PPT/ (Raíz)")

    def actualizar_lista_equipos(self):
        equipos = [f"ADM{i}" for i in range(1, self.num_admins.get() + 1)]
        if self.ruta_excel.get():
            df = pd.read_excel(self.ruta_excel.get(), engine='openpyxl')
            if 'Sala' in df.columns:
                salas = sorted(list(set(str(s).strip() for s in df['Sala'].dropna() if str(s).strip())))
                equipos.extend(salas)
        self.txt_equipos.delete("1.0", tk.END)
        self.txt_equipos.insert("1.0", "\n".join(equipos))

    def generar_paquete(self):
        if not self.ruta_excel.get():
            messagebox.showwarning("Falta Excel", "Selecciona primero el archivo Excel del congreso.")
            return

        base = obtener_ruta_base()
        salida = os.path.join(base, "PAQUETE_USB_CONGRESO")
        shutil.rmtree(salida, ignore_errors=True)
        os.makedirs(salida, exist_ok=True)

        equipos_finales = [line.strip() for line in self.txt_equipos.get("1.0", tk.END).splitlines() if line.strip()]
        config_data = {
            "workgroup": self.workgroup.get().strip().upper(),
            "orden_carpetas": self.orden_carpetas,
            "equipos": equipos_finales
        }

        with open(os.path.join(salida, "config_congreso.json"), "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)

        shutil.copy2(self.ruta_excel.get(), os.path.join(salida, "programa.xlsx"))

        copiado_algo = False
        exe_win = os.path.join(base, "CongresoApp.exe")
        app_mac = os.path.join(base, "CongresoApp.app")

        if os.path.exists(exe_win):
            shutil.copy2(exe_win, os.path.join(salida, "CongresoApp.exe"))
            copiado_algo = True
        if os.path.exists(app_mac):
            shutil.copytree(app_mac, os.path.join(salida, "CongresoApp.app"))
            copiado_algo = True

        if not copiado_algo:
            messagebox.showwarning("Aviso", "Se ha generado config_congreso.json y programa.xlsx, pero asegúrate de poner CongresoApp.exe o CongresoApp.app junto a ellos.")
        else:
            messagebox.showinfo("¡Listo!", f"Carpeta generada correctamente:\n{salida}\n\nLista para copiar al pendrive y usar en Windows y Mac Intel.")

        if SISTEMA == 'Windows':
            os.startfile(salida)
        elif SISTEMA == 'Darwin':
            subprocess.Popen(["open", salida])

if __name__ == "__main__":
    root = tk.Tk()
    app = ConfiguradorApp(root)
    root.mainloop()
