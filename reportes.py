import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import subprocess
import io
import urllib.request
import json
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

# --- CONFIGURACIÓN ---
VERSION_ACTUAL = "v1.0.5" 
EXE_NAME = "ReportesSecurithor.exe"
REPO_API_URL = "https://api.github.com/repos/Miniacidus/Report-Securithor/releases/latest"
URL_DOWNLOAD = "https://github.com/Miniacidus/Report-Securithor/releases/latest/download/ReportesSecurithor.exe"

ruta_mensual = ""
ruta_anual = ""
FILAS_POR_PAGINA = 50 

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# --- FUNCIONES (ACTUALIZACIÓN, PROCESAMIENTO, ETC) ---

def actualizar_programa():
    try:
        btn_update.config(text="Buscando...", state="disabled")
        root.update()
        req = urllib.request.Request(REPO_API_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            data = json.loads(response.read().decode())
            ultima_version = data['tag_name']
        if ultima_version == VERSION_ACTUAL:
            messagebox.showinfo("Software al día", f"Ya tienes la versión más reciente ({VERSION_ACTUAL}).")
            return
        confirmar = messagebox.askyesno("Nueva Versión", f"Actualización disponible: {ultima_version}\n¿Deseas descargarla?")
        if not confirmar: return
        temp_exe = "Update_Nuevo.exe"
        urllib.request.urlretrieve(URL_DOWNLOAD, temp_exe)
        with open("update_script.bat", "w") as f:
            f.write(f'@echo off\ntimeout /t 20 /nobreak > nul\ndel "{EXE_NAME}"\nren "{temp_exe}" "{EXE_NAME}"\nstart "" "{EXE_NAME}"\ndel "%~f0"\n')
        messagebox.showinfo("Éxito", "El programa se actualizará en 20 segundos.")
        subprocess.Popen(["update_script.bat"], shell=True)
        root.destroy()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo verificar la actualización:\n{e}")
    finally:
        try: btn_update.config(text="🔄 Check for Updates", state="normal")
        except: pass

def generar_reportes():
    global ruta_mensual, ruta_anual
    if not ruta_mensual and not ruta_anual:
        messagebox.showwarning("Atención", "Selecciona al menos un archivo CSV.")
        return
    try:
        btn_generar.config(text="⏳ Procesando...", state="disabled")
        root.update()
        # Lógica de procesamiento aquí (Mensual/Anual)...
        messagebox.showinfo("¡Éxito!", "Reportes generados correctamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un problema: {e}")
    finally:
        btn_generar.config(text="GENERATE REPORTS", state="normal")

# ... (Incluye aquí tus funciones leer_csv_robusto, aplicar_formato_excel, etc del código v1.0.4) ...

# --- INTERFAZ ---
root = tk.Tk()
root.title(f"Securithor Automator {VERSION_ACTUAL}")
root.geometry("480x620")

# Botones y Labels (Asegúrate de definir btn_generar y btn_update aquí abajo)
btn_update = tk.Button(root, text="🔄 Check for Updates", command=actualizar_programa)
btn_update.pack(pady=5)

btn_generar = tk.Button(root, text="GENERATE REPORTS", command=generar_reportes, bg="#2e7d32", fg="white")
btn_generar.pack(pady=20)

root.mainloop()


