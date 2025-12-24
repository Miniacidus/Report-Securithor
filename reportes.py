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

# --- CONFIGURACIÓN DE VERSIÓN Y GITHUB ---
VERSION_ACTUAL = "v1.0.6" 
EXE_NAME = "ReportesSecurithor.exe"
REPO_API_URL = "https://api.github.com/repos/Miniacidus/Report-Securithor/releases/latest"
URL_DOWNLOAD = "https://github.com/Miniacidus/Report-Securithor/releases/latest/download/ReportesSecurithor.exe"

# Variables Globales inicializadas
ruta_mensual = ""
ruta_anual = ""
FILAS_POR_PAGINA = 50 

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# --- SISTEMA DE ACTUALIZACIÓN ---
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

        if messagebox.askyesno("Nueva Versión", f"Actualización disponible: {ultima_version}\n¿Descargar?"):
            temp_exe = "Update_Nuevo.exe"
            urllib.request.urlretrieve(URL_DOWNLOAD, temp_exe)
            with open("update_script.bat", "w") as f:
                f.write(f'@echo off\ntitle Actualizando...\necho Espere 20 segundos...\ntimeout /t 20 /nobreak > nul\ndel "{EXE_NAME}"\nren "{temp_exe}" "{EXE_NAME}"\nstart "" "{EXE_NAME}"\ndel "%~f0"\n')
            messagebox.showinfo("Éxito", "El programa se reiniciará en 20 segundos.")
            subprocess.Popen(["update_script.bat"], shell=True)
            root.destroy()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo actualizar: {e}")
    finally:
        try: btn_update.config(text="🔄 Check for Updates", state="normal")
        except: pass

# --- MOTOR DE PROCESAMIENTO ---
def corregir_fecha(texto):
    if pd.isna(texto): return pd.NaT
    texto = str(texto).lower().replace('p. m.', 'pm').replace('a. m.', 'am').replace('p.m.', 'pm').replace('a.m.', 'am')
    try: return pd.to_datetime(texto, dayfirst=True)
    except: return pd.NaT

def leer_csv_robusto(ruta):
    try:
        with open(ruta, 'r', encoding='utf-8-sig', errors='replace') as f:
            contenido = f.read()
        df = pd.read_csv(io.StringIO(contenido), header=3)
        df.columns = df.columns.str.strip()
        posibles = ['Llegada', 'Recibido', 'Fecha', 'Arrival', 'Received']
        col_f = next((c for c in posibles if c in df.columns), None)
        df = df.rename(columns={col_f: 'Llegada'})
        df['FechaCompleta'] = df['Llegada'].apply(corregir_fecha)
        df = df.dropna(subset=['FechaCompleta'])
        df['Cuenta'] = df['Cuenta'].astype(str).str.replace(r'\.0$', '', regex=True)
        df['Cuenta_Num'] = pd.to_numeric(df['Cuenta'], errors='coerce').fillna(0)
        return df
    except Exception as e:
        raise Exception(f"Error en CSV: {e}")

def aplicar_formato_excel(ruta_excel, modo="asistencia"):
    wb = load_workbook(ruta_excel)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Ajustes de impresión para 31 días
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.sheet_properties.pageSetUpPr.fitToPage = True 
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0 
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        
        for row in ws.iter_rows():
            for cell in row:
                cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                cell.alignment = Alignment(horizontal='center', vertical='center')
    wb.save(ruta_excel)

def generar_reportes():
    global ruta_mensual, ruta_anual
    if not ruta_mensual and not ruta_anual:
        messagebox.showwarning("Atención", "Selecciona un archivo.")
        return
    
    try: # CORRECCIÓN: Ahora el try tiene su except y finally
        btn_generar.config(text="⏳ Procesando...", state="disabled")
        root.update()
        
        if ruta_mensual:
            df = leer_csv_robusto(ruta_mensual)
            fecha_max = df['FechaCompleta'].max()
            df['Dia'] = df['FechaCompleta'].dt.day
            df['Hora'] = df['FechaCompleta'].dt.strftime('%H:%M')
            pivot = df.pivot_table(index=['Cuenta_Num', 'Cuenta'], columns='Dia', values='Hora', aggfunc='max').fillna('/')
            out = f"1_Reporte_Asistencia_{MESES_ES[fecha_max.month]}.xlsx"
            pivot.to_excel(out)
            aplicar_formato_excel(out, "asistencia")
            
        messagebox.showinfo("¡Éxito!", "Proceso completado.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un problema: {e}")
    finally:
        btn_generar.config(text="GENERATE REPORTS", state="normal")

# --- INTERFAZ ---
def sel_m(): 
    global ruta_mensual
    r = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if r: ruta_mensual = r; lbl_m.config(text=os.path.basename(r))

root = tk.Tk()
root.title(f"Securithor Automator {VERSION_ACTUAL}")
root.geometry("400x400")

lbl_m = tk.Label(root, text="(vacio)")
lbl_m.pack()
tk.Button(root, text="Select Monthly CSV", command=sel_m).pack()
btn_update = tk.Button(root, text="🔄 Check for Updates", command=actualizar_programa)
btn_update.pack(pady=10)
btn_generar = tk.Button(root, text="GENERATE REPORTS", command=generar_reportes, bg="green", fg="white")
btn_generar.pack(pady=20)

root.mainloop()
