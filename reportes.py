import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import subprocess
import io
import urllib.request
import json
import webbrowser # <--- NUEVO: Para abrir el navegador
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

# --- CONFIGURACIÓN DE VERSIÓN Y GITHUB ---
VERSION_ACTUAL = "v1.0.14" 
EXE_NAME = "ReportesSecurithor.exe"
REPO_API_URL = "https://api.github.com/repos/Miniacidus/Report-Securithor/releases/latest"
URL_RELEASES = "https://github.com/Miniacidus/Report-Securithor/releases" # <--- NUEVO: Link directo

# Variables globales
ALTO_FILA = 30
ruta_mensual = ""
ruta_anual = ""
FILAS_POR_PAGINA = 30

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# --- SISTEMA DE ACTUALIZACIÓN (MODO MANUAL WEB) ---

def actualizar_programa():
    try:
        btn_update.config(text="Buscando...", state="disabled")
        root.update()

        # 1. Consultamos a GitHub cuál es la última versión
        req = urllib.request.Request(REPO_API_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            data = json.loads(response.read().decode())
            ultima_version = data['tag_name']

        # 2. Comparamos
        if ultima_version == VERSION_ACTUAL:
            messagebox.showinfo("Software al día", f"Ya tienes la versión más reciente ({VERSION_ACTUAL}).")
        else:
            # 3. MODIFICADO: En lugar de descargar, preguntamos si abrir la web
            msg = f"Nueva versión disponible: {ultima_version}\n\n¿Deseas ir a la página de descargas?"
            if messagebox.askyesno("Actualización Manual", msg):
                webbrowser.open(URL_RELEASES)

    except Exception as e:
        # Si falla la conexión, ofrecemos abrir la web por si acaso
        if messagebox.askyesno("Error de conexión", f"No se pudo verificar la versión.\n¿Quieres abrir la página de descargas manualmente?"):
             webbrowser.open(URL_RELEASES)
    finally:
        try: btn_update.config(text="🔄 Check for Updates", state="normal")
        except: pass

# --- GESTIÓN DE ARCHIVOS DE CONFIGURACIÓN ---

def asegurar_archivos_config():
    if not os.path.exists("bajas.txt"):
        with open("bajas.txt", "w", encoding="utf-8") as f: f.write("")
    if not os.path.exists("nombres.txt"):
        with open("nombres.txt", "w", encoding="utf-8") as f:
            f.write("Cuenta, Nombre")

def abrir_txt(archivo):
    asegurar_archivos_config()
    try: subprocess.Popen(["notepad.exe", archivo])
    except: messagebox.showerror("Error", f"No se pudo abrir {archivo}")

def cargar_bajas():
    if not os.path.exists("bajas.txt"): return []
    with open("bajas.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

def cargar_diccionario_nombres():
    nombres = {}
    if os.path.exists("nombres.txt"):
        with open("nombres.txt", "r", encoding="utf-8") as f:
            for line in f:
                if "," in line:
                    partes = line.strip().split(",", 1)
                    nombres[partes[0].strip()] = partes[1].strip()
    return nombres

# --- MOTOR DE PROCESAMIENTO (INTACTO v1.0.7) ---

def corregir_fecha(texto):
    if pd.isna(texto): return pd.NaT
    texto = str(texto).lower().replace('p. m.', 'pm').replace('a. m.', 'am').replace('p.m.', 'pm').replace('a.m.', 'am')
    try: return pd.to_datetime(texto, dayfirst=True)
    except: return pd.NaT

def leer_csv_robusto(ruta):
    try:
        with open(ruta, 'r', encoding='utf-8-sig', errors='replace') as f:
            contenido = f.read()
        if len(contenido) < 10:
            with open(ruta, 'r', encoding='latin-1', errors='replace') as f:
                contenido = f.read()

        df = pd.read_csv(io.StringIO(contenido), header=3)
        df.columns = df.columns.str.strip()

        posibles = ['Llegada', 'Recibido', 'Fecha', 'Arrival', 'Received']
        col_f = next((c for c in posibles if c in df.columns), None)
        if not col_f: raise Exception("Formato de columnas no reconocido.")

        df = df.rename(columns={col_f: 'Llegada'})
        df['FechaCompleta'] = df['Llegada'].apply(corregir_fecha)
        df = df.dropna(subset=['FechaCompleta'])
        
        df['Cuenta'] = df['Cuenta'].astype(str).str.replace(r'\.0$', '', regex=True)
        df['Cuenta_Num'] = pd.to_numeric(df['Cuenta'], errors='coerce').fillna(0)
        
        bajas = cargar_bajas()
        df = df[~df['Cuenta'].isin(bajas)]
        mapa = cargar_diccionario_nombres()
        df['Nombre del Cliente'] = df['Cuenta'].map(mapa).fillna('')
        
        return df
    except Exception as e:
        raise Exception(f"Error procesando el archivo CSV: {e}")

def aplicar_formato_excel(ruta_excel, modo="asistencia"):
    """Formatea el Excel para impresión profesional a escala con filas ajustadas"""
    wb = load_workbook(ruta_excel)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        borde_fino = Side(style="thin")
        
        color_fondo = "4472C4" 
        if sheet_name == "Sin_Reportar": color_fondo = "FFC000" 
        if sheet_name == "Solo_Sucesos": color_fondo = "9BC2E6" 

        for row in ws.iter_rows():
            # --- NUEVO: APLICAR ALTURA A LA FILA ACTUAL ---
            # Esto estira la fila verticalmente según el número ALTO_FILA
            ws.row_dimensions[row[0].row].height = ALTO_FILA

            for cell in row:
                cell.border = Border(left=borde_fino, right=borde_fino, top=borde_fino, bottom=borde_fino)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if cell.row == 1:
                    cell.font = Font(bold=True, color="FFFFFF", size=10)
                    cell.fill = PatternFill("solid", fgColor=color_fondo)
                else: cell.font = Font(size=9)
        
        ws.column_dimensions['A'].width = 12 
        ws.column_dimensions['B'].width = 42 
        ws.column_dimensions['C'].width = 18
        if ws.max_column >= 4: ws.column_dimensions['D'].width = 20

        if modo == "asistencia":
            for col in range(3, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(col)].width = 5.5

        # --- AJUSTES DE IMPRESIÓN (31 DÍAS EN UNA HOJA) ---
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.sheet_properties.pageSetUpPr.fitToPage = True 
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0 
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.print_title_rows = '1:1'
        
        for i in range(FILAS_POR_PAGINA + 2, ws.max_row + 1, FILAS_POR_PAGINA):
            ws.row_breaks.append(Break(id=i-1))
            
    wb.save(ruta_excel)

def generar_reportes():
    global ruta_mensual, ruta_anual
    if not ruta_mensual and not ruta_anual:
        messagebox.showwarning("Atención", "Selecciona al menos un archivo CSV.")
        return
    
    try: 
        btn_generar.config(text="⏳ Procesando...", state="disabled")
        root.update()
        
        if ruta_mensual:
            df = leer_csv_robusto(ruta_mensual)
            fecha_max = df['FechaCompleta'].max()
            df['Dia'] = df['FechaCompleta'].dt.day
            df['Hora'] = df['FechaCompleta'].dt.strftime('%H:%M')
            pivot = df.pivot_table(index=['Cuenta_Num', 'Cuenta', 'Nombre del Cliente'], columns='Dia', values='Hora', aggfunc='max')
            pivot = pivot.sort_index(level=0).droplevel(0).fillna('/')
            
            out = os.path.join(os.path.dirname(ruta_mensual), f"1_Reporte_Asistencia_{MESES_ES[fecha_max.month]}_{fecha_max.year}.xlsx")
            pivot.to_excel(out)
            aplicar_formato_excel(out, "asistencia")
            
        if ruta_anual:
            df_a = leer_csv_robusto(ruta_anual)
            fecha_ref_a = df_a['FechaCompleta'].max()
            u = df_a.groupby(['Cuenta_Num', 'Cuenta', 'Nombre del Cliente'])['FechaCompleta'].max().reset_index()
            u['Dias'] = (fecha_ref_a - u['FechaCompleta']).dt.days
            f = u[u['Dias'] > 4].copy()
            
            def detectar_test(cta):
                has_88 = df_a[df_a['Cuenta'] == cta]['Alarma'].astype(str).str.contains('88').any()
                return "" if has_88 else "Solo sucesos"

            f['Notas'] = f['Cuenta'].apply(detectar_test)
            f['Desde'] = (f['FechaCompleta'] + pd.Timedelta(days=1)).dt.strftime('%d/%m/%Y')
            f_ord = f.sort_values('Cuenta_Num')
            
            df_sin = f_ord[f_ord['Notas'] == ''][['Cuenta', 'Nombre del Cliente', 'Desde', 'Notas']]
            df_suc = f_ord[f_ord['Notas'] == 'Solo sucesos'][['Cuenta', 'Nombre del Cliente', 'Desde', 'Notas']]
            
            out_a = os.path.join(os.path.dirname(ruta_anual), f"2_Reporte_Fallas_Anual_{fecha_ref_a.year}.xlsx")
            with pd.ExcelWriter(out_a, engine='openpyxl') as writer:
                df_sin.to_excel(writer, sheet_name='Sin_Reportar', index=False)
                df_suc.to_excel(writer, sheet_name='Solo_Sucesos', index=False)
            aplicar_formato_excel(out_a, "fallas")
            
        messagebox.showinfo("¡Éxito!", "Reportes generados correctamente.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un problema: {e}")
    
    finally:
        btn_generar.config(text="GENERATE REPORTS", state="normal")

# --- INTERFAZ GRÁFICA ---

def sel_m(): 
    global ruta_mensual
    r = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if r: ruta_mensual = r; lbl_m.config(text=os.path.basename(r), fg="#2e7d32")

def sel_a(): 
    global ruta_anual
    r = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if r: ruta_anual = r; lbl_a.config(text=os.path.basename(r), fg="#2e7d32")

def limpiar():
    global ruta_mensual, ruta_anual
    ruta_mensual = ""; ruta_anual = ""
    lbl_m.config(text="(vacio)", fg="gray"); lbl_a.config(text="(vacio)", fg="gray")

root = tk.Tk()
root.title(f"Securithor Automator {VERSION_ACTUAL}")
root.geometry("480x620")
root.resizable(False, False)

tk.Label(root, text="System Report Generator", font=("Arial", 16, "bold")).pack(pady=15)

f1 = tk.LabelFrame(root, text=" 1. Monthly Attendance ", padx=10, pady=10)
f1.pack(pady=5, padx=20, fill="x")
tk.Button(f1, text="Select CSV", command=sel_m).pack(); lbl_m = tk.Label(f1, text="(vacio)", fg="gray"); lbl_m.pack()

f2 = tk.LabelFrame(root, text=" 2. Annual Failures ", padx=10, pady=10)
f2.pack(pady=5, padx=20, fill="x")
tk.Button(f2, text="Select CSV", command=sel_a).pack(); lbl_a = tk.Label(f2, text="(vacio)", fg="gray"); lbl_a.pack()

fr = tk.Frame(root); fr.pack(pady=10)
tk.Button(fr, text="🧹 Clear", command=limpiar, bg="#fffde7", width=10).grid(row=0, column=0, padx=5)
tk.Button(fr, text="📝 Bajas", command=lambda: abrir_txt("bajas.txt"), bg="#ffebee", width=10).grid(row=0, column=1, padx=5)
tk.Button(fr, text="👤 Names", command=lambda: abrir_txt("nombres.txt"), bg="#e3f2fd", width=10).grid(row=0, column=2, padx=5)

# Botón modificado: Mantiene el check de versión pero redirige a la web
btn_update = tk.Button(root, text="🔄 Check for Updates", command=actualizar_programa, bg="#eeeeee", font=("Arial", 9))
btn_update.pack(pady=5)

btn_generar = tk.Button(root, text="GENERATE REPORTS", command=generar_reportes, bg="#2e7d32", fg="white", font=("Arial", 12, "bold"), height=2, width=25, cursor="hand2")
btn_generar.pack(pady=20)

asegurar_archivos_config()
root.mainloop()