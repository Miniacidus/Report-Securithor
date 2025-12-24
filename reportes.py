import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import subprocess
import io
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

# --- CONFIGURACIÓN ---
VERSION = "v5.2 Mes y Año Automático"
FILAS_POR_PAGINA = 50 
ruta_mensual = ""
ruta_anual = ""

# Diccionario para nombres de meses en español
MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# --- GESTIÓN DE ARCHIVOS ---
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
        if len(contenido) < 10:
            with open(ruta, 'r', encoding='latin-1', errors='replace') as f:
                contenido = f.read()
        df = pd.read_csv(io.StringIO(contenido), header=3)
        df.columns = df.columns.str.strip()
        posibles = ['Llegada', 'Recibido', 'Fecha', 'Arrival', 'Received']
        col_f = next((c for c in posibles if c in df.columns), None)
        if not col_f: raise Exception(f"No se halló la columna de fecha.")
        df = df.rename(columns={col_f: 'Llegada'})
        df['FechaCompleta'] = df['Llegada'].apply(corregir_fecha)
        df = df.dropna(subset=['FechaCompleta'])
        df['Cuenta'] = df['Cuenta'].astype(str).str.replace(r'\.0$', '', regex=True)
        df['Cuenta_Num'] = pd.to_numeric(df['Cuenta'], errors='coerce').fillna(0)
        bajas = cargar_bajas(); df = df[~df['Cuenta'].isin(bajas)]
        mapa = cargar_diccionario_nombres(); df['Nombre del Cliente'] = df['Cuenta'].map(mapa).fillna('')
        return df
    except Exception as e: raise Exception(f"Error procesando {os.path.basename(ruta)}:\n{e}")

def aplicar_formato_excel(ruta_excel, modo="asistencia"):
    wb = load_workbook(ruta_excel)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        borde = Side(style="thin")
        color_fondo = "4472C4"
        if sheet_name == "Sin_Reportar": color_fondo = "FFC000"
        if sheet_name == "Solo_Sucesos": color_fondo = "9BC2E6"
        for row in ws.iter_rows():
            for cell in row:
                cell.border = Border(left=borde, right=borde, top=borde, bottom=borde)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if cell.row == 1:
                    cell.font = Font(bold=True, color="FFFFFF", size=10)
                    cell.fill = PatternFill("solid", fgColor=color_fondo)
                else: cell.font = Font(size=9)
        ws.column_dimensions['A'].width = 12 
        ws.column_dimensions['B'].width = 45 
        ws.column_dimensions['C'].width = 20
        if ws.max_column >= 4: ws.column_dimensions['D'].width = 25
        if modo == "asistencia":
            for col in range(3, ws.max_column + 1): ws.column_dimensions[get_column_letter(col)].width = 6
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.print_title_rows = '1:1'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        for i in range(FILAS_POR_PAGINA + 2, ws.max_row + 1, FILAS_POR_PAGINA): ws.row_breaks.append(Break(id=i-1))
    wb.save(ruta_excel)

def generar_reportes():
    global ruta_mensual, ruta_anual
    if not ruta_mensual and not ruta_anual:
        messagebox.showwarning("Atención", "Selecciona archivos CSV.")
        return
    try:
        btn_generar.config(text="⏳ Procesando...", state="disabled")
        root.update()
        
        if ruta_mensual:
            df = leer_csv_robusto(ruta_mensual)
            # Obtener mes y año para el nombre del archivo
            fecha_ref = df['FechaCompleta'].max()
            mes_nombre = MESES_ES[fecha_ref.month]
            anio_val = fecha_ref.year
            
            df['Dia'] = df['FechaCompleta'].dt.day
            df['Hora'] = df['FechaCompleta'].dt.strftime('%H:%M')
            pivot = df.pivot_table(index=['Cuenta_Num', 'Cuenta', 'Nombre del Cliente'], columns='Dia', values='Hora', aggfunc='max')
            pivot = pivot.sort_index(level=0).droplevel(0).fillna('/')
            
            # --- NOMBRE PERSONALIZADO ---
            nombre_archivo = f"1_Reporte_Asistencia_{mes_nombre}_{anio_val}.xlsx"
            out = os.path.join(os.path.dirname(ruta_mensual), nombre_archivo)
            pivot.to_excel(out)
            aplicar_formato_excel(out, "asistencia")
            
        if ruta_anual:
            df_a = leer_csv_robusto(ruta_anual)
            fecha_ref_a = df_a['FechaCompleta'].max()
            anio_val_a = fecha_ref_a.year
            
            u = df_a.groupby(['Cuenta_Num', 'Cuenta', 'Nombre del Cliente'])['FechaCompleta'].max().reset_index()
            u['Dias'] = (fecha_ref_a - u['FechaCompleta']).dt.days
            f = u[u['Dias'] > 4].copy()
            def detectar_nota(cta):
                tiene_test = df_a[df_a['Cuenta'] == cta]['Alarma'].astype(str).str.contains('88').any()
                return "" if tiene_test else "Solo sucesos"
            f['Notas'] = f['Cuenta'].apply(detectar_nota)
            f['Desde'] = (f['FechaCompleta'] + pd.Timedelta(days=1)).dt.strftime('%d/%m/%Y')
            f_ord = f.sort_values('Cuenta_Num')
            df_sin = f_ord[f_ord['Notas'] == ''][['Cuenta', 'Nombre del Cliente', 'Desde', 'Notas']]
            df_suc = f_ord[f_ord['Notas'] == 'Solo sucesos'][['Cuenta', 'Nombre del Cliente', 'Desde', 'Notas']]
            
            # --- NOMBRE PERSONALIZADO ---
            nombre_archivo_a = f"2_Reporte_Fallas_Anual_{anio_val_a}.xlsx"
            out_a = os.path.join(os.path.dirname(ruta_anual), nombre_archivo_a)
            with pd.ExcelWriter(out_a, engine='openpyxl') as writer:
                df_sin.to_excel(writer, sheet_name='Sin_Reportar', index=False)
                df_suc.to_excel(writer, sheet_name='Solo_Sucesos', index=False)
            aplicar_formato_excel(out_a, "fallas")
            
        messagebox.showinfo("¡Éxito!", "Reportes generados con fecha en el nombre.")
    except Exception as e: messagebox.showerror("Error", str(e))
    finally: btn_generar.config(text="GENERAR REPORTES", state="normal")

# --- INTERFAZ ---
def sel_m(): 
    global ruta_mensual
    r = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
    if r: ruta_mensual = r; lbl_m.config(text=os.path.basename(r), fg="green")
def sel_a(): 
    global ruta_anual
    r = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
    if r: ruta_anual = r; lbl_a.config(text=os.path.basename(r), fg="green")
def limpiar():
    global ruta_mensual, ruta_anual
    ruta_mensual = ""; ruta_anual = ""; lbl_m.config(text="(vacio)", fg="gray"); lbl_a.config(text="(vacio)", fg="gray")

root = tk.Tk()
root.title(f"Securithor {VERSION}"); root.geometry("480x580"); root.resizable(False, False)
tk.Label(root, text="Generador de Reportes", font=("Arial", 16, "bold")).pack(pady=20)
f1 = tk.LabelFrame(root, text=" 1. Asistencia (Mensual) ", padx=10, pady=10); f1.pack(pady=10, padx=20, fill="x")
tk.Button(f1, text="Cargar CSV", command=sel_m).pack(); lbl_m = tk.Label(f1, text="(vacio)", fg="gray"); lbl_m.pack()
f2 = tk.LabelFrame(root, text=" 2. Fallas (Anual) ", padx=10, pady=10); f2.pack(pady=10, padx=20, fill="x")
tk.Button(f2, text="Cargar CSV", command=sel_a).pack(); lbl_a = tk.Label(f2, text="(vacio)", fg="gray"); lbl_a.pack()
fr = tk.Frame(root); fr.pack(pady=10)
tk.Button(fr, text="🧹 Limpiar", command=limpiar, bg="#fff9c4", width=10).grid(row=0, column=0, padx=5)
tk.Button(fr, text="📝 Bajas", command=lambda: abrir_txt("bajas.txt"), bg="#ffcdd2", width=10).grid(row=0, column=1, padx=5)
tk.Button(fr, text="👤 Nombres", command=lambda: abrir_txt("nombres.txt"), bg="#bbdefb", width=10).grid(row=0, column=2, padx=5)
btn_generar = tk.Button(root, text="GENERAR REPORTES", command=generar_reportes, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), height=2, width=25, cursor="hand2")
btn_generar.pack(pady=20); asegurar_archivos_config(); root.mainloop()