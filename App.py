import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import os
import sys
import ctypes
DOMINIOS_PERMITIDOS = {
    "gmail.com", "outlook.com", "outlook.cl", "hotmail.com", 
    "hotmail.cl", "live.com", "msn.com", "yahoo.com", "yahoo.es", "yahoo.cl","icloud.com"
}

try:
    mi_app_id = 'miempresa.limpiadorcorreos.v1'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(mi_app_id)
except Exception:
    pass

def obtener_ruta_recurso(ruta_relativa):
    """Obtiene la ruta absoluta para recursos, funciona en dev y con PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, ruta_relativa)
    return os.path.join(os.path.abspath("."), ruta_relativa)

def limpiar_y_validar_correo(correo):
    if pd.isna(correo) or str(correo).strip() == "": 
        return pd.Series([str(correo), "", False, "Celda vacía"])
        
    correo_limpio = str(correo).strip().lower()
    
    patron = r'^[\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,}$'
    if not re.match(patron, correo_limpio):
        return pd.Series([str(correo), correo_limpio, False, "Estructura incorrecta"])
        
    partes = correo_limpio.split('@')
    dominio = partes[1] if len(partes) == 2 else ""
    
    if dominio not in DOMINIOS_PERMITIDOS:
        return pd.Series([str(correo), correo_limpio, False, "Dominio no permitido"])

    return pd.Series([str(correo), correo_limpio, True, "Válido"])

def seleccionar_archivo_entrada():
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel de entrada",
        filetypes=[("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
    )
    if archivo:
        var_ruta_entrada.set(archivo)
        lbl_estado.config(text="Archivo de entrada cargado.", fg="#2e7d32")

def seleccionar_archivo_salida():
    archivo = filedialog.asksaveasfilename(
        title="Guardar archivo de correos válidos como...",
        defaultextension=".xlsx",
        initialfile=var_nombre_salida.get() if var_nombre_salida.get() else "correos_validos.xlsx",
        filetypes=[("Archivos de Excel", "*.xlsx")]
    )
    if archivo:
        var_nombre_salida.set(archivo)

def procesar():
    archivo_entrada = var_ruta_entrada.get()
    if not archivo_entrada:
        messagebox.showwarning("Atención", "Por favor selecciona un archivo Excel (.xlsx) de entrada primero.")
        return

    nombre_salida = var_nombre_salida.get().strip()
    if not nombre_salida:
        nombre_salida = "correos_validos.xlsx"

    if not nombre_salida.lower().endswith(".xlsx"):
        nombre_salida += ".xlsx"

    lbl_estado.config(text="Procesando archivo, por favor espera...", fg="#1565c0")
    root.update_idletasks()

    try:
        df = pd.read_excel(archivo_entrada)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo Excel:\n{e}")
        lbl_estado.config(text="Error al abrir el archivo.", fg="#c62828")
        return

    columna_objetivo = None
    for col in df.columns:
        if str(col).lower() in ['correo', 'correos', 'email', 'emails', 'mail']:
            columna_objetivo = col
            break
            
    if not columna_objetivo:
        messagebox.showerror("Error", "No se encontró una columna llamada 'Correo' o 'Email' en la primera fila del Excel.")
        lbl_estado.config(text="Error: Columna de correo no encontrada.", fg="#c62828")
        return

    # 1. Limpieza y validación
    df[['correo_original', 'correo_limpio', 'es_valido', 'motivo']] = df[columna_objetivo].apply(limpiar_y_validar_correo)

    df_validos = df[df['es_valido'] == True].copy()
    df_invalidos = df[df['es_valido'] == False].copy()

    # 2. Detección y eliminación de duplicados en válidos
    total_validos_antes = len(df_validos)
    df_validos = df_validos.drop_duplicates(subset=['correo_limpio'], keep='first')
    duplicados_eliminados = total_validos_antes - len(df_validos)

    # 3. Reestructurar columnas: rut | nombre | correo
    df_validos_final = pd.DataFrame({
        'rut': 'No informado',
        'nombre': 'Estimado cliente',
        'correo': df_validos['correo_limpio']
    })

    df_invalidos_final = pd.DataFrame({
        'rut': 'No informado',
        'nombre': 'Estimado Cliente',
        'correo': df_invalidos['correo_limpio'],
        'motivo': df_invalidos['motivo']
    })

    # 4. Definir rutas de salida
    carpeta_origen = os.path.dirname(archivo_entrada)
    
    if os.path.isabs(nombre_salida):
        ruta_validos = nombre_salida
        carpeta_destino = os.path.dirname(nombre_salida)
    else:
        ruta_validos = os.path.join(carpeta_origen, nombre_salida)
        carpeta_destino = carpeta_origen

    # Nombre para el archivo de rechazados en la misma carpeta de salida
    base_nombre_salida = os.path.splitext(os.path.basename(ruta_validos))[0]
    ruta_invalidos = os.path.join(carpeta_destino, f"{base_nombre_salida}_rechazados.xlsx")

    try:
        # Exportar Excel de válidos
        df_validos_final.to_excel(ruta_validos, index=False)

        # Exportar Excel de rechazados
        with pd.ExcelWriter(ruta_invalidos, engine='openpyxl') as writer:
            if not df_invalidos_final.empty:
                df_invalidos_final.to_excel(writer, sheet_name='Todos los Rechazados', index=False)
                grupos = df_invalidos_final.groupby('motivo')
                for motivo, sub_df in grupos:
                    nombre_hoja = str(motivo)[:31]
                    sub_df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            else:
                pd.DataFrame({'rut': [], 'nombre': [], 'correo': []}).to_excel(writer, sheet_name='Sin Rechazos', index=False)

    except PermissionError:
        messagebox.showerror("Error de permisos", "Cierra los archivos Excel si los tienes abiertos e intenta de nuevo.")
        lbl_estado.config(text="Error: Permiso denegado al guardar.", fg="#c62828")
        return
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al guardar los archivos:\n{e}")
        lbl_estado.config(text="Error al guardar resultados.", fg="#c62828")
        return

    resumen = (
        f"✅ PROCESO COMPLETADO CON ÉXITO\n\n"
        f"• Total de registros analizados: {len(df)}\n"
        f"• Duplicados eliminados: {duplicados_eliminados}\n"
        f"• Correos válidos guardados: {len(df_validos_final)}\n"
        f"• Correos rechazados: {len(df_invalidos_final)}\n\n"
        f"Archivos generados:\n"
        f"1. {os.path.basename(ruta_validos)}\n"
        f"2. {os.path.basename(ruta_invalidos)}"
    )
    
    lbl_estado.config(text="¡Proceso finalizado correctamente!", fg="#2e7d32")
    messagebox.showinfo("Limpieza Completada", resumen)

# --- INTERFAZ GRÁFICA TKINTER ---
root = tk.Tk()
root.title("Limpiador de Correos Excel")
root.geometry("560x340")
root.resizable(False, False)

# Cargar icono de la ventana (evita errores si el archivo no está presente)
try:
    ruta_icono = obtener_ruta_recurso("icono.ico")
    root.iconbitmap(ruta_icono)
except Exception:
    pass

frame = tk.Frame(root, padx=20, pady=15)
frame.pack(fill="both", expand=True)

lbl_titulo = tk.Label(frame, text="Limpiador y Validador de Correos", font=("Segoe UI", 13, "bold"))
lbl_titulo.pack(pady=(0, 10))

# 1. Archivo de Entrada
lbl_entrada = tk.Label(frame, text="1. Archivo Excel de entrada (.xlsx):", font=("Segoe UI", 9, "bold"))
lbl_entrada.pack(anchor="w")

frame_entrada = tk.Frame(frame)
frame_entrada.pack(fill="x", pady=(2, 10))

var_ruta_entrada = tk.StringVar()
entry_entrada = tk.Entry(frame_entrada, textvariable=var_ruta_entrada, state="readonly", font=("Segoe UI", 9))
entry_entrada.pack(side="left", fill="x", expand=True, padx=(0, 8))

btn_entrada = tk.Button(frame_entrada, text="Buscar...", command=seleccionar_archivo_entrada, font=("Segoe UI", 9), width=12)
btn_entrada.pack(side="right")

# 2. Nombre o Ruta del Archivo de Salida
lbl_salida = tk.Label(frame, text="2. Nombre del archivo de salida (Válidos):", font=("Segoe UI", 9, "bold"))
lbl_salida.pack(anchor="w")

frame_salida = tk.Frame(frame)
frame_salida.pack(fill="x", pady=(2, 10))

var_nombre_salida = tk.StringVar(value="correos_validos.xlsx")
entry_salida = tk.Entry(frame_salida, textvariable=var_nombre_salida, font=("Segoe UI", 9))
entry_salida.pack(side="left", fill="x", expand=True, padx=(0, 8))

btn_salida = tk.Button(frame_salida, text="Guardar como...", command=seleccionar_archivo_salida, font=("Segoe UI", 9), width=12)
btn_salida.pack(side="right")

# Botón principal
btn_procesar = tk.Button(
    frame, 
    text="🚀 Procesar y Limpiar Excel", 
    command=procesar, 
    bg="#0078d4", 
    fg="white", 
    font=("Segoe UI", 11, "bold"), 
    relief="flat",
    cursor="hand2",
    pady=7
)
btn_procesar.pack(fill="x", pady=15)

lbl_estado = tk.Label(frame, text="Esperando selección de archivo...", font=("Segoe UI", 9, "italic"), fg="gray")
lbl_estado.pack()

if __name__ == "__main__":
    root.mainloop()