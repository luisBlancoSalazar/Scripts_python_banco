import pandas as pd
from pathlib import Path
import sys
import tkinter as tk
from tkinter import messagebox

# --- CONFIGURACIÓN ---
PATRON_ARCHIVO_ENTRADA = 'Restructuraciones*.xlsx'
NOMBRE_ARCHIVO_SALIDA = 'reporte_final_reestructuraciones.xlsx'

def obtener_ruta_base():
    """
    Obtiene la ruta base del script, ya sea que se ejecute como .py o como .exe compilado.
    """
    if getattr(sys, 'frozen', False):
        # Estamos ejecutando como un .exe compilado
        return Path(sys.executable).parent
    else:
        # Estamos ejecutando como un script .py normal
        return Path(__file__).parent

def encontrar_archivo_entrada(ruta_base):
    """Busca en la ruta base un archivo que coincida con el patrón."""
    archivos_encontrados = list(ruta_base.glob(PATRON_ARCHIVO_ENTRADA))
    if not archivos_encontrados:
        return None
    return archivos_encontrados[0]

def main():
    """
    Función principal que lee, procesa y exporta los datos.
    """
    # Ocultar la ventana principal de tkinter
    root = tk.Tk()
    root.withdraw()

    try:
        ruta_script = obtener_ruta_base()
        ruta_entrada = encontrar_archivo_entrada(ruta_script)

        if ruta_entrada is None:
            messagebox.showerror("Error", f"No se encontró ningún archivo que coincida con el patrón '{PATRON_ARCHIVO_ENTRADA}'.\n\nAsegúrate de que el archivo de Excel esté en la misma carpeta que el ejecutable.")
            return

        ruta_salida = ruta_script / NOMBRE_ARCHIVO_SALIDA
        
        df = pd.read_excel(ruta_entrada)
        
        # ... (El resto de tu lógica de procesamiento es exactamente la misma)
        df['FECHA DE RESTRUCTURACION'] = pd.to_datetime(df['FECHA DE RESTRUCTURACION'], dayfirst=True)
        df_ordenado = df.sort_values(by=['NUMERO DE SOLICITUD', 'FECHA DE RESTRUCTURACION'], ascending=True)
        df_vigente = df_ordenado[df_ordenado['ESTADO DE CREDITO ACTUAL'] == 'VIGENTE'].copy()
        df_agrupado = df_vigente.groupby("NUMERO DE SOLICITUD").agg(
            AGENCIA=('AGENCIA', 'first'),
            SALDO_CREDITO=('SALDO CREDITO A LA FECHA', 'last')
        ).reset_index()
        resumen_por_agencia = df_agrupado.groupby('AGENCIA').agg(
            NUMERO_DE_REESTRUCTURACIONES=('NUMERO DE SOLICITUD', 'count'),
            SALDO_CREDITO_TOTAL=('SALDO_CREDITO', 'sum')
        ).reset_index()
        df_sin_duplicados = df_vigente.drop_duplicates(subset='NUMERO DE SOLICITUD', keep='last')

        with pd.ExcelWriter(ruta_salida) as writer:
            df_sin_duplicados.to_excel(writer, sheet_name='Detalle', index=False)
            resumen_por_agencia.to_excel(writer, sheet_name='Resumen por Agencia', index=False)
        
        messagebox.showinfo("Proceso Completado", f"¡Éxito! 🎉\n\nEl reporte '{NOMBRE_ARCHIVO_SALIDA}' ha sido generado en la misma carpeta.")

    except KeyError as e:
        messagebox.showerror("Error de Columna", f"No se encontró la columna {e}.\n\nVerifica que los nombres de las columnas en tu archivo Excel sean los correctos.")
    except Exception as e:
        messagebox.showerror("Error Inesperado", f"Ocurrió un error inesperado:\n\n{e}")

if __name__ == "__main__":
    main()