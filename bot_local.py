from flask import Flask, jsonify, send_file
from flask_cors import CORS
import os
import glob
import subprocess
import pandas as pd
import pythoncom
import win32com.client as win32

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# Rutas del sistema
CARPETA_DESCARGAS = r"C:\Users\Usuario\Downloads"
CARPETA_PROYECTO = r"C:\Users\Usuario\Documents\JC\PACIENTES CITADOS"
ARCHIVO_EXCEL = os.path.join(CARPETA_PROYECTO, "CITADOS DE SETIEMBRE.xlsx")

def procesar_excel_y_github(ruta_txt):
    """Estructura la Hoja 2, formatea fechas, actualiza el rango de origen y refresca las tablas dinámicas de la Hoja 1"""
    try:
        pythoncom.CoInitialize()

        print(f"\n📄 Convertidor activado para: {ruta_txt}")
        print("📊 Desglosando datos e inyectando en la Hoja 2...")

        # 1. Leer el TXT separando por '|'
        df_txt = pd.read_csv(ruta_txt, sep='|', encoding='latin1', on_bad_lines='skip', dtype=str)

        # Normalizar nombres de columnas eliminando espacios extras
        df_txt.columns = df_txt.columns.str.strip()

        # Corregir y formatear columna de fecha para evitar confusión entre 09/03 y 09/09
        col_fecha = None
        for col in ['FECHA', 'FECHA_CITA', 'FECHACITA', 'FEC_CITA']:
            if col in df_txt.columns:
                col_fecha = col
                break

        if col_fecha:
            # Convertir a datetime respetando el día primero (DD/MM/YYYY)
            df_txt[col_fecha] = pd.to_datetime(df_txt[col_fecha], dayfirst=True, errors='coerce')
            # Formatear estrictamente con 2 dígitos en día y mes (ej: 03/09/2026)
            df_txt[col_fecha] = df_txt[col_fecha].dt.strftime('%d/%m/%Y')

        # 2. Abrir Excel mediante win32com
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(ARCHIVO_EXCEL)
        
        # Buscar o limpiar Hoja 2
        try:
            ws_hoja2 = wb.Worksheets("Hoja2")
        except:
            try:
                ws_hoja2 = wb.Worksheets("Hoja 2")
            except:
                ws_hoja2 = wb.Worksheets(2)

        ws_hoja2.Cells.ClearContents()

        # Escribir encabezados
        for col_num, col_name in enumerate(df_txt.columns, 1):
            ws_hoja2.Cells(1, col_num).Value = str(col_name)

        # Escribir datos
        datos = df_txt.fillna('').values.tolist()
        if datos:
            filas = len(datos)
            columnas = len(datos[0])
            rango_destino = ws_hoja2.Range(ws_hoja2.Cells(2, 1), ws_hoja2.Cells(filas + 1, columnas))
            rango_destino.Value = datos

        print("🔄 Redefiniendo origen de datos y actualizando Tablas Dinámicas en Hoja 1...")

        # 3. Actualizar Tablas Dinámicas en Hoja 1
        ws_hoja1 = wb.Worksheets("Hoja1")
        
        # Determinar nuevo rango de la Hoja 2
        ult_fila = ws_hoja2.UsedRange.Rows.Count
        ult_col = ws_hoja2.UsedRange.Columns.Count
        nuevo_rango = f"'{ws_hoja2.Name}'!R1C1:R{ult_fila}C{ult_col}"

        # Recorrer y actualizar cada tabla dinámica en Hoja 1
        for pt in ws_hoja1.PivotTables():
            try:
                pt.ChangePivotCache(wb.PivotCaches().Create(SourceType=win32.constants.xlDatabase, SourceData=nuevo_rango))
                pt.RefreshTable()
            except Exception as e_pt:
                print(f"Aviso al refrescar tabla dinámica: {e_pt}")

        wb.Save()
        wb.Close(True)
        excel.Quit()

        print("✅ ¡Tablas dinámicas y datos de Hoja 1/Hoja 2 actualizados al 100%!")

        # 4. Git Pull / Rebase / Add / Commit / Push
        print("🚀 Sincronizando con GitHub...")
        os.chdir(CARPETA_PROYECTO)

        subprocess.run(["git", "pull", "origin", "main", "--rebase"], check=False)
        subprocess.run(["git", "add", "."], check=True)
        subprocess.run(["git", "commit", "-m", "Auto-update: Formato de fechas, Rango de tabla dinámica y Hoja 2 actualizados"], check=False)
        subprocess.run(["git", "push", "origin", "main"], check=True)

        print("✅ ¡Publicado con éxito en GitHub!")
        return True, "Proceso completado e interfaz sincronizada."

    except Exception as e:
        print(f"❌ Error durante el procesamiento: {e}")
        try:
            wb.Close(False)
            excel.Quit()
        except:
            pass
        return False, str(e)

@app.route('/api/sincronizar', methods=['POST'])
def sincronizar():
    print("\n--------------------------------------------------")
    print("🔄 Petición de sincronización recibida...")

    # Buscar el archivo TXT más reciente en Downloads
    patron = os.path.join(CARPETA_DESCARGAS, "*.txt")
    archivos_txt = glob.glob(patron)

    if not archivos_txt:
        return jsonify({"error": "No se encontraron archivos .txt en la carpeta Downloads"}), 400

    archivo_mas_reciente = max(archivos_txt, key=os.path.getmtime)

    exito, mensaje = procesar_excel_y_github(archivo_mas_reciente)

    if exito:
        return send_file(ARCHIVO_EXCEL, as_attachment=True)
    else:
        return jsonify({"error": mensaje}), 500

if __name__ == '__main__':
    print("🚀 Servidor local activo en http://127.0.0.1:5000")
    app.run(host='127.0.0.1', port=5000, debug=False)