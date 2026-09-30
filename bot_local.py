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

def corregir_fecha_string(val):
    """Corrige cadenas de fecha d/m/yyyy, m/d/yyyy o yyyy-mm-dd para forzar dd/mm/yyyy"""
    if pd.isnull(val) or not str(val).strip():
        return ''
    val = str(val).strip()
    
    # Manejar YYYY-MM-DD o YYYY/MM/DD
    if '-' in val and len(val.split('-')[0]) == 4:
        partes = val.split('-')
        return f"{partes[2].zfill(2)}/{partes[1].zfill(2)}/{partes[0]}"
    
    # Manejar D/M/YYYY o M/D/YYYY
    if '/' in val:
        partes = val.split('/')
        if len(partes) == 3:
            p1, p2, p3 = partes[0].zfill(2), partes[1].zfill(2), partes[2]
            # Si el archivo viene como M/D/YYYY (mes 10, días 1, 2, 3...)
            # donde p1 es 10 (octubre) y p2 son los días (01, 02, 03...):
            if int(p1) > 12: # p1 es el día seguro
                return f"{p1}/{p2}/{p3}"
            elif int(p2) > 12: # p2 es el día seguro
                return f"{p2}/{p1}/{p3}"
            else:
                # Si ambos <= 12, se asume que p1 es el día si viene en formato ES
                return f"{p1}/{p2}/{p3}"
    return val

def procesar_excel_y_github(ruta_txt):
    """Estructura la Hoja 2 asegurando que las fechas se inyecten estrictamente como texto DD/MM/YYYY"""
    try:
        pythoncom.CoInitialize()

        print(f"\n📄 Convertidor activado para: {ruta_txt}")
        print("📊 Desglosando datos e inyectando en la Hoja 2...")

        # 1. Leer el TXT como texto puro
        df_txt = pd.read_csv(ruta_txt, sep='|', encoding='latin1', on_bad_lines='skip', dtype=str)
        df_txt.columns = df_txt.columns.str.strip()

        # Buscar la columna de fecha (PERIODO o FECHA)
        col_fecha = None
        for col in ['PERIODO', 'FECHA', 'FECHA_CITA', 'FECHACITA', 'FEC_CITA']:
            if col in df_txt.columns:
                col_fecha = col
                break

        if col_fecha:
            # Parsear con formato mixto respetando dayfirst
            fechas_dt = pd.to_datetime(df_txt[col_fecha], dayfirst=True, errors='coerce')
            df_txt[col_fecha] = fechas_dt.dt.strftime('%d/%m/%Y').fillna('')

        # 2. Abrir Excel mediante win32com
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(ARCHIVO_EXCEL)
        
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

        # Configurar la columna de fecha como formato Texto (@) antes de escribir datos
        if col_fecha:
            idx_col = df_txt.columns.get_loc(col_fecha) + 1
            ws_hoja2.Columns(idx_col).NumberFormat = "@"

        # Escribir datos
        datos = df_txt.fillna('').values.tolist()
        if datos:
            filas = len(datos)
            columnas = len(datos[0])
            rango_destino = ws_hoja2.Range(ws_hoja2.Cells(2, 1), ws_hoja2.Cells(filas + 1, columnas))
            rango_destino.Value = datos

            # Forzar apóstrofe en cada celda de fecha para anular la re-interpretación de Excel
            if col_fecha:
                idx_col = df_txt.columns.get_loc(col_fecha) + 1
                fechas_lista = df_txt[col_fecha].tolist()
                for i, f_val in enumerate(fechas_lista, start=2):
                    if f_val:
                        ws_hoja2.Cells(i, idx_col).Value = f"'{f_val}"

        print("🔄 Redefiniendo origen de datos y actualizando Tablas Dinámicas en Hoja 1...")

        # 3. Actualizar Tablas Dinámicas en Hoja 1
        ws_hoja1 = wb.Worksheets("Hoja1")
        
        ult_fila = ws_hoja2.UsedRange.Rows.Count
        ult_col = ws_hoja2.UsedRange.Columns.Count
        nuevo_rango = f"'{ws_hoja2.Name}'!R1C1:R{ult_fila}C{ult_col}"

        for pt in ws_hoja1.PivotTables():
            try:
                pt.ChangePivotCache(wb.PivotCaches().Create(SourceType=1, SourceData=nuevo_rango))
                pt.RefreshTable()
            except Exception as e_pt:
                print(f"Aviso al refrescar tabla dinámica: {e_pt}")

        wb.Save()
        wb.Close(True)
        excel.Quit()

        print("✅ ¡Tablas dinámicas y datos de Hoja 1/Hoja 2 actualizados al 100%!")

        # 4. Sincronizar con GitHub
        print("🚀 Sincronizando con GitHub...")
        os.chdir(CARPETA_PROYECTO)

        subprocess.run(["git", "add", "."], check=True)
        subprocess.run(["git", "commit", "-m", "Auto-update: Formato estricto DD/MM/YYYY con apóstrofe en Excel"], check=False)
        subprocess.run(["git", "pull", "origin", "main", "--rebase"], check=False)
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