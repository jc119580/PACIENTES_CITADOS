from flask import Flask, jsonify, send_file
from flask_cors import CORS
import os
import glob
import subprocess
import pandas as pd
import datetime
import pythoncom
import win32com.client as win32

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# Rutas del sistema
CARPETA_DESCARGAS = r"C:\Users\Usuario\Downloads"
CARPETA_PROYECTO = r"C:\Users\Usuario\Documents\JC\PACIENTES CITADOS"
ARCHIVO_EXCEL = os.path.join(CARPETA_PROYECTO, "CITADOS DE SETIEMBRE.xlsx")

def fecha_a_numero_excel(dt):
    if pd.isnull(dt):
        return None
    fecha_base = datetime.datetime(1899, 12, 30)
    delta = dt - fecha_base
    return delta.days + (delta.seconds / 86400.0)

def leer_txt_valido(ruta_txt):
    """Intenta leer el TXT en UTF-8 y si falla usa latin1. Verifica que tenga columnas de citados."""
    for enc in ['utf-8', 'latin1']:
        try:
            df = pd.read_csv(ruta_txt, sep='|', encoding=enc, on_bad_lines='skip', dtype=str)
            df.columns = df.columns.str.strip()
            cols_upper = [c.upper() for c in df.columns]
            if any(c in cols_upper for c in ['PERIODO', 'FECHA', 'SUBACTIVIDAD', 'HRAS_PROG', 'CANTPACI_HRA_ATEN']):
                return df
        except Exception:
            continue
    return None

def buscar_txt_citados():
    """Busca el archivo .txt más reciente en Downloads que realmente corresponda al reporte de citados."""
    archivos_txt = glob.glob(os.path.join(CARPETA_DESCARGAS, "*.txt"))
    # Ordenar del más reciente al más antiguo
    archivos_txt.sort(key=os.path.getmtime, reverse=True)
    for ruta in archivos_txt:
        df = leer_txt_valido(ruta)
        if df is not None and len(df) > 0:
            return ruta, df
    return None, None

def procesar_excel_y_github(ruta_txt, df_txt):
    excel = None
    wb = None
    try:
        pythoncom.CoInitialize()
        print(f"\n📄 Convertidor activado para: {ruta_txt}")

        # 1. Corregir errores de codificación en Ñ y tildes (ej. NIÃ‘O -> NIÑO)
        reemplazos = {
            r"Ã‘": "Ñ",
            r"Ã±": "ñ",
            r"NIÃ['´‘]?O": "NIÑO",
            r"Ã¡": "á", r"Ã©": "é", r"Ã­": "í", r"Ã³": "ó", r"Ãº": "ú",
            r"Ã": "Á", r"Ã‰": "É", r"Ã": "Í", r"Ã“": "Ó", r"Ãš": "Ú"
        }
        df_txt = df_txt.replace(reemplazos, regex=True)

        # Identificar la columna de fecha
        col_fecha = None
        for col in ['PERIODO', 'FECHA', 'FECHA_CITA', 'FECHACITA', 'FEC_CITA']:
            if col in df_txt.columns:
                col_fecha = col
                break

        # 2. Abrir Excel mediante win32com
        excel = win32.DispatchEx('Excel.Application')
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

        datos = df_txt.fillna('').values.tolist()
        
        if datos:
            filas = len(datos)
            columnas = len(datos[0])
            rango_destino = ws_hoja2.Range(ws_hoja2.Cells(2, 1), ws_hoja2.Cells(filas + 1, columnas))
            rango_destino.Value = datos

            if col_fecha:
                idx_col = df_txt.columns.get_loc(col_fecha) + 1
                col_rango = ws_hoja2.Range(ws_hoja2.Cells(2, idx_col), ws_hoja2.Cells(filas + 1, idx_col))
                col_rango.NumberFormat = "dd/mm/yyyy"

                fechas_raw = df_txt[col_fecha].tolist()
                for i, val in enumerate(fechas_raw, start=2):
                    if val and str(val).strip():
                        try:
                            dt = pd.to_datetime(str(val).strip(), dayfirst=True, errors='coerce')
                            if pd.notnull(dt):
                                ws_hoja2.Cells(i, idx_col).Value2 = fecha_a_numero_excel(dt)
                        except Exception:
                            pass

        print("🔄 Actualizando Tablas Dinámicas en Hoja 1...")
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
        wb = None
        excel = None

        print("✅ ¡Excel actualizado!")

        # 3. Sincronizar con GitHub
        os.chdir(CARPETA_PROYECTO)
        subprocess.run(["git", "add", "."], check=True)
        subprocess.run(["git", "commit", "-m", "Auto-update: Limpieza UTF-8 (NIÑO) y fechas Value2"], check=False)
        subprocess.run(["git", "push", "origin", "main", "--force"], check=True)

        return True, "Sincronizado correctamente."

    except Exception as e:
        print(f"❌ Error durante el procesamiento: {e}")
        try:
            if wb: wb.Close(False)
            if excel: excel.Quit()
        except:
            pass
        return False, str(e)

@app.route('/api/sincronizar', methods=['POST'])
def sincronizar():
    print("\n--------------------------------------------------")
    print("🔄 Petición de sincronización recibida...")

    ruta_txt, df_txt = buscar_txt_citados()
    if not ruta_txt:
        return jsonify({"error": "No se encontró ningún archivo .txt válido de citados en Downloads."}), 400

    exito, mensaje = procesar_excel_y_github(ruta_txt, df_txt)

    if exito:
        return send_file(ARCHIVO_EXCEL, as_attachment=True)
    else:
        return jsonify({"error": mensaje}), 500

if __name__ == '__main__':
    print("🚀 Servidor local activo en http://127.0.0.1:5000")
    app.run(host='127.0.0.1', port=5000, debug=False)