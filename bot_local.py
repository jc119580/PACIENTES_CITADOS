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
    """Estructura la Hoja 2, actualiza el rango de origen y refresca las tablas dinámicas de la Hoja 1"""
    try:
        pythoncom.CoInitialize()

        print(f"\n📄 Convertidor activado para: {ruta_txt}")
        print("📊 Desglosando datos e inyectando en la Hoja 2...")
        
        # 1. Leer el TXT separando por '|'
        df_txt = pd.read_csv(ruta_txt, sep='|', encoding='latin1', on_bad_lines='skip', dtype=str)

        # 2. Abrir Excel mediante win32com
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(ARCHIVO_EXCEL)
        ws_hoja1 = wb.Worksheets(1)
        ws_hoja2 = wb.Worksheets(2)

        # Limpiar datos anteriores de la Hoja 2
        ws_hoja2.Cells.Clear()

        # Preparar matriz de datos
        headers = [str(col).strip() for col in df_txt.columns]
        filas = [[str(val).strip() if pd.notna(val) else "" for val in row] for row in df_txt.values]
        
        matriz_completa = [headers] + filas
        num_filas = len(matriz_completa)
        num_cols = len(headers)

        # Inyectar datos en la Hoja 2
        rango_destino = ws_hoja2.Range(
            ws_hoja2.Cells(1, 1),
            ws_hoja2.Cells(num_filas, num_cols)
        )
        rango_destino.Value = matriz_completa

        print("🔄 Redefiniendo origen de datos y actualizando Tablas Dinámicas en Hoja 1...")
        
        # Determinar el nombre de la Hoja 2 para la referencia de rango de la tabla dinámica
        nombre_hoja2 = ws_hoja2.Name
        # Obtener la letra de la última columna
        def col_to_letter(col_idx):
            letter = ''
            while col_idx > 0:
                col_idx, remainder = divmod(col_idx - 1, 26)
                letter = chr(65 + remainder) + letter
            return letter

        ultima_col_letra = col_to_letter(num_cols)
        nuevo_rango_origen = f"'{nombre_hoja2}'!R1C1:R{num_filas}C{num_cols}"

        # Refrescar y actualizar el rango origen de cada Tabla Dinámica en la Hoja 1
        for pt in ws_hoja1.PivotTables():
            try:
                pt.ChangePivotCache(
                    wb.PivotCaches().Create(
                        SourceType=1, # xlDatabase
                        SourceData=nuevo_rango_origen
                    )
                )
                pt.Update()
            except Exception as e_pt:
                print(f"⚠️ Nota al actualizar tabla dinámica: {str(e_pt)}")

        # Refresco global de seguridad
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        
        wb.Save()
        wb.Close()
        excel.Quit()
        print("✅ ¡Tablas dinámicas y datos de Hoja 1/Hoja 2 actualizados al 100%!")

        # 3. Subir cambios a GitHub
        print("🚀 Sincronizando con GitHub...")
        subprocess.run(["git", "branch", "-M", "main"], cwd=CARPETA_PROYECTO)
        subprocess.run(["git", "add", "."], cwd=CARPETA_PROYECTO)
        subprocess.run(["git", "commit", "-m", "Auto-update: Rango de tabla dinámica y Hoja 2 actualizados"], cwd=CARPETA_PROYECTO)
        subprocess.run(["git", "pull", "origin", "main", "--rebase"], cwd=CARPETA_PROYECTO)
        subprocess.run(["git", "push", "-u", "origin", "main"], cwd=CARPETA_PROYECTO)
        print("✅ ¡Publicado con éxito en GitHub!")

    except Exception as e:
        print(f"⚠️ Error en el procesamiento: {str(e)}")
    finally:
        pythoncom.CoUninitialize()

@app.route('/api/sincronizar', methods=['POST', 'OPTIONS'])
def sincronizar():
    try:
        print("\n--------------------------------------------------")
        print("🔄 Petición de sincronización recibida...")
        
        # Buscar el archivo .TXT más reciente en Descargas
        archivos_txt = sorted(
            glob.glob(os.path.join(CARPETA_DESCARGAS, "*.txt")),
            key=os.path.getmtime,
            reverse=True
        )

        if archivos_txt:
            ruta_txt = archivos_txt[0]
            procesar_excel_y_github(ruta_txt)
            print(f"📤 Enviando archivo Excel actualizado a la web: {ARCHIVO_EXCEL}")
            return send_file(ARCHIVO_EXCEL, as_attachment=True)
        else:
            print("⚠️ No se encontró ningún archivo .TXT en la carpeta Descargas.")
            return jsonify({"error": "No hay archivos .txt en la carpeta Descargas."}), 404

    except Exception as e:
        print(f"❌ Error durante el proceso: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("🚀 Servidor local activo en http://127.0.0.1:5000")
    app.run(host='127.0.0.1', port=5000)