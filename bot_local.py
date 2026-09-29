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
    """Lee el TXT descargado, actualiza la Hoja 2, refresca la Hoja 1 y sube a GitHub"""
    try:
        pythoncom.CoInitialize()

        print(f"\n📄 Archivo .TXT encontrado en Descargas: {ruta_txt}")
        print("📊 Leyendo datos e inyectando en la Hoja 2...")
        
        # 1. Leer el TXT separado por tabuladores
        df_txt = pd.read_csv(ruta_txt, sep='\t', encoding='latin1', on_bad_lines='skip')

        # 2. Abrir Excel mediante la API de Windows
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(ARCHIVO_EXCEL)
        ws_hoja2 = wb.Worksheets(2)
        ws_hoja2.Cells.ClearContents() # Limpiar Hoja 2

        # Escribir Encabezados
        for col_num, col_name in enumerate(df_txt.columns, 1):
            ws_hoja2.Cells(1, col_num).Value = str(col_name)

        # Escribir Filas
        datos = df_txt.values.tolist()
        for row_idx, row_data in enumerate(datos, 2):
            for col_idx, val in enumerate(row_data, 1):
                ws_hoja2.Cells(row_idx, col_idx).Value = str(val) if pd.notna(val) else ""

        print("🔄 Recalculando y actualizando Tablas Dinámicas (Hoja 1)...")
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        
        wb.Save()
        wb.Close()
        excel.Quit()
        print("✅ Excel 'CITADOS DE SETIEMBRE.xlsx' actualizado y guardado correctamente.")

        # 3. Git Push a GitHub
        print("🚀 Sincronizando con GitHub...")
        subprocess.run(["git", "add", "."], cwd=CARPETA_PROYECTO)
        subprocess.run(["git", "commit", "-m", "Auto-update: Datos sincronizados"], cwd=CARPETA_PROYECTO)
        subprocess.run(["git", "push", "origin", "main"], cwd=CARPETA_PROYECTO)
        print("✅ ¡Publicado con éxito en GitHub!")

    except Exception as e:
        print(f"⚠️ Nota sobre Git/Excel: {str(e)}")
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