from flask import Flask, jsonify, send_file
from flask_cors import CORS
import os
import glob
import subprocess
import pandas as pd
import pythoncom
import win32com.client as win32
from playwright.sync_api import sync_playwright
from datetime import datetime

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# Credenciales y Rutas configuradas
USUARIO_ESSALUD = "42283343"
PASSWORD_ESSALUD = "CANTUARIAS8j"
CARPETA_DESCARGAS = r"C:\Users\Usuario\Downloads"
CARPETA_PROYECTO = r"C:\Users\Usuario\Documents\JC\PACIENTES CITADOS"
ARCHIVO_EXCEL = os.path.join(CARPETA_PROYECTO, "CITADOS DE SETIEMBRE.xlsx")

def descargar_reporte_essalud_headless():
    """Conecta a EsSalud, inicia sesión y descarga el .TXT automáticamente en segundo plano"""
    print("\n🤖 Iniciando automatización de descarga en SEGUNDO PLANO...")
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-web-security", "--disable-site-isolation-trials"]
        )
        context = browser.new_context(accept_downloads=True, no_viewport=True)
        page = context.new_page()
        page.set_default_timeout(180000)

        print("🌐 Conectando a EsSalud...")
        page.goto("http://appsgasistexpl.essalud.gob.pe/explotaDatos/index.html", wait_until="domcontentloaded")
        page.wait_for_timeout(3000)

        # 1. Localizar el frame de login
        frame_login = None
        for _ in range(20):
            for f in page.frames:
                try:
                    if f.locator("input[name='txtUsuario']").count() > 0:
                        frame_login = f
                        break
                except Exception:
                    continue
            if frame_login:
                break
            page.wait_for_timeout(500)

        if not frame_login:
            frame_login = page

        # 2. Iniciar sesión con tus credenciales
        print("🔑 Autenticando con credenciales...")
        frame_login.locator("input[name='txtUsuario']").first.fill(USUARIO_ESSALUD)
        frame_login.locator("input[name='txtClave']").first.fill(PASSWORD_ESSALUD)

        cas_select = frame_login.locator("select[name='cmbCas']").first
        if cas_select.count() > 0:
            try:
                cas_select.select_option(label="H.I ALBRECHT")
            except Exception:
                opts = cas_select.locator("option").all_text_contents()
                for o in opts:
                    if "ALBRECHT" in o.upper():
                        cas_select.select_option(label=o)
                        break

        frame_login.locator("input[value='Ingresar']").first.click()
        page.wait_for_timeout(5000)

        # 3. Navegar en el menú
        frame_menu = None
        for _ in range(20):
            for f in page.frames:
                try:
                    if f.locator("text=CONSULTA EXTERNA").count() > 0:
                        frame_menu = f
                        break
                except Exception:
                    continue
            if frame_menu:
                break
            page.wait_for_timeout(500)

        if not frame_menu:
            frame_menu = page

        print("📍 Seleccionando 'CONSULTA EXTERNA' > 'PROG. HORAS EFECTIVAS'...")
        frame_menu.locator("text=CONSULTA EXTERNA").first.click()
        page.wait_for_timeout(1000)
        frame_menu.locator("text=PROG. HORAS EFECTIVAS").first.click()
        page.wait_for_timeout(2500)

        # 4. Configurar fechas del mes actual y selector TXT
        hoy = datetime.now()
        fecha_inicio = f"01/{hoy.strftime('%m/%Y')}"
        ultimo_dia = 31 if hoy.month == 12 else (datetime(hoy.year, hoy.month + 1, 1) - datetime(hoy.year, hoy.month, 1)).days
        fecha_fin = f"{ultimo_dia:02d}/{hoy.strftime('%m/%Y')}"

        print(f"📅 Rango configurado: {fecha_inicio} al {fecha_fin}")
        frame_menu.locator("input[name='txtFecIni']").first.fill(fecha_inicio)
        frame_menu.locator("input[name='txtFecFin']").first.fill(fecha_fin)
        
        select_tipo = frame_menu.locator("select[name='cmbTipoArc']").first
        try:
            select_tipo.select_option(label="*.TXT para XLS")
        except Exception:
            select_tipo.select_option(value="TXT")

        page.wait_for_timeout(1500)

        # 5. Descargar reporte automáticamente
        print("📥 Disparando descarga del archivo .TXT...")
        btn_imprimir = frame_menu.locator("input[value='Imprimir']").first

        with page.expect_download(timeout=180000) as download_info:
            btn_imprimir.click(force=True)
        
        download = download_info.value
        ruta_guardada = os.path.join(CARPETA_DESCARGAS, download.suggested_filename)
        download.save_as(ruta_guardada)
        browser.close()
        
        print(f"✅ Descarga completada en segundo plano: {ruta_guardada}")
        return ruta_guardada

def procesar_excel_y_github(ruta_txt):
    """Procesa el TXT, actualiza la Hoja 2, reajusta las tablas dinámicas de la Hoja 1 y sube a GitHub"""
    try:
        pythoncom.CoInitialize()

        print(f"\n📄 Procesando archivo: {ruta_txt}")
        print("📊 Inyectando datos en la Hoja 2...")
        
        df_txt = pd.read_csv(ruta_txt, sep='|', encoding='latin1', on_bad_lines='skip', dtype=str)

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(ARCHIVO_EXCEL)
        ws_hoja1 = wb.Worksheets(1)
        ws_hoja2 = wb.Worksheets(2)

        ws_hoja2.Cells.Clear()

        headers = [str(col).strip() for col in df_txt.columns]
        filas = [[str(val).strip() if pd.notna(val) else "" for val in row] for row in df_txt.values]
        
        matriz_completa = [headers] + filas
        num_filas = len(matriz_completa)
        num_cols = len(headers)

        rango_destino = ws_hoja2.Range(
            ws_hoja2.Cells(1, 1),
            ws_hoja2.Cells(num_filas, num_cols)
        )
        rango_destino.Value = matriz_completa

        print("🔄 Redefiniendo origen y actualizando Tablas Dinámicas...")
        nombre_hoja2 = ws_hoja2.Name
        nuevo_rango_origen = f"'{nombre_hoja2}'!R1C1:R{num_filas}C{num_cols}"

        for pt in ws_hoja1.PivotTables():
            try:
                pt.ChangePivotCache(
                    wb.PivotCaches().Create(
                        SourceType=1,
                        SourceData=nuevo_rango_origen
                    )
                )
                pt.Update()
            except Exception:
                pass

        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        
        wb.Save()
        wb.Close()
        excel.Quit()
        print("✅ Excel actualizado y guardado correctamente.")

        print("🚀 Publicando cambios en GitHub...")
        subprocess.run(["git", "branch", "-M", "main"], cwd=CARPETA_PROYECTO)
        subprocess.run(["git", "add", "."], cwd=CARPETA_PROYECTO)
        subprocess.run(["git", "commit", "-m", "Auto-update: Descarga automática y tablas dinámicas sincronizadas"], cwd=CARPETA_PROYECTO)
        subprocess.run(["git", "pull", "origin", "main", "--rebase"], cwd=CARPETA_PROYECTO)
        subprocess.run(["git", "push", "-u", "origin", "main"], cwd=CARPETA_PROYECTO)
        print("✅ ¡Publicado con éxito en GitHub!")

    except Exception as e:
        print(f"⚠️ Error procesando Excel/Git: {str(e)}")
    finally:
        pythoncom.CoUninitialize()

@app.route('/api/sincronizar', methods=['POST', 'OPTIONS'])
def sincronizar():
    try:
        print("\n--------------------------------------------------")
        print("🔄 Petición de sincronización recibida...")
        
        ruta_txt = None
        try:
            # 1. Intentar la descarga automática en segundo plano
            ruta_txt = descargar_reporte_essalud_headless()
        except Exception as e_bot:
            print(f"⚠️ Descarga directa en vivo no disponible ({e_bot}). Usando el archivo .TXT más reciente...")
            archivos_txt = sorted(
                glob.glob(os.path.join(CARPETA_DESCARGAS, "*.txt")),
                key=os.path.getmtime,
                reverse=True
            )
            if archivos_txt:
                ruta_txt = archivos_txt[0]

        if ruta_txt and os.path.exists(ruta_txt):
            procesar_excel_y_github(ruta_txt)
            return send_file(ARCHIVO_EXCEL, as_attachment=True)
        else:
            return jsonify({"error": "No hay ningún archivo disponible para procesar."}), 404

    except Exception as e:
        print(f"❌ Error durante el proceso: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("🚀 Servidor local activo en http://127.0.0.1:5000")
    app.run(host='127.0.0.1', port=5000)