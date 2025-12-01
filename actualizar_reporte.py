import pandas as pd
import glob
import os
import xlwings as xw
import sys

# --- CONFIGURACIÓN ---
# Aquí defines las rutas. Si cambian, solo editas estas líneas.
CARPETA_INSUMOS = 'Sucursales'
ARCHIVO_DESTINO = 'Plantilla_Dashboard.xlsm'
HOJA_DATOS = 'BD_DATA'

print("==========================================")
print("   INICIANDO ROBOT DE CONSOLIDACIÓN BANCARIA")
print("==========================================")

# ---------------------------------------------------------
# PASO 1: EXTRACT (Lectura y Consolidación)
# ---------------------------------------------------------
print(f"📂 Buscando archivos en la carpeta '{CARPETA_INSUMOS}'...")

# Buscamos todos los CSV
archivos = glob.glob(os.path.join(CARPETA_INSUMOS, "*.csv"))

if not archivos:
    print("❌ ERROR: No encontré archivos .csv en la carpeta.")
    print("   Asegúrate de haber generado los datos dummy primero.")
    sys.exit() # Detiene el programa si no hay insumos

print(f"🔎 Se encontraron {len(archivos)} archivos. Procesando...")

lista_tablas = []

for archivo in archivos:
    try:
        df_temp = pd.read_csv(archivo)
        lista_tablas.append(df_temp)
    except Exception as e:
        print(f"⚠️ Advertencia: No se pudo leer {archivo}. Razón: {e}")

# Unimos todo en un solo DataFrame
if lista_tablas:
    df_consolidado = pd.concat(lista_tablas, ignore_index=True)
    total_ventas = df_consolidado['Monto_Venta'].sum()
    print(f"✅ Consolidación exitosa.")
    print(f"   - Total Registros: {len(df_consolidado)}")
    print(f"   - Monto Total: ${total_ventas:,.0f}")
else:
    print("❌ Error: No se pudieron consolidar datos.")
    sys.exit()

# ---------------------------------------------------------
# PASO 2: LOAD (Inyección en Excel con xlwings)
# ---------------------------------------------------------
print("\n🚀 Iniciando Excel para actualizar el reporte...")

# Iniciamos Excel en modo invisible para el usuario
app = xw.App(visible=False)

try:
    # Verificamos si el archivo existe antes de abrirlo
    if not os.path.exists(ARCHIVO_DESTINO):
        print(f"❌ ERROR: No encuentro la plantilla '{ARCHIVO_DESTINO}'.")
        raise FileNotFoundError

    wb = app.books.open(ARCHIVO_DESTINO)
    
    # Seleccionamos la hoja. Si no existe, dará error.
    try:
        sheet = wb.sheets[HOJA_DATOS]
    except:
        print(f"❌ ERROR: La hoja '{HOJA_DATOS}' no existe en el Excel.")
        raise

    # Limpiamos datos viejos (desde A2 hacia abajo y hacia la derecha)
    # Usamos .api.UsedRange para asegurar limpieza profunda si es necesario,
    # pero .expand() suele ser suficiente.
    if sheet.range('A2').value is not None:
        sheet.range('A2').expand().clear_contents()
    
    print("🧹 Datos antiguos eliminados.")

    # Pegamos los nuevos datos
    # index=False (sin numero de fila), header=False (sin titulos, ya están en el Excel)
    sheet.range('A2').options(index=False, header=False).value = df_consolidado
    
    print("💉 Datos nuevos inyectados correctamente.")

    # Guardamos
    wb.save()
    print("💾 Archivo guardado.")

except Exception as e:
    print(f"❌ Ocurrió un error crítico en Excel: {e}")

finally:
    # IMPORTANTE: Siempre cerrar Excel y matar el proceso
    if 'wb' in locals():
        wb.close()
    app.quit()
    print("✅ Proceso finalizado. Excel cerrado.")

print("\n==========================================")
print("   TAREA COMPLETADA CON ÉXITO")
print("   Abre 'Plantilla_Dashboard.xlsm' y actualiza.")
print("==========================================")