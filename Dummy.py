import pandas as pd
import numpy as np
import os

# 1. Configuración
carpeta_destino = 'Sucursales'
if not os.path.exists(carpeta_destino):
    os.makedirs(carpeta_destino)



# A) Agregamos las sucursales
sucursales = ['Bogota', 'Medellin', 'Cali', 'Barranquilla', 'Bucaramanga', 'Cartagena', 'Pereira']

# B) Definimos cuántas ventas queremos por cada archivo
CANTIDAD_DATOS = 2000  

# ------------------------------

productos = ['Credito Consumo', 'Tarjeta Visa', 'Libranza', 'Microcredito']

# 2. Bucle para generar los archivos masivos
for ciudad in sucursales:
    datos = {
        'Fecha': '2025-11-27',
        'Sucursal': ciudad,
        
        # AQUÍ USAMOS LA VARIABLE CANTIDAD_DATOS
        'Producto': [np.random.choice(productos) for _ in range(CANTIDAD_DATOS)], 
        'Monto_Venta': np.random.randint(100000, 5000000, size=CANTIDAD_DATOS),   
        'ID_Asesor': np.random.randint(100, 199, size=CANTIDAD_DATOS)             
    }
    
    df_temp = pd.DataFrame(datos)
    
    nombre_archivo = f'{carpeta_destino}/Ventas_{ciudad}.csv'
    df_temp.to_csv(nombre_archivo, index=False)
    
    # Formato {:,.0f} pone los puntos de mil para leerlo mejor
    print(f" Generado: {ciudad} con {CANTIDAD_DATOS:,.0f} filas.")
