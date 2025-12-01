# 📊 Automatización de Reporting Financiero (ETL Pipeline)

![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![Excel](https://img.shields.io/badge/Excel-VBA-green)
![Status](https://img.shields.io/badge/Status-Completed-success)

## 💼 Descripción del Proyecto
Este proyecto resuelve un problema común en la banca y finanzas corporativas: la consolidación manual de múltiples fuentes de datos. 

Desarrollé un **Robot de Software (RPA)** que automatiza el ciclo completo de reporting:
1.  **EXTRACT:** Detecta y lee reportes de ventas fragmentados (CSV) de 7 sucursales distintas.
2.  **TRANSFORM:** Consolida y limpia más de **35,000 registros** usando Pandas.
3.  **LOAD:** Inyecta la data procesada en una plantilla corporativa de Excel (`.xlsm`) respetando macros y Dashboards existentes.

**Impacto:** Reducción del tiempo operativo de 30 minutos a **5 segundos** con 0% de error humano.

---

## 🛠️ Tecnologías Usadas
*   **Python:** Lógica central y manipulación de datos.
    *   `pandas`: Para el procesamiento masivo de datos (ETL).
    *   `xlwings`: Para la interacción avanzada con Excel y Macros.
    *   `glob` / `os`: Para la gestión dinámica de directorios de archivos.
*   **Microsoft Excel:** Frontend para el usuario final.
    *   Tablas Dinámicas conectadas a rangos dinámicos.
    *   Macros (VBA) para actualización de interfaz.

---

## 📸 Demostración Visual

### 1. El Problema (Inputs Fragmentados)
*El script detecta automáticamente n cantidad de archivos en la carpeta objetivo.*
![Input Data](PON_AQUI_EL_LINK_DE_TU_IMAGEN_1)

### 2. El Proceso (Ejecución del Script)
*Consolidación de 35.000 registros en tiempo real.*
![Console Output](PON_AQUI_EL_LINK_DE_TU_IMAGEN_2)

### 3. El Resultado (Dashboard Automatizado)
*Dashboard interactivo que se actualiza con un solo clic.*
![Dashboard Final](PON_AQUI_EL_LINK_DE_TU_IMAGEN_3)

---

## 🚀 Cómo ejecutar este proyecto

1.  **Clonar el repositorio:**
    ```bash
    git clone https://github.com/TU_USUARIO/Financial-Reporting-Automation.git
    ```
2.  **Instalar dependencias:**
    ```bash
    pip install pandas xlwings openpyxl
    ```
3.  **Generar datos de prueba (Opcional):**
    Ejecuta el script de generación para crear los archivos CSV simulados.
4.  **Ejecutar el Robot:**
    ```bash
    python actualizar_reporte.py
    ```
5.  **Ver Resultado:**
    Abre `Plantilla_Dashboard.xlsm` y haz clic en "ACTUALIZAR DATOS".

---

## 🛡️ Nota de Privacidad y Ética
**Datos Sintéticos:** Toda la información financiera, transacciones e identificaciones mostradas en este repositorio son **ficticias**. Fueron generadas algorítmicamente (`NumPy`) con fines académicos y demostrativos. Este proyecto no contiene información real de ninguna entidad financiera.

---
*Desarrollado por [Jhon Vairon Jaramillo Riascos] - Estudiante de Administración Financiera*
