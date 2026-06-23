SCRIPT: exportar_por_asesor.py

Objetivo:
- Leer el Excel fuente.
- Agrupar por la columna "Asesor".
- Crear un Excel diferente por cada asesor.
- Cada archivo contiene únicamente las filas del asesor correspondiente.
- También genera un resumen: 00_resumen_exportacion_por_asesor.xlsx

Instalación:
    pip install pandas xlsxwriter openpyxl

Comando recomendado en PowerShell:
    python .\exportar_por_asesor.py "PROCESOS DE VENTA  (1)(1).xlsx" --sheet "Export" --output-dir "exports_por_asesor"

Notas:
- El script tolera espacios extra y mayúsculas/minúsculas al buscar la columna "Asesor".
- Si hay filas sin asesor, las exporta en un archivo llamado Sin Asesor.xlsx.
- Los nombres de archivos se limpian para evitar caracteres inválidos en Windows.
