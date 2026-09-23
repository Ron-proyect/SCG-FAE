import pandas as pd
from docxtpl import DocxTemplate
import os

# Nombres de tus archivos (Asegúrate de que se llamen así)
EXCEL_FILE = 'datos.xlsx'
WORD_TEMPLATE = 'plantilla.docx'
CARPETA_SALIDA = 'REPORTES_LISTOS'

def crear_reportes():
    # Crear carpeta de salida si no existe
    if not os.path.exists(CARPETA_SALIDA):
        os.makedirs(CARPETA_SALIDA)

    try:
        # 1. Leer el Excel
        df = pd.read_excel(EXCEL_FILE)
        print(f"✅ Excel leído correctamente. Filas encontradas: {len(df)}")
        
        # 2. Generar los documentos uno por uno
        for indice, fila in df.iterrows():
            doc = DocxTemplate(WORD_TEMPLATE)
            
            # Convertimos la fila en datos que Word entiende
            contexto = fila.to_dict()
            doc.render(contexto)
            
            # Guardamos el archivo (le ponemos un número para no confundirlos)
            nombre_archivo = f"{CARPETA_SALIDA}/Reporte_Fila_{indice + 1}.docx"
            doc.save(nombre_archivo)
            print(f"✔ Generado: {nombre_archivo}")
            
        print("\n🚀 ¡PROCESO TERMINADO! Revisa la carpeta REPORTES_LISTOS")
            
    except Exception as e:
        print(f"❌ ERROR: {e}")
        print(f"Asegúrate de que '{EXCEL_FILE}' y '{WORD_TEMPLATE}' estén en la carpeta.")

if __name__ == "__main__":
    crear_reportes()