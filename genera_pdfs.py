import os
from pyexcelerate import Workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Carpeta con los archivos Excel generados
carpeta_excel = 'archivos_generados'
carpeta_pdf = 'archivos_pdf'

# Crear carpeta PDF si no existe
os.makedirs(carpeta_pdf, exist_ok=True)


def excel_to_pdf_multiplatform(input_excel_path, output_pdf_path):
    # Leer el archivo Excel
    wb = Workbook()
    wb.read(input_excel_path)
    ws = wb.get_sheet(1)  # Obtener la primera hoja
    
    # Crear PDF
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    width, height = letter
    
    # Configuración de texto
    text = c.beginText(40, height - 40)
    text.setFont("Helvetica", 10)
    
    # Escribir datos de Excel en PDF
    for row in ws.Range(ws.getUsedRange()):
        line = " | ".join([str(cell.Value) if cell.Value else "" for cell in row])
        text.textLine(line)
    
    c.drawText(text)
    c.save()



# Convertir todos los archivos Excel a PDF
for archivo in os.listdir(carpeta_excel):
    if archivo.endswith('.xlsx'):
        excel_path = os.path.join(carpeta_excel, archivo)
        pdf_name = os.path.splitext(archivo)[0] + '.pdf'
        pdf_path = os.path.join(carpeta_pdf, pdf_name)
        
        excel_to_pdf_multiplatform(excel_path, pdf_path)
        print(f"Convertido: {archivo} → {pdf_name}")

print("Conversión completada!")
