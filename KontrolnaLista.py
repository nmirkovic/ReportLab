import pandas as pd
from PyPDF2 import PdfWriter, PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

# putanja
data = pd.read_excel('C:/Users/.../Kontrolna Lista.xlsx')
data = data.fillna('')

table_style = TableStyle([
    ('ALIGN', (0, 0), (1, 3), 'RIGHT'),
    ('FONTNAME', (0, 0), (1, 3), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (1, 3), 10),
    ('BOTTOMPADDING', (0, 0), (1, 3), 5),
    # Tablica 1
    ('BOX', (0, 0), (1, 3), 2, colors.black),
    ('GRID', (1, 0), (1, 3), 1, colors.black),
    # Tablica 2 (od kupac)
    ('BOX', (2, 0), (3, 8), 2, colors.black),
    ('GRID', (2, 0), (3, 8), 1, colors.black),
    # Tablica 3 (poklopac toplate)
    ('BOX', (0, 5), (1, 8), 2, colors.black),
    # Tablica 4 (radnik, datum, potpis)
    ('GRID', (0, 11), (3, 23), 1, colors.black),
    ('GRID', (1, 10), (3, 10), 1, colors.black),
    # Tablica 5 (Zavrsna kontrola)
    ('BOX', (0, 25), (1, 31), 2, colors.black),
    ('GRID', (1, 26), (1, 31), 1, colors.black),
    # Tablica 6 (Prijava proizvodnje)
    ('GRID', (0, 33), (1, 33), 1, colors.black),
])

table_data = [list(data.columns)]
for row in data.values:
    table_data.append(list(row))

pdf_output = PdfWriter()
pdf_template = SimpleDocTemplate("Kontrolna Lista.pdf", pagesize=letter)
pdf_elements = []

table = Table(table_data)
table.setStyle(table_style)
pdf_elements.append(table)

pdf_template.build(pdf_elements, onFirstPage=lambda canvas,
                   doc: canvas.setFont("Helvetica-Bold", 14))
