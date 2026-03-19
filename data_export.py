import os
import textwrap
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
SUMMARY_FILE = os.path.join(OUTPUT_DIR, 'resumen_asistencia.xlsx')
PDF_FILE = os.path.join(OUTPUT_DIR, 'reporte_asistencia.pdf')

print('=== ETAPA EXPORT ===')

if not os.path.exists(SUMMARY_FILE):
    raise FileNotFoundError('No existe resumen_asistencia.xlsx. Ejecuta primero data_transform.py')

resumen_general = pd.read_excel(SUMMARY_FILE, sheet_name='Resumen_General')
resumen_empleado = pd.read_excel(SUMMARY_FILE, sheet_name='Incidencias_Empleado')
resumen_area = pd.read_excel(SUMMARY_FILE, sheet_name='Resumen_Area')
resumen_dia = pd.read_excel(SUMMARY_FILE, sheet_name='Resumen_Dia')

def small_table_from_df(df, title, max_rows=8):
    styles = getSampleStyleSheet()
    elements = [Paragraph(f'<b>{title}</b>', styles['Heading3']), Spacer(1, 0.2 * cm)]

    show = df.head(max_rows).copy()
    for col in show.columns:
        show[col] = show[col].astype(str).apply(
            lambda x: '\n'.join(textwrap.wrap(x, width=20)) if len(x) > 20 else x
        )

    data = [list(show.columns)] + show.values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 0.4 * cm))
    return elements

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='JustifyCustom', alignment=4, leading=14))

doc = SimpleDocTemplate(
    PDF_FILE,
    pagesize=A4,
    rightMargin=1.8 * cm,
    leftMargin=1.8 * cm,
    topMargin=1.5 * cm,
    bottomMargin=1.5 * cm
)

story = []
story.append(Paragraph('Reporte Automático de Control de Asistencia', styles['Title']))
story.append(Spacer(1, 0.3 * cm))
story.append(Paragraph(
    'Documento generado automáticamente por Jenkins y Python. Resume incidencias de asistencia, tardanzas, faltas y distribución por área.',
    styles['BodyText']
))
story.append(Spacer(1, 0.5 * cm))

data_general = [list(resumen_general.columns)] + resumen_general.values.tolist()
table_general = Table(data_general, colWidths=[9 * cm, 4 * cm])
table_general.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.beige]),
]))
story.append(Paragraph('<b>Resumen General</b>', styles['Heading2']))
story.append(table_general)
story.append(Spacer(1, 0.5 * cm))

total_tardanzas = int(resumen_general.loc[resumen_general['Indicador'] == 'Tardanzas', 'Valor'].iloc[0])
total_faltas = int(resumen_general.loc[resumen_general['Indicador'] == 'Faltas', 'Valor'].iloc[0])

comentario = (
    f'Se identificaron <b>{total_tardanzas}</b> tardanzas y <b>{total_faltas}</b> faltas en el periodo analizado. '
    'El reporte permite reconocer al personal con mayores incidencias y las áreas que requieren seguimiento.'
)
story.append(Paragraph(comentario, styles['JustifyCustom']))
story.append(Spacer(1, 0.4 * cm))

for block in small_table_from_df(resumen_empleado, 'Top de empleados con más incidencias', max_rows=10):
    story.append(block)
for block in small_table_from_df(resumen_area, 'Resumen por área', max_rows=10):
    story.append(block)
for block in small_table_from_df(resumen_dia, 'Resumen por día', max_rows=10):
    story.append(block)

doc.build(story)
print(f'PDF generado correctamente: {PDF_FILE}')
print('EXPORT finalizado correctamente.')