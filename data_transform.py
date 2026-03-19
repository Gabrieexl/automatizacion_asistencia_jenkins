import os
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_DIR, 'temp')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
INPUT_RAW = os.path.join(TEMP_DIR, 'asistencia_raw.csv')
DETAIL_FILE = os.path.join(OUTPUT_DIR, 'asistencia_detalle_procesada.csv')
SUMMARY_FILE = os.path.join(OUTPUT_DIR, 'resumen_asistencia.xlsx')

os.makedirs(OUTPUT_DIR, exist_ok=True)

print('=== ETAPA TRANSFORM ===')

if not os.path.exists(INPUT_RAW):
    raise FileNotFoundError('No existe asistencia_raw.csv. Ejecuta primero data_read.py')

df = pd.read_csv(INPUT_RAW, encoding='utf-8-sig')

df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
df['Minutos_Tardanza'] = pd.to_numeric(df['Minutos_Tardanza'], errors='coerce').fillna(0)
df['Estado'] = df['Estado'].astype(str).str.strip().str.upper()

if 'Área' in df.columns:
    area_col = 'Área'
elif 'Area' in df.columns:
    area_col = 'Area'
else:
    raise ValueError("No se encontró la columna 'Área' o 'Area'.")

if 'Observación' in df.columns:
    obs_col = 'Observación'
elif 'Observacion' in df.columns:
    obs_col = 'Observacion'
else:
    obs_col = None

before = len(df)
df = df.dropna(subset=['Fecha', 'Nombre'])
after = len(df)
print(f'Registros válidos después de limpieza: {after} (eliminados: {before - after})')

mapping = {
    'ASISTIÓ': 'ASISTIO',
    'ASISTIO': 'ASISTIO',
    'TARDANZA': 'TARDANZA',
    'FALTA': 'FALTA',
    'PERMISO': 'PERMISO',
    'VACACIONES': 'VACACIONES'
}
df['Estado_Normalizado'] = df['Estado'].map(mapping).fillna(df['Estado'])

resumen_general = pd.DataFrame([
    {'Indicador': 'Total de registros', 'Valor': len(df)},
    {'Indicador': 'Asistencias normales', 'Valor': int((df['Estado_Normalizado'] == 'ASISTIO').sum())},
    {'Indicador': 'Tardanzas', 'Valor': int((df['Estado_Normalizado'] == 'TARDANZA').sum())},
    {'Indicador': 'Faltas', 'Valor': int((df['Estado_Normalizado'] == 'FALTA').sum())},
    {'Indicador': 'Permisos', 'Valor': int((df['Estado_Normalizado'] == 'PERMISO').sum())},
    {'Indicador': 'Vacaciones', 'Valor': int((df['Estado_Normalizado'] == 'VACACIONES').sum())},
])

incidencias = df[df['Estado_Normalizado'].isin(['TARDANZA', 'FALTA'])].copy()

resumen_empleado = (
    incidencias.groupby(['ID_Empleado', 'Nombre', area_col], dropna=False)
    .agg(
        Total_Incidencias=('Estado_Normalizado', 'count'),
        Total_Tardanzas=('Estado_Normalizado', lambda x: (x == 'TARDANZA').sum()),
        Total_Faltas=('Estado_Normalizado', lambda x: (x == 'FALTA').sum()),
        Minutos_Tardanza_Acumulados=('Minutos_Tardanza', 'sum')
    )
    .reset_index()
    .sort_values(['Total_Incidencias', 'Total_Faltas', 'Minutos_Tardanza_Acumulados'], ascending=False)
)

resumen_area = (
    df.groupby(area_col, dropna=False)
    .agg(
        Registros=('Estado_Normalizado', 'count'),
        Tardanzas=('Estado_Normalizado', lambda x: (x == 'TARDANZA').sum()),
        Faltas=('Estado_Normalizado', lambda x: (x == 'FALTA').sum()),
        Permisos=('Estado_Normalizado', lambda x: (x == 'PERMISO').sum()),
        Vacaciones=('Estado_Normalizado', lambda x: (x == 'VACACIONES').sum())
    )
    .reset_index()
    .sort_values('Registros', ascending=False)
)

resumen_dia = (
    df.groupby(df['Fecha'].dt.date)
    .agg(
        Registros=('Estado_Normalizado', 'count'),
        Tardanzas=('Estado_Normalizado', lambda x: (x == 'TARDANZA').sum()),
        Faltas=('Estado_Normalizado', lambda x: (x == 'FALTA').sum())
    )
    .reset_index()
)

top_incidencias = resumen_empleado.head(10).copy()

if obs_col is None:
    df['Observación_Final'] = ''
    obs_final = 'Observación_Final'
else:
    obs_final = obs_col

cols_detalle = [
    'Fecha',
    'ID_Empleado',
    'Nombre',
    area_col,
    'Hora_Entrada_Programada',
    'Hora_Entrada_Real',
    'Hora_Salida_Programada',
    'Hora_Salida_Real',
    'Estado',
    'Estado_Normalizado',
    'Minutos_Tardanza',
    'Incidencia',
    obs_final
]

for col in cols_detalle:
    if col not in df.columns:
        df[col] = ''

df[cols_detalle].to_csv(DETAIL_FILE, index=False, encoding='utf-8-sig')
print(f'Detalle procesado generado: {DETAIL_FILE}')

with pd.ExcelWriter(SUMMARY_FILE, engine='openpyxl') as writer:
    resumen_general.to_excel(writer, index=False, sheet_name='Resumen_General')
    resumen_empleado.to_excel(writer, index=False, sheet_name='Incidencias_Empleado')
    resumen_area.to_excel(writer, index=False, sheet_name='Resumen_Area')
    resumen_dia.to_excel(writer, index=False, sheet_name='Resumen_Dia')
    top_incidencias.to_excel(writer, index=False, sheet_name='Top_Incidencias')

print(f'Resumen Excel generado: {SUMMARY_FILE}')
print('TRANSFORM finalizado correctamente.')