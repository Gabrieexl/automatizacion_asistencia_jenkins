import os
import pandas as pd
import unicodedata

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
TEMP_DIR = os.path.join(BASE_DIR, 'temp')
INPUT_FILE = os.path.join(DATA_DIR, 'asistencia_control_jenkins.xlsx')
SHEET_NAME = 'Asistencia_Febrero_2026'
OUTPUT_RAW = os.path.join(TEMP_DIR, 'asistencia_raw.csv')

os.makedirs(TEMP_DIR, exist_ok=True)

print('=== ETAPA READ ===')
print(f'Buscando archivo: {INPUT_FILE}')


# FUNCION PARA ELIMINAR TILDES
def quitar_tildes(texto):
    texto = str(texto)
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join([c for c in texto if not unicodedata.combining(c)])
    return texto


if not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(
        'No se encontró el archivo Excel de asistencia. '
        'Ubícalo en la carpeta data con el nombre asistencia_control_jenkins.xlsx'
    )

try:
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
except ValueError:
    xls = pd.ExcelFile(INPUT_FILE)
    raise ValueError(
        f"No se encontró la hoja '{SHEET_NAME}'. Hojas disponibles: {xls.sheet_names}"
    )

print(f'Registros leídos: {len(df)}')
print(f'Columnas originales: {list(df.columns)}')


# NORMALIZAR COLUMNAS (SIN TILDES)
columnas_normalizadas = {}
for col in df.columns:
    nueva = quitar_tildes(col).strip()
    columnas_normalizadas[col] = nueva

df = df.rename(columns=columnas_normalizadas)

print(f'Columnas normalizadas: {list(df.columns)}')


# COLUMNAS REQUERIDAS YA NORMALIZADAS
required_columns = [
    'Fecha',
    'ID_Empleado',
    'Nombre',
    'Area',
    'Hora_Entrada_Real',
    'Hora_Salida_Real',
    'Estado',
    'Minutos_Tardanza'
]

missing = [col for col in required_columns if col not in df.columns]

if missing:
    raise ValueError(f'Faltan columnas obligatorias en el Excel: {missing}')


df.to_csv(OUTPUT_RAW, index=False, encoding='utf-8-sig')

print(f'Archivo temporal generado: {OUTPUT_RAW}')
print('READ finalizado correctamente.')