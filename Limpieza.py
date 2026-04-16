import pandas as pd

# Load the Excel file to inspect its contents
file_path = 'Libro1.xlsx'
excel_data = pd.ExcelFile(file_path)

# Display sheet names to understand the structure of the file
sheet_names = excel_data.sheet_names
print(sheet_names)

# Load each sheet to inspect its contents
contrato_df = pd.read_excel(file_path, sheet_name='CONTRATO')
servicios_df = pd.read_excel(file_path, sheet_name='SERVICIOS')
estados_df = pd.read_excel(file_path, sheet_name='ESTADOS')

# Display the first few rows of each dataframe to understand their structure and content
print(contrato_df.head())
print(servicios_df.head())
print(estados_df.head())

# Renombrar columnas para evitar espacios y caracteres especiales en 'CONTRATO'
contrato_df.columns = contrato_df.columns.str.replace(' ', '_').str.replace('\n', '_').str.upper()

# Convertir columnas de fechas a datetime
contrato_df['FECHA_CONTRATO'] = pd.to_datetime(contrato_df['FECHA_CONTRATO'], errors='coerce')
contrato_df['FECHA_INSTALACION'] = pd.to_datetime(contrato_df['FECHA_INSTALACION'], errors='coerce')

# Llenar valores nulos de 'CÓDIGO_IPLUS' con 'DESCONOCIDO'
contrato_df['CÓDIGO_IPLUS'].fillna('DESCONOCIDO', inplace=True)

# Eliminar filas donde las fechas son nulas
contrato_df.dropna(subset=['FECHA_CONTRATO', 'FECHA_INSTALACION'], inplace=True)

# Estandarizar los valores en las columnas categóricas
contrato_df['SUCURSAL'] = contrato_df['SUCURSAL'].str.strip().str.upper()
contrato_df['ACTUAL_PLUS'] = contrato_df['ACTUAL_PLUS'].str.strip().str.upper()
contrato_df['COD_ESTADO_TT'] = contrato_df['COD_ESTADO_TT'].str.strip().str.upper()

# Mostrar los cambios realizados
print(contrato_df.head())

# Renombrar columnas para evitar espacios y caracteres especiales en 'SERVICIOS'
servicios_df.columns = servicios_df.columns.str.replace(' ', '_').str.replace('\n', '_').str.upper()

# Estandarizar los valores en las columnas categóricas
servicios_df['SERVICIO'] = servicios_df['SERVICIO'].str.strip().str.upper()
servicios_df['ESTADO'] = servicios_df['ESTADO'].str.strip().str.upper()

# Mostrar los cambios realizados
print(servicios_df.head())

# Obtener listas únicas de estados de ambas columnas de la hoja 'ESTADOS'
estados_iplus_unicos = estados_df['ESTADO IPLUS'].str.upper().unique()
estados_tt_unicos = estados_df['ESTADO TT'].str.upper().unique()

# Verificar si todos los estados en 'SERVICIOS' están en alguna de las listas de estados únicos
servicios_estados_unicos = servicios_df['ESTADO'].unique()
contrato_estados_unicos = contrato_df['ACTUAL_PLUS'].unique()

# Mostrar estados únicos de 'SERVICIOS' y 'CONTRATO' que no están en ninguna de las listas
servicios_estados_invalidos = [estado for estado in servicios_estados_unicos if estado not in estados_iplus_unicos and estado not in estados_tt_unicos]
contrato_estados_invalidos = [estado for estado in contrato_estados_unicos if estado not in estados_iplus_unicos and estado not in estados_tt_unicos]

print(servicios_estados_invalidos, contrato_estados_invalidos)

# Crear un diccionario de mapeo para alinear los estados inválidos
mapeo_estados = {
    'SUSPENDIDOS': 'SUSPENDIDO',
    'ACTIVOS': 'ACTIVO',
    'CUENTAS INCOBRABLES': 'INACTIVO'
}

# Aplicar el mapeo a las hojas 'SERVICIOS' y 'CONTRATO'
servicios_df['ESTADO'] = servicios_df['ESTADO'].map(mapeo_estados).fillna(servicios_df['ESTADO'])
contrato_df['ACTUAL_PLUS'] = contrato_df['ACTUAL_PLUS'].map(mapeo_estados).fillna(contrato_df['ACTUAL_PLUS'])

# Verificar nuevamente los estados únicos después del mapeo
servicios_estados_unicos = servicios_df['ESTADO'].unique()
contrato_estados_unicos = contrato_df['ACTUAL_PLUS'].unique()

print(servicios_estados_unicos, contrato_estados_unicos)

# Guardar los dataframes limpiados de nuevo en un archivo Excel
output_path = 'DatosPlusNormalizados.xlsx'

with pd.ExcelWriter(output_path) as writer:
    contrato_df.to_excel(writer, sheet_name='CONTRATO', index=False)
    servicios_df.to_excel(writer, sheet_name='SERVICIOS', index=False)
    estados_df.to_excel(writer, sheet_name='ESTADOS', index=False)

print("Archivo guardado en:", output_path)
