# Importar librerias
import pandas as pd
import numpy as np

print('Los archivos de ce las cedulas en excel se tienen que guardar en 8.0(Extended)')

# Convertir el URL de GoogleSheet a CSV
def get_sheet_csv(url):
    sheet_id = url.split('/d/')[1].split('/')[0]
    csv_url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv'
    return csv_url

# Solicitar la URL del archivo de GoogleSheet donde se encuentran las rutas de los certificados
url = input("Ingrese la URL del archivo de GoogleSheet: ")
csv_url = get_sheet_csv(url)
df = pd.read_csv(csv_url)



'''CONFRONTA MENSUAL'''

# Obtener las filas donde confronta_mensual es True
mensual_rows = df[df['confronta_mensual'] == True]

for index, row in mensual_rows.iterrows():
    try:
        # Obtener las rutas específicas para esta empresa
        cedula_path = row['cedula_mensual']
        emision_path = row['emision']
        nombre_corto = row['nombre_corto']
        mes = row['mes']
        año = row['año']

        # Procesar cédula mensual
        sua_mensual = pd.read_excel(cedula_path, engine='xlrd', header=None)
        sua_mensual.columns = ['nss', 
                               'nombre', 
                               'dias', 
                               'sdi', 
                               'licencia', 
                               'incapacidades', 
                               'ausentismos', 
                               'cuota_fija', 
                               'excedente_patronal', 
                               'excedente_obrero',  
                               'prestaciones_patronal', 
                               'prestaciones_obrero', 
                               'gastos_medicos_patronal', 
                               'gastos_medicos_obrero',
                               'riesgo_trabajo',
                               'invalidez_vida_patronal',
                               'invalidez_vida_obrero',
                               'guarderia',
                               'total_patronal',
                               'total_obrero',
                               'total'
                               ]
        # Mover los nombres de los trabajadores a la columna 'nombre'
        names = sua_mensual['incapacidades'].apply(lambda x: not str(x).isdigit())
        sua_mensual.loc[names, 'nombre'] = sua_mensual.loc[names, 'incapacidades']
        sua_mensual.loc[names, 'incapacidades'] = np.nan
        
        # Eliminar las 2 primeras filas desde la columna 'dias' hasta 'total'
        cols_to_shift = sua_mensual.columns[2:]
        sua_mensual[cols_to_shift] = sua_mensual[cols_to_shift].shift(-2)
        
        # Eliminar todas las filas que no contengan un numero y sean menores a 31 en la columna 'dias'
        if len(sua_mensual) == 22:
            sua_mensual = sua_mensual[sua_mensual['dias'].apply(lambda x: (str(float(x)).replace('.0', '').isdigit() and float(x) <= 31) if pd.notnull(x) else False)]
            sua_mensual = sua_mensual.iloc[:-1]
        else:
            sua_mensual = sua_mensual[sua_mensual['dias'].apply(lambda x: (str(float(x)).replace('.0', '').isdigit() and float(x) <= 31) if pd.notnull(x) else False)]
        
        # Definir el regex para los nss
        regex_nss = r'^\d{2}-\d{2}-\d{2}-\d{4}-\d$'
        # Defiir regex para las fechas
        regex_fecha = r'^\d{2}/\d{2}/\d{4}$'
        
        # Count how many NSS values match the regex pattern
        matching_nss = sua_mensual['nss'].str.match(regex_nss).sum()
        # If only one NSS matches the pattern, remove the last row
        if matching_nss == 1:
            sua_mensual = sua_mensual.iloc[:-1]
        
        # Asegurarse de que la columna 'nombre' sea de tipo string
        sua_mensual['nombre'] = sua_mensual['nombre'].astype(str)
        
        # Borrar cualquier valor en la columna nombre que cumpla con el regex de fecha para dejar solo los nombres
        mask = sua_mensual['nombre'].str.match(regex_fecha, na=False)
        sua_mensual.loc[mask, 'nombre'] = None
        
        # Rellenar valores que no cumplen con el regex usando el valor anterior
        mask = ~sua_mensual['nss'].str.match(regex_nss, na=False)
        sua_mensual.loc[mask, 'nss'] = np.nan
        sua_mensual['nss'] = sua_mensual['nss'].ffill()
        sua_mensual['nombre'] = sua_mensual['nombre'].ffill()
        
        # Reemplazar los '/' en la columna 'nombre' por 'Ñ'
        sua_mensual['nombre'] = sua_mensual['nombre'].str.replace('/', 'Ñ')
        
        # Agrupar por nss y sumar los valores
        sua_mensual = sua_mensual.groupby('nss').sum().reset_index()
        
        # Eliminar los '-' en la columna 'nss'
        sua_mensual['nss'] = sua_mensual['nss'].str.replace('-', '')
        
        # Ordenar por nombre y resetear el índice
        sua_mensual = sua_mensual.sort_values('nombre').reset_index(drop=True)
        
        # Convertir columnas a su tipo correspondiente
        sua_mensual[['dias', 'licencia', 'incapacidades', 'ausentismos']] = sua_mensual[['dias', 'licencia', 'incapacidades', 'ausentismos']].astype(int)
        sua_mensual['sdi'] = sua_mensual['sdi'].astype(float)
        sua_mensual.iloc[:, 7:] = sua_mensual.iloc[:, 7:].astype(float)
        
        print(f"DataFrame de cédula mensual creado exitosamente para {nombre_corto}")

        # Procesar emisión mensual
        emision_mensual = pd.read_excel(emision_path, sheet_name=1, header=4, dtype={'NSS': str})
        emision_mensual.columns = ['nss',
                                   'nombre',
                                   'origen',
                                   'tipo',
                                   'fecha',
                                   'dias',
                                   'sdi',
                                   'cuota_fija',
                                   'excedente_patronal',
                                   'excedente_obrero',
                                   'prestaciones_patronal',
                                   'prestaciones_obrero',
                                   'gastos_medicos_patronal',
                                   'gastos_medicos_obrero',
                                   'riesgo_trabajo',
                                   'invalidez_vida_patronal',
                                   'invalidez_vida_obrero',
                                   'guarderia',
                                   'total'
                                ]
        
        # Eliminar filas que en la columna 'tipo' tengan el numero 2
        emision_mensual = emision_mensual[emision_mensual['tipo'] != 2]
        
        # Eliminar columnas innecesarias
        emision_mensual.drop(columns=['origen', 'tipo', 'fecha'], inplace=True)
        
        # Reemplazar todos los '#' de la columna nombre por 'Ñ'
        emision_mensual['nombre'] = emision_mensual['nombre'].str.replace('#', 'Ñ')
        
        # Convertir columnas numéricas a float explícitamente
        numeric_columns = [
            'cuota_fija', 'excedente_patronal', 'excedente_obrero',
            'prestaciones_patronal', 'prestaciones_obrero',
            'gastos_medicos_patronal', 'gastos_medicos_obrero',
            'riesgo_trabajo', 'invalidez_vida_patronal',
            'invalidez_vida_obrero', 'guarderia', 'total'
        ]
        
        for col in numeric_columns:
            emision_mensual[col] = pd.to_numeric(emision_mensual[col], errors='coerce')
        
        # Convertir días y SDI
        emision_mensual['dias'] = pd.to_numeric(emision_mensual['dias'], errors='coerce')
        emision_mensual['sdi'] = pd.to_numeric(emision_mensual['sdi'], errors='coerce')
        
        # Convertir 'excedente_patronal' y 'excedente_obrero' a float
        emision_mensual['excedente_patronal'] = emision_mensual['excedente_patronal'].astype(float)
        emision_mensual['excedente_obrero'] = emision_mensual['excedente_obrero'].astype(float)
        
        # Eliminar espacios al final de cada valor de 'nombre'
        emision_mensual['nombre'] = emision_mensual['nombre'].str.rstrip()
        
        # Agrupar por nss y sumar los valores
        emision_mensual = emision_mensual.groupby('nss').sum().reset_index()
        
        # Ordenar por nombre y resetear el índice
        emision_mensual = emision_mensual.sort_values('nombre').reset_index(drop=True)
        
        print(f"DataFrame de emisión mensual creado exitosamente para {nombre_corto}")

        # Crear las confrontas
        sua_vs_ema = sua_mensual.copy()
        ema_vs_sua = emision_mensual.copy()
        
        # Eliminar las columnas 'total_patronal' y 'total_obrero' de sua_vs_ema
        sua_vs_ema.drop(columns=['total_patronal', 'total_obrero'], inplace=True)
        
        # Convertir 'excede_patronal' y 'excede_obrero' a float
        sua_vs_ema['excedente_patronal'] = sua_vs_ema['excedente_patronal'].astype(float)
        sua_vs_ema['excedente_obrero'] = sua_vs_ema['excedente_obrero'].astype(float)
        
        # Iterar sobre sua_vs_ema
        for index, row in sua_vs_ema.iterrows():
            nss = row['nss']
            emision_mensual_row = emision_mensual[emision_mensual['nss'] == nss]
            
            if not emision_mensual_row.empty:
                sua_vs_ema.at[index, 'dias'] = row['dias'] - emision_mensual_row.iloc[0]['dias']
                sua_vs_ema.at[index, 'sdi'] = row['sdi'] - emision_mensual_row.iloc[0]['sdi']
                sua_vs_ema.at[index, 'cuota_fija'] = row['cuota_fija'] - emision_mensual_row.iloc[0]['cuota_fija']
                sua_vs_ema.at[index, 'excedente_patronal'] = row['excedente_patronal'] - emision_mensual_row.iloc[0]['excedente_patronal']
                sua_vs_ema.at[index, 'excedente_obrero'] = row['excedente_obrero'] - emision_mensual_row.iloc[0]['excedente_obrero']
                sua_vs_ema.at[index, 'prestaciones_patronal'] = row['prestaciones_patronal'] - emision_mensual_row.iloc[0]['prestaciones_patronal']
                sua_vs_ema.at[index, 'prestaciones_obrero'] = row['prestaciones_obrero'] - emision_mensual_row.iloc[0]['prestaciones_obrero']
                sua_vs_ema.at[index, 'gastos_medicos_patronal'] = row['gastos_medicos_patronal'] - emision_mensual_row.iloc[0]['gastos_medicos_patronal']
                sua_vs_ema.at[index, 'gastos_medicos_obrero'] = row['gastos_medicos_obrero'] - emision_mensual_row.iloc[0]['gastos_medicos_obrero']
                sua_vs_ema.at[index, 'riesgo_trabajo'] = row['riesgo_trabajo'] - emision_mensual_row.iloc[0]['riesgo_trabajo']
                sua_vs_ema.at[index, 'invalidez_vida_patronal'] = row['invalidez_vida_patronal'] - emision_mensual_row.iloc[0]['invalidez_vida_patronal']
                sua_vs_ema.at[index, 'invalidez_vida_obrero'] = row['invalidez_vida_obrero'] - emision_mensual_row.iloc[0]['invalidez_vida_obrero']
                sua_vs_ema.at[index, 'guarderia'] = row['guarderia'] - emision_mensual_row.iloc[0]['guarderia']
                sua_vs_ema.at[index, 'total'] = row['total'] - emision_mensual_row.iloc[0]['total']
            else:
                sua_vs_ema.at[index, 'dias'] = np.nan
                sua_vs_ema.at[index, 'sdi'] = np.nan
                sua_vs_ema.at[index, 'cuota_fija'] = np.nan
                sua_vs_ema.at[index, 'excedente_patronal'] = np.nan
                sua_vs_ema.at[index, 'excedente_obrero'] = np.nan
                sua_vs_ema.at[index, 'prestaciones_patronal'] = np.nan
                sua_vs_ema.at[index, 'prestaciones_obrero'] = np.nan
                sua_vs_ema.at[index, 'gastos_medicos_patronal'] = np.nan
                sua_vs_ema.at[index, 'gastos_medicos_obrero'] = np.nan
                sua_vs_ema.at[index, 'riesgo_trabajo'] = np.nan
                sua_vs_ema.at[index, 'invalidez_vida_patronal'] = np.nan
                sua_vs_ema.at[index, 'invalidez_vida_obrero'] = np.nan
                sua_vs_ema.at[index, 'guarderia'] = np.nan
                sua_vs_ema.at[index, 'total'] = np.nan
                
        sua_vs_ema['observacion'] = ''
        
        # Añadir observaciones basadas en diferentes condiciones
        for index, row in sua_vs_ema.iterrows():
            observations = []
            nss = row['nss']
            emision_row = emision_mensual[emision_mensual['nss'] == nss]
            
            if pd.isna(row['total']):
                observations.append('TRABAJADOR NO SE LOCALIZO EN EMISION')
            else:
                if row['nombre'] != emision_row.iloc[0]['nombre']:
                    observations.append('NOMBRE DIFERENTE')
                if row['dias'] != 0:
                    observations.append('DIFERENCIA EN DIAS')
                if row['total'] != 0:
                    if row['incapacidades'] > 0:
                        diff = (emision_row.iloc[0]['total'] / emision_row.iloc[0]['dias']) * row['incapacidades']
                        observations.append(f'DIFERENCIA POR INCAPACIDAD: {diff}')
                    else:
                        observations.append('DIFERENCIA EN CUOTAS')
                
            if observations:
                sua_vs_ema.at[index, 'observacion'] = ', '.join(observations)
            else:
                sua_vs_ema.at[index, 'observacion'] = 'SIN DIFERENCIAS'
                        
        # Iterar sobre ema_vs_sua
        for index, row in ema_vs_sua.iterrows():
            nss = row['nss']
            sua_mensual_row = sua_mensual[sua_mensual['nss'] == nss]
            
            if not sua_mensual_row.empty:
                ema_vs_sua.at[index, 'dias'] = row['dias'] - sua_mensual_row.iloc[0]['dias']
                ema_vs_sua.at[index, 'sdi'] = row['sdi'] - sua_mensual_row.iloc[0]['sdi']
                ema_vs_sua.at[index, 'cuota_fija'] = row['cuota_fija'] - sua_mensual_row.iloc[0]['cuota_fija']
                ema_vs_sua.at[index, 'excedente_patronal'] = row['excedente_patronal'] - sua_mensual_row.iloc[0]['excedente_patronal']
                ema_vs_sua.at[index, 'excedente_obrero'] = row['excedente_obrero'] - sua_mensual_row.iloc[0]['excedente_obrero']
                ema_vs_sua.at[index, 'prestaciones_patronal'] = row['prestaciones_patronal'] - sua_mensual_row.iloc[0]['prestaciones_patronal']
                ema_vs_sua.at[index, 'prestaciones_obrero'] = row['prestaciones_obrero'] - sua_mensual_row.iloc[0]['prestaciones_obrero']
                ema_vs_sua.at[index, 'gastos_medicos_patronal'] = row['gastos_medicos_patronal'] - sua_mensual_row.iloc[0]['gastos_medicos_patronal']
                ema_vs_sua.at[index, 'gastos_medicos_obrero'] = row['gastos_medicos_obrero'] - sua_mensual_row.iloc[0]['gastos_medicos_obrero']
                ema_vs_sua.at[index, 'riesgo_trabajo'] = row['riesgo_trabajo'] - sua_mensual_row.iloc[0]['riesgo_trabajo']
                ema_vs_sua.at[index, 'invalidez_vida_patronal'] = row['invalidez_vida_patronal'] - sua_mensual_row.iloc[0]['invalidez_vida_patronal']
                ema_vs_sua.at[index, 'invalidez_vida_obrero'] = row['invalidez_vida_obrero'] - sua_mensual_row.iloc[0]['invalidez_vida_obrero']
                ema_vs_sua.at[index, 'guarderia'] = row['guarderia'] - sua_mensual_row.iloc[0]['guarderia']
                ema_vs_sua.at[index, 'total'] = row['total'] - sua_mensual_row.iloc[0]['total']
            else:
                ema_vs_sua.at[index, 'dias'] = np.nan
                ema_vs_sua.at[index, 'sdi'] = np.nan
                ema_vs_sua.at[index, 'cuota_fija'] = np.nan
                ema_vs_sua.at[index, 'excedente_patronal'] = np.nan
                ema_vs_sua.at[index, 'excedente_obrero'] = np.nan
                ema_vs_sua.at[index, 'prestaciones_patronal'] = np.nan
                ema_vs_sua.at[index, 'prestaciones_obrero'] = np.nan
                ema_vs_sua.at[index, 'gastos_medicos_patronal'] = np.nan
                ema_vs_sua.at[index, 'gastos_medicos_obrero'] = np.nan
                ema_vs_sua.at[index, 'riesgo_trabajo'] = np.nan
                ema_vs_sua.at[index, 'invalidez_vida_patronal'] = np.nan
                ema_vs_sua.at[index, 'invalidez_vida_obrero'] = np.nan
                ema_vs_sua.at[index, 'guarderia'] = np.nan
                ema_vs_sua.at[index, 'total'] = np.nan
                
        ema_vs_sua['observacion'] = ''
        
        # Add observations based on different conditions
        for index, row in ema_vs_sua.iterrows():
            observations = []
            nss = row['nss']
            sua_row = sua_mensual[sua_mensual['nss'] == nss]
            
            if pd.isna(row['total']):
                observations.append('TRABAJADOR NO SE LOCALIZO EN CEDULA')
            else:
                if row['nombre'] != sua_row.iloc[0]['nombre']:
                    observations.append('NOMBRE DIFERENTE')
                if row['dias'] != 0:
                    observations.append('DIFERENCIA EN DIAS')
                if row['total'] != 0:
                    observations.append('DIFERENCIA EN CUOTAS')
                
            if observations:
                ema_vs_sua.at[index, 'observacion'] = ', '.join(observations)
            else:
                ema_vs_sua.at[index, 'observacion'] = 'SIN DIFERENCIAS'
            
        print(f"DataFrame de confronta mensual creado exitosamente para {nombre_corto}")

        # Guardar los resultados para esta empresa específica
        filename = f'01_MENSUAL_{nombre_corto}_{mes}_{año}.xlsx'
        with pd.ExcelWriter(filename) as writer:
            sua_mensual.to_excel(writer, sheet_name='sua_mensual', index=False, float_format='%.2f')
            emision_mensual.to_excel(writer, sheet_name='emision_mensual', index=False, float_format='%.2f')  
            sua_vs_ema.to_excel(writer, sheet_name='sua_vs_ema', index=False, float_format='%.2f')
            ema_vs_sua.to_excel(writer, sheet_name='ema_vs_sua', index=False, float_format='%.2f')
        print(f"DataFrames guardados exitosamente en {filename}")

    except Exception as e:
        print(f"Error procesando {nombre_corto}: {e}")



'''CONFRONTA BIMESTRAL'''

if df['confronta_bimestral'].any():
    # Obtener las filas donde confronta_bimestral es True
    bimestral_paths = df[df['confronta_bimestral'] == True]['cedula_bimestral']
    for path in bimestral_paths:
        try:
            
            '''CEDULA BIMESTRAL'''
            
            sua_bimestral = pd.read_excel(path, engine='xlrd', header=None)
            sua_bimestral.columns = ['nss',
                                     'nombre',
                                     'dias',
                                     'sdi',
                                     'licencia',
                                     'incapacidades',
                                     'ausentismos',
                                     'retiro',
                                     'cv_patronal',
                                     'cv_obrero',
                                     'rcv_suma',
                                     'aportacion_patronal',
                                     'borrar_columna',
                                     'factor',
                                     'borrar_columna',
                                     'amortizacion',
                                     'infonavit_suma',
                                     'numero_credito',
                                     'borrar_columna',
                                     ]
            
            # Mover los nombres de los trabajadores a la columna 'nombre'
            names = sua_bimestral['incapacidades'].apply(lambda x: not str(x).isdigit())
            sua_bimestral.loc[names, 'nombre'] = sua_bimestral.loc[names, 'incapacidades']
            sua_bimestral.loc[names, 'incapacidades'] = np.nan
            
            # Borrar columnas innecesarias
            sua_bimestral.drop(columns=['borrar_columna'], inplace=True)
            sua_bimestral.drop(columns=['numero_credito'], inplace=True)
            
            # Eliminar las 2 primeras filas desde la columna 'dias' hasta 'numero_credito'
            cols_to_shift = sua_bimestral.columns[2:]
            sua_bimestral[cols_to_shift] = sua_bimestral[cols_to_shift].shift(-2)
            
            # Eliminar todas las filas que no contengan un int en la columna 'dias' o sean mayores a 62
            sua_bimestral = sua_bimestral[sua_bimestral['dias'].apply(lambda x: str(x).isdigit() and int(x) <= 62)]
            
            # Definir el regex para los nss
            regex_nss = r'^\d{2}-\d{2}-\d{2}-\d{4}-\d$'
            
            # Rellenar valores que no cumplen con el regex usando el valor anterior
            mask = ~sua_bimestral['nss'].str.match(regex_nss, na=False)
            sua_bimestral.loc[mask, 'nss'] = np.nan
            sua_bimestral['nss'] = sua_bimestral['nss'].ffill()
            sua_bimestral['nombre'] = sua_bimestral['nombre'].ffill()
            
            # Reemplazar los '/' en la columna 'nombre' por 'Ñ'
            sua_bimestral['nombre'] = sua_bimestral['nombre'].str.replace('/', 'Ñ')
            
            # Rellenar lo valores ausentes con 0
            sua_bimestral.fillna(0, inplace=True)
            
            # Convertir columnas a su tipo correspondiente
            sua_bimestral[['dias', 'licencia', 'incapacidades', 'ausentismos']] = sua_bimestral[['dias', 'licencia', 'incapacidades', 'ausentismos']].astype(int)
            sua_bimestral['sdi'] = sua_bimestral['sdi'].astype(float)
            
            # Agrupar por nss y utilizar agg por que hay columnas con diferentes funciones
            agg_dict = {
                'nombre': 'first',
                'dias': 'sum',
                'sdi': 'sum',
                'licencia': 'sum',
                'incapacidades': 'sum',
                'ausentismos': 'sum',
                'retiro': 'sum',
                'cv_patronal': 'sum',
                'cv_obrero': 'sum',
                'rcv_suma': 'sum',
                'aportacion_patronal': 'sum',
                'factor': 'first',
                'amortizacion': 'sum',
                'infonavit_suma': 'sum',
            }
            sua_bimestral = sua_bimestral.groupby('nss').agg(agg_dict).reset_index()
            
            # Crear columna total para sumar las columnas rcv_suma y infonavit_suma
            sua_bimestral['total'] = sua_bimestral['rcv_suma'] + sua_bimestral['infonavit_suma']
            
            # Ordenar por nombre y resetear el índice
            sua_bimestral = sua_bimestral.sort_values('nombre').reset_index(drop=True)
            
            print("DataFrame bimestral creado exitosamente")
        except Exception as e:
            print(f"Error al leer la cedula bimestral: {e}")
            
if df['confronta_bimestral'].any():
    # Obtener las filas donde emision_bimestral es True
    emision_paths = df[df['confronta_bimestral'] == True]['emision']
    for path in emision_paths:
        try:
            
            '''EMISION BIMESTRAL'''
            
            emision_bimestral = pd.read_excel(path, sheet_name=2, header=4, dtype={'NSS': str})
            emision_bimestral.columns = ['nss',
                                       'nombre',
                                       'origen',
                                       'tipo',
                                       'fecha',
                                       'dias',
                                       'sdi',
                                       'retiro',
                                       'cv_patronal',
                                       'cv_obrero',
                                       'rcv_suma',
                                       'aportacion_patronal',
                                       'tipo_factor',
                                       'factor',
                                       'numero_credito',
                                       'amortizacion',
                                       'infonavit_suma',
                                       'total'
                                       ]
            
            # Eliminar filas que en la columna 'tipo' tengan el numero 2
            emision_bimestral = emision_bimestral[emision_bimestral['tipo'] != 2]
            
            # Eliminar columnas innecesarias
            emision_bimestral.drop(columns=['origen', 'tipo', 'fecha'], inplace=True)
            
            # Reemplazar todos los '#' de la columna nombre por 'Ñ'
            emision_bimestral['nombre'] = emision_bimestral['nombre'].str.replace('#', 'Ñ')
            
            # Convertir columna factor a float, reemplazando guiones por ceros y eliminando comas
            emision_bimestral['factor'] = emision_bimestral['factor'].replace('-', '0')
            emision_bimestral['factor'] = emision_bimestral['factor'].str.replace(',', '')
            emision_bimestral['factor'] = emision_bimestral['factor'].astype(float)
            
            # Agrupar por nss y sumar los valores
            emision_bimestral = emision_bimestral.groupby('nss').sum().reset_index()
            
            # Ordenar por nombre y resetear el índice
            emision_bimestral = emision_bimestral.sort_values('nombre').reset_index(drop=True)
            
            print("DataFrames de emisiones bimestrales creados exitosamente")
            
            # Crear las confrontas
            sua_vs_eba = sua_bimestral.copy()
            eba_vs_sua = emision_bimestral.copy()
            
            
            
            
            
            
        except Exception as e:
            print(f"Error al leer la emision bimestral: {e}")