import pandas as pd
import numpy as np


sua_mensual = pd.read_excel(r'E:\01 TRABAJO\HRI\Documentos\01 IMSS\ALEJANDRO\CUINBA\02 SUA\2025\01 Enero\Pago\cedula oportuno obr-pat_gbl.xls', engine='xlrd', header=None)
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

names = sua_mensual['incapacidades'].apply(lambda x: not str(x).isdigit())
sua_mensual.loc[names, 'nombre'] = sua_mensual.loc[names, 'incapacidades']
sua_mensual.loc[names, 'incapacidades'] = np.nan

cols_to_shift = sua_mensual.columns[2:]
sua_mensual[cols_to_shift] = sua_mensual[cols_to_shift].shift(-2)

if len(sua_mensual) == 22:
    sua_mensual = sua_mensual[sua_mensual['dias'].apply(lambda x: (str(float(x)).replace('.0', '').isdigit() and float(x) <= 31) if pd.notnull(x) else False)]
    sua_mensual = sua_mensual.iloc[:-1]
else:
    sua_mensual = sua_mensual[sua_mensual['dias'].apply(lambda x: (str(float(x)).replace('.0', '').isdigit() and float(x) <= 31) if pd.notnull(x) else False)]

regex_nss = r'^\d{2}-\d{2}-\d{2}-\d{4}-\d$'
# Count how many NSS values match the regex pattern
matching_nss = sua_mensual['nss'].str.match(regex_nss).sum()

# If only one NSS matches the pattern, remove the last row
if matching_nss == 1:
    sua_mensual = sua_mensual.iloc[:-1]

print(sua_mensual.tail(10))