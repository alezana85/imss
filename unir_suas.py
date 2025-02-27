import pyodbc
import os
from pathlib import Path

def merge_mdb_files(source_folder, output_file, password):
    """
    Combina múltiples archivos .mdb protegidos con contraseña en uno solo.
    
    Args:
        source_folder (str): Ruta a la carpeta que contiene los archivos .mdb
        output_file (str): Ruta donde se guardará el archivo combinado
        password (str): Contraseña para acceder a los archivos .mdb
    """
    # Lista para almacenar los nombres de las tablas
    table_names = []
    
    # Obtener el primer archivo .mdb para extraer la estructura
    first_file = next(Path(source_folder).glob('*.mdb'))
    
    # Conectar al primer archivo para obtener los nombres de las tablas
    # Incluir la contraseña en la cadena de conexión
    conn_str = f'Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={first_file};PWD={password};'
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    
    # Obtener nombres de todas las tablas
    for row in cursor.tables():
        if row.table_type == 'TABLE':
            table_names.append(row.table_name)
    
    # Crear nuevo archivo de Access
    import shutil
    shutil.copy(first_file, output_file)
    
    # Conectar al archivo de salida con contraseña
    output_conn_str = f'Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={output_file};PWD={password};'
    output_conn = pyodbc.connect(output_conn_str)
    output_cursor = output_conn.cursor()
    
    # Procesar cada archivo .mdb
    for mdb_file in Path(source_folder).glob('*.mdb'):
        if str(mdb_file) == output_file:
            continue
            
        print(f'Procesando: {mdb_file}')
        
        # Conectar al archivo fuente con contraseña
        source_conn_str = f'Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_file};PWD={password};'
        source_conn = pyodbc.connect(source_conn_str)
        source_cursor = source_conn.cursor()
        
        # Procesar cada tabla
        for table_name in table_names:
            try:
                # Leer datos de la tabla fuente
                source_cursor.execute(f'SELECT * FROM [{table_name}]')
                rows = source_cursor.fetchall()
                
                # Insertar datos en la tabla de destino
                for row in rows:
                    try:
                        columns = ', '.join(['?' for _ in range(len(row))])
                        insert_sql = f'INSERT INTO [{table_name}] VALUES ({columns})'
                        output_cursor.execute(insert_sql, row)
                    except pyodbc.IntegrityError:
                        # Si hay un error de duplicado, continuar con el siguiente registro
                        continue
                
                output_conn.commit()
                
            except Exception as e:
                print(f'Error procesando tabla {table_name} en archivo {mdb_file}: {str(e)}')
                continue
        
        source_conn.close()
    
    output_conn.close()
    print('Proceso completado.')

# Ejemplo de uso
if __name__ == '__main__':
    source_folder = r'E:\TRABAJOS IMSS\HRI\Documentos\01 IMSS\01 ACCES\03 SEPARACIONES\2025\01 Enero\LUIS'
    # Modificar la ruta de salida para incluir el nombre del archivo
    output_file = os.path.join(source_folder, 'SUA.MDB')
    password = 'S5@N52V49'  # Contraseña para los archivos
    merge_mdb_files(source_folder, output_file, password)