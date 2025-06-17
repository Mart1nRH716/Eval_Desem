import os
import pandas as pd
import pyodbc

# -------------------------------
# CONFIGURACIÓN
# -------------------------------
CARPETA_REPORTES = "Reportes/OOAD"  
HOJA_EXCEL = "Base Publicación" 

CONEXION_SQL = (
    "DRIVER={SQL Server};"
    "SERVER=localhost;"          
    "DATABASE=DBTRABAJODAS;"  # Las pruebas se realizaron en local para despues migrar las tabalas al 103    
    "UID=;" # ------------------Poner usuario  y contraseña              
    "PWD=;"          
)
TABLA_DESTINO = "ResultadosEvaluacionOOAD"  

def procesar_excel(ruta_archivo):
    #Leemos el archivo de Excel y lo almacenamos en nun dataFrame
    try:
        
        df = pd.read_excel(ruta_archivo, sheet_name=HOJA_EXCEL)
        print(f"\nProcesando archivo: {os.path.basename(ruta_archivo)} - Filas: {len(df)}")
        
        columnas_requeridas = ['Año', 'mes_i', 'cve_del', 'Delegación', 'tipo', 
                              'Proceso_Normativa', 'Ponderación', 'Calificación', 'Logro']
        
        for col in columnas_requeridas:
            if col not in df.columns:
                raise ValueError(f"La columna '{col}' no se encuentra en el archivo Excel")
        
        # Agrupar por Delegación y tipo, 
        df_agrupado = df.groupby(['Delegación', 'tipo'])['Calificación'].mean().reset_index()
        df_agrupado = df_agrupado.sort_values(['tipo', 'Calificación'], ascending=[True, False])
        
        # Asignar ranking dentro de cada grupo de tipo
        df_agrupado['Lugar_que_ocupa'] = df_agrupado.groupby('tipo')['Calificación'].rank(ascending=False, method='min')
        
        # Unir el ranking al dataframe original
        df_final = pd.merge(df, df_agrupado[['Delegación', 'tipo', 'Lugar_que_ocupa']], 
                           on=['Delegación', 'tipo'], how='left')
        
        df_final = df_final.sort_values(['tipo', 'Lugar_que_ocupa'])
        
        return df_final
        
    except Exception as e:
        print(f"Error al procesar el archivo {ruta_archivo}: {e}")
        return None

def insertar_dataframe_sql(df, conexion_str, tabla_destino):
    #Insertar el dataFrame obtenido
    if df is None or df.empty:
        print("DataFrame vacío, no hay datos para insertar")
        return
    
    try:
        conn = pyodbc.connect(conexion_str)
        cursor = conn.cursor()
        
        # Creamo la tabla si no existe
        cursor.execute(f"""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='{tabla_destino}' AND xtype='U')
        CREATE TABLE {tabla_destino} (
            Año INT,
            mes_i INT,
            cve_del NVARCHAR(10),
            Delegación NVARCHAR(100),
            tipo NVARCHAR(50),
            Proceso_Normativa NVARCHAR(100),
            Ponderación FLOAT,
            Calificación FLOAT,
            Logro FLOAT,
            Lugar_que_ocupa INT
        )
        """)
        
        # Insertar datos con cursor
        total_insertados = 0
        for _, row in df.iterrows():
            try:
                cursor.execute(f"""
                INSERT INTO {tabla_destino} (
                    Año, mes_i, cve_del, Delegación, tipo, 
                    Proceso_Normativa, Ponderación, Calificación, Logro, Lugar_que_ocupa
                ) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, 
                row['Año'], row['mes_i'], row['cve_del'], row['Delegación'], row['tipo'],
                row['Proceso_Normativa'], row['Ponderación'], row['Calificación'], 
                row['Logro'], row['Lugar_que_ocupa'])
                total_insertados += 1
            except pyodbc.IntegrityError:
                print(f"Advertencia: Registro duplicado omitido - Delegación: {row['Delegación']}, Tipo: {row['tipo']}")
                continue
            except Exception as e:
                print(f"Error al insertar registro: {e}")
                continue
        
        conn.commit()
        print(f"Se insertaron {total_insertados}/{len(df)} registros en la tabla {tabla_destino}")
        
    except Exception as e:
        print(f"Error al conectar con SQL Server: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

def procesar_carpeta(carpeta):
    #Busca todos los xlsx dentro de la ruta para insertar todos los registros
    if not os.path.exists(carpeta):
        print(f"Error: La carpeta {carpeta} no existe")
        return
    
    archivos_excel = [f for f in os.listdir(carpeta) if f.endswith('.xlsx')]
    
    if not archivos_excel:
        print(f"No se encontraron archivos Excel (.xlsx) en la carpeta {carpeta}")
        return
    
    total_procesados = 0
    total_insertados = 0
    
    for archivo in archivos_excel:
        ruta_completa = os.path.join(carpeta, archivo)
        df = procesar_excel(ruta_completa)
        
        if df is not None:
            insertar_dataframe_sql(df, CONEXION_SQL, TABLA_DESTINO)
            total_procesados += 1
            total_insertados += len(df)
    
    print(f"\nResumen:")
    print(f"- Archivos procesados: {total_procesados}/{len(archivos_excel)}")
    print(f"- Total registros insertados: {total_insertados}")

if __name__ == "__main__":
    procesar_carpeta(CARPETA_REPORTES)