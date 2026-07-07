import os
import glob
import pandas as pd
import numpy as np
from datetime import datetime
from sqlmodel import Session, select
from database import init_db, engine, Expedicion, Anexo1

def get_operador(path: str) -> str:
    path_lower = path.lower()
    if "lider" in path_lower:
        return "lider"
    elif "tasacop" in path_lower:
        return "tasacop"
    elif "toptur" in path_lower:
        return "toptur"
    return "desconocido"

def limpiar_nan(val):
    if pd.isna(val) or val is None or str(val).strip().lower() in ("nan", "nat", ""):
        return None
    return val

def importar_anexo1(file_path: str):
    operador = get_operador(file_path)
    print(f"📥 Importando Anexo 1 desde: {file_path} (Operador: {operador})...")
    
    xl = pd.ExcelFile(file_path)
    sheet_names = [s for s in xl.sheet_names if s not in ('TAPA', 'Servicios')]
    
    records_to_add = []
    
    for sheet_name in sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            # Validar que tenga el tamaño mínimo
            if len(df) < 36:
                print(f"⚠️ Ignorando hoja '{sheet_name}': tamaño insuficiente ({len(df)} filas)")
                continue
                
            # Extraer metadatos de la hoja (Fila 6, 0-indexed)
            servicio = str(df.iloc[6, 1]).strip()
            sentido = str(df.iloc[6, 2]).strip()
            
            # Recorrer periodos (Filas 12 a 35)
            for r_idx in range(12, 36):
                periodo_val = df.iloc[r_idx, 1]
                if pd.isna(periodo_val):
                    continue
                periodo = int(periodo_val)
                horario = str(df.iloc[r_idx, 2]).strip()
                
                # Mapeo de columnas por día
                dias_config = [
                    ("Laboral", 3, 4),
                    ("Sábado", 5, 6),
                    ("Domingo / Festivo", 7, 8)
                ]
                
                for tipo_dia, td_col, freq_col in dias_config:
                    td_val = limpiar_nan(df.iloc[r_idx, td_col])
                    freq_val = limpiar_nan(df.iloc[r_idx, freq_col])
                    
                    if td_val is not None or freq_val is not None:
                        # Si freq_val es un float o int, convertirlo
                        freq_esperada = float(freq_val) if freq_val is not None else None
                        
                        record = Anexo1(
                            operador=operador,
                            servicio=servicio,
                            sentido=sentido,
                            periodo=periodo,
                            horario=horario,
                            tipo_dia=tipo_dia,
                            tipo_demanda=str(td_val) if td_val is not None else None,
                            frecuencia_esperada=freq_esperada
                        )
                        records_to_add.append(record)
                        
        except Exception as e:
            print(f"❌ Error procesando hoja '{sheet_name}' de '{file_path}': {e}")
            
    # Guardar en base de datos
    if records_to_add:
        with Session(engine) as session:
            session.add_all(records_to_add)
            session.commit()
        print(f"✅ Se importaron {len(records_to_add)} registros de Anexo 1.")
    else:
        print("⚠️ No se encontraron registros válidos de Anexo 1.")

def importar_expediciones(file_path: str):
    operador = get_operador(file_path)
    print(f"📥 Importando Expediciones desde: {file_path} (Operador: {operador})...")
    
    # Comprobar formato leyendo los primeros 100 bytes
    is_html = False
    try:
        with open(file_path, 'rb') as f:
            header_bytes = f.read(100)
            if b"html" in header_bytes or b"HTML" in header_bytes:
                is_html = True
    except Exception as e:
        print(f"❌ Error leyendo archivo para determinar formato {file_path}: {e}")
        return

    try:
        if is_html:
            dfs = pd.read_html(file_path)
            df = dfs[0]
        else:
            df = pd.read_excel(file_path)
    except Exception as e:
        print(f"❌ Error al leer el archivo {file_path}: {e}")
        return
        
    records_to_add = []
    
    # Detectar el tipo de mapeo según las columnas del archivo
    columns_set = set(df.columns)
    
    # Mapeo según el formato (HTML vs Excel de Tasacop)
    for idx, row in df.iterrows():
        try:
            if is_html:
                # Formato HTML (Lider y Toptur)
                chofer = limpiar_nan(row.get('Chofer'))
                sentido = str(row.get('Sentido', ''))
                periodo = int(row.get('Periodo', 0)) if not pd.isna(row.get('Periodo')) else 0
                tipo_demanda = limpiar_nan(row.get('Tipo demanda'))
                frecuencia = float(row.get('Frecuencia')) if not pd.isna(row.get('Frecuencia')) else None
                inicio_exp = str(row.get('Inicio Expedicion', ''))
                fin_exp = str(row.get('Fin Expedicion', ''))
                servicio = str(row.get('Servicio', ''))
                propietario = limpiar_nan(row.get('Propietario'))
                
                # Checkpoints son de "1" a "22"
                pois = {}
                for p_num in range(1, 23):
                    pois[f"p{p_num}"] = limpiar_nan(row.get(str(p_num)))
            else:
                # Formato Excel estándar (Tasacop)
                chofer = limpiar_nan(row.get('Conductor'))
                sentido = str(row.get('Dirección', ''))
                periodo = int(row.get('Período', 0)) if not pd.isna(row.get('Período')) else 0
                tipo_demanda = limpiar_nan(row.get('Tipo Demanda'))
                frecuencia = float(row.get('Frecuencia Exigida')) if not pd.isna(row.get('Frecuencia Exigida')) else None
                inicio_exp = ""
                fin_exp = ""
                servicio = str(row.get('Variante', ''))  # Variante actúa como servicio
                propietario = None
                
                # Checkpoints tienen formato "01" a "22"
                pois = {}
                for p_num in range(1, 23):
                    pois[f"p{p_num}"] = limpiar_nan(row.get(f"{p_num:02d}"))

            exp_data = {
                'operador': operador,
                'Fecha': str(row.get('Fecha', '')),
                'Inicio_Expedicion': inicio_exp,
                'Fin_Expedicion': fin_exp,
                'Folio_TS': limpiar_nan(row.get('Folio TS')),
                'ID_Exp': int(row.get('ID Exp')) if not pd.isna(row.get('ID Exp')) else None,
                'Bus': limpiar_nan(row.get('Bus')),
                'Chofer': chofer,
                'Propietario': propietario,
                'Variante': str(row.get('Variante', '')),
                'Servicio': servicio,
                'Periodo': periodo,
                'Sentido': sentido,
                'Estado': limpiar_nan(row.get('Estado')),
                'Causa': limpiar_nan(row.get('Causa')),
                'Tipo_demanda': tipo_demanda,
                'Frecuencia': frecuencia,
                'Vel_Promedio': float(row.get('Vel.Promedio')) if not pd.isna(row.get('Vel.Promedio')) else None,
                'Vel_Maxima': float(row.get('Vel.Maxima')) if not pd.isna(row.get('Vel.Maxima')) else None,
            }
            
            # Combinar datos base con checkpoints
            exp_data.update(pois)
            
            exp = Expedicion(**exp_data)
            records_to_add.append(exp)
            
        except Exception as e:
            print(f"⚠️ Error procesando fila {idx} en '{file_path}': {e}")
            
    # Guardar en base de datos
    if records_to_add:
        # Limpiar registros anteriores del mismo operador y archivos similares para evitar duplicados
        # (Esto es opcional, pero ayuda a mantener la base limpia en re-ejecuciones)
        # with Session(engine) as session:
        #     session.execute(f"DELETE FROM expediciones WHERE operador = '{operador}'")
        #     session.commit()
            
        batch_size = 1000
        total_records = len(records_to_add)
        
        with Session(engine) as session:
            for i in range(0, total_records, batch_size):
                batch = records_to_add[i:i+batch_size]
                session.add_all(batch)
                session.commit()
        print(f"✅ Se importaron {total_records} expediciones.")
    else:
        print("⚠️ No se encontraron expediciones válidas.")

def main():
    print("🚀 Iniciando importación de datos en SQLite...")
    # Asegurar que las tablas están creadas
    init_db()
    
    # Limpiar tablas para evitar duplicados
    from sqlalchemy import text
    with Session(engine) as session:
        session.execute(text("DELETE FROM expediciones"))
        session.execute(text("DELETE FROM anexo_1"))
        session.commit()
    print("🧹 Tablas 'expediciones' y 'anexo_1' limpiadas con éxito.")
    
    # 1. Buscar y procesar Anexos 1
    anexo_files = glob.glob("data/**/*A1_*.xlsx", recursive=True)
    print(f"Se encontraron {len(anexo_files)} archivos de Anexo 1.")
    for f in anexo_files:
        importar_anexo1(f)
        
    # 2. Buscar y procesar Expediciones
    # Cambiamos el patrón glob para que incluya .xlsx y archivos sin prefijo de operador
    exp_files = glob.glob("data/**/*expediciones*.xls*", recursive=True)
    print(f"\nSe encontraron {len(exp_files)} archivos de expediciones.")
    for f in exp_files:
        importar_expediciones(f)
        
    print("\n🏁 Proceso de importación finalizado con éxito!")

if __name__ == "__main__":
    main()
