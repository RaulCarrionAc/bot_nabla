import os
import io
import tempfile
import pandas as pd
import openpyxl
from typing import Dict, Any, List, Optional
from sqlalchemy import text
from sqlmodel import Session
from database import engine, Anexo1, PuntoControlPO, Anexo5Record


def limpiar_nan(val: Any) -> Optional[Any]:
    if pd.isna(val) or val is None or str(val).strip().lower() in ("nan", "nat", ""):
        return None
    return val


def actualizar_anexo_1_desde_bytes(excel_bytes: bytes, operador: str) -> Dict[str, Any]:
    """
    Parsea y actualiza el Anexo 1 (Frecuencias) para una empresa/operador específico,
    sobreescribiendo los registros existentes en la tabla 'anexo_1' de SQLite.
    100% en memoria/base de datos sin dejar archivos residuales en disco.
    """
    operador_clean = operador.strip().lower()
    if "tasa" in operador_clean:
        operador_clean = "tasacop"
    
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(excel_bytes)
        tmp_path = tmp.name
        
    try:
        xl = pd.ExcelFile(tmp_path)
        sheet_names = [s for s in xl.sheet_names if s not in ('TAPA', 'Servicios', 'Resumen')]
        
        records_to_add: List[Anexo1] = []
        servicios_set = set()
        
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(tmp_path, sheet_name=sheet_name, header=None)
                if len(df) < 15:
                    continue
                    
                servicio = str(df.iloc[6, 1]).strip()
                sentido = str(df.iloc[6, 2]).strip()
                servicios_set.add(f"{servicio}_{sentido}")
                
                max_r = min(len(df), 36)
                for r_idx in range(12, max_r):
                    periodo_val = df.iloc[r_idx, 1]
                    if pd.isna(periodo_val):
                        continue
                    try:
                        periodo = int(periodo_val)
                    except Exception:
                        continue
                        
                    horario = str(df.iloc[r_idx, 2]).strip()
                    
                    dias_config = [
                        ("Laboral", 3, 4),
                        ("Sábado", 5, 6),
                        ("Domingo / Festivo", 7, 8)
                    ]
                    
                    for tipo_dia, td_col, freq_col in dias_config:
                        if td_col < len(df.columns) and freq_col < len(df.columns):
                            td_val = limpiar_nan(df.iloc[r_idx, td_col])
                            freq_val = limpiar_nan(df.iloc[r_idx, freq_col])
                            
                            if td_val is not None or freq_val is not None:
                                try:
                                    freq_esperada = float(freq_val) if freq_val is not None else None
                                except Exception:
                                    freq_esperada = None
                                    
                                record = Anexo1(
                                    operador=operador_clean,
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
                print(f"⚠️ Error procesando hoja '{sheet_name}' de Anexo 1: {e}")
                
        xl.close()
        
        if not records_to_add:
            return {
                "success": False,
                "error": "No se encontraron tablas válidas de frecuencias en las hojas del archivo."
            }
            
        # Sobreescribir en la base de datos SQLite
        with Session(engine) as session:
            session.execute(text(f"DELETE FROM anexo_1 WHERE LOWER(operador) = '{operador_clean}'"))
            session.add_all(records_to_add)
            session.commit()
            
        return {
            "success": True,
            "tipo": "A1",
            "operador": operador_clean,
            "total_registros": len(records_to_add),
            "servicios_actualizados": len(servicios_set)
        }
        
    finally:
        try:
            if 'xl' in locals():
                xl.close()
        except Exception:
            pass
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def actualizar_anexo_5_desde_bytes(excel_bytes: bytes, operador: str) -> Dict[str, Any]:
    """
    Parsea y actualiza el Anexo 5 (Puntos de Control / PO A5 / LPP) para una empresa/operador específico.
    Sobreescribe las tablas 'puntos_control_po' y 'anexo_5' en SQLite.
    100% en base de datos SQLite sin almacenar archivos en disco.
    """
    operador_clean = operador.strip().lower()
    if "tasa" in operador_clean:
        operador_clean = "tasacop"
    
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(excel_bytes)
        tmp_path = tmp.name
        
    try:
        wb = openpyxl.load_workbook(tmp_path, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        
        records_pc: List[PuntoControlPO] = []
        records_lpp: List[Anexo5Record] = []
        servicios_pc_set = set()
        servicios_lpp_set = set()
        
        # 1. Procesar Hoja 'PC' (Puntos de Control y distancias geodésicas)
        pc_sheet_name = next((s for s in sheet_names if s.strip().upper() == "PC"), None)
        if pc_sheet_name:
            from handlers.velocidades import load_sheet_with_dynamic_header
            df_pc = load_sheet_with_dynamic_header(tmp_path, pc_sheet_name)
            
            for col in df_pc.columns:
                c_clean = str(col).strip().lower()
                if "distancia" in c_clean and "origen" in c_clean:
                    df_pc.rename(columns={col: "Distancia al origen"}, inplace=True)
                elif "correlativo" in c_clean or ("punto" in c_clean and "control" in c_clean):
                    df_pc.rename(columns={col: "Correlativo Punto de Control"}, inplace=True)
                elif "sentido" in c_clean:
                    df_pc.rename(columns={col: "Sentido"}, inplace=True)
                elif "servicio" in c_clean:
                    df_pc.rename(columns={col: "Servicio"}, inplace=True)
            
            if {'Servicio', 'Sentido', 'Correlativo Punto de Control', 'Distancia al origen'}.issubset(set(df_pc.columns)):
                df_pc['Distancia al origen'] = pd.to_numeric(df_pc['Distancia al origen'], errors='coerce')
                df_pc['Correlativo Punto de Control'] = pd.to_numeric(df_pc['Correlativo Punto de Control'], errors='coerce')
                df_pc['Sentido'] = pd.to_numeric(df_pc['Sentido'], errors='coerce')
                
                df_clean_pc = df_pc.dropna(subset=['Servicio', 'Sentido', 'Correlativo Punto de Control', 'Distancia al origen'])
                
                for _, row in df_clean_pc.iterrows():
                    serv_str = str(row['Servicio']).strip()
                    sentido_int = int(row['Sentido'])
                    servicios_pc_set.add(f"{serv_str}_{sentido_int}")
                    
                    record = PuntoControlPO(
                        operador=operador_clean,
                        servicio=serv_str,
                        sentido=sentido_int,
                        correlativo=int(row['Correlativo Punto de Control']),
                        distancia_origen=float(row['Distancia al origen'])
                    )
                    records_pc.append(record)

        # 2. Procesar Hoja 'LPP' (Lista de Pasadas Programadas para IP / Puntualidad)
        lpp_sheet_name = next((s for s in sheet_names if s.strip().upper() == "LPP"), None)
        if lpp_sheet_name:
            df_lpp_raw = pd.read_excel(tmp_path, sheet_name=lpp_sheet_name, header=None)
            
            # Buscar fila de encabezados que contenga 'Servicio' y 'Hora' o 'TPP'
            header_row = 0
            for idx, row in df_lpp_raw.iterrows():
                row_str = " ".join([str(v).lower() for v in row.dropna()])
                if "servicio" in row_str and ("hora" in row_str or "tpp" in row_str or "pasada" in row_str):
                    header_row = idx
                    break
                    
            df_lpp = pd.read_excel(tmp_path, sheet_name=lpp_sheet_name, skiprows=header_row)
            
            col_map = {}
            for c in df_lpp.columns:
                c_clean = str(c).strip().lower().replace("\n", " ")
                if "servicio" in c_clean:
                    col_map[c] = "Servicio"
                elif "sentido" in c_clean:
                    col_map[c] = "Sentido"
                elif "correlativo" in c_clean or ("punto" in c_clean and "control" in c_clean):
                    col_map[c] = "correlativo_pc"
                elif "anterior" in c_clean:
                    col_map[c] = "IPP_anterior"
                elif "programada" in c_clean or "tpp" in c_clean:
                    col_map[c] = "TPP"
                elif "posterior" in c_clean:
                    col_map[c] = "IPP_posterior"
                elif "tipo" in c_clean and "d" in c_clean:
                    col_map[c] = "tipo_dia"
                    
            df_lpp = df_lpp.rename(columns=col_map)
            if 'Servicio' in df_lpp.columns and 'TPP' in df_lpp.columns:
                df_lpp_clean = df_lpp.dropna(subset=['Servicio', 'TPP'])
                for _, row in df_lpp_clean.iterrows():
                    srv = str(row['Servicio']).strip()
                    sen = str(row.get('Sentido', 'Ida')).strip()
                    c_pc = int(row.get('correlativo_pc', 1)) if pd.notna(row.get('correlativo_pc')) else 1
                    tpp_str = str(row.get('TPP', '')).strip()
                    ipp_ant = str(row.get('IPP_anterior', '')).strip() if pd.notna(row.get('IPP_anterior')) else None
                    ipp_post = str(row.get('IPP_posterior', '')).strip() if pd.notna(row.get('IPP_posterior')) else None
                    td = str(row.get('tipo_dia', 'Laboral')).strip() if pd.notna(row.get('tipo_dia')) else 'Laboral'
                    
                    servicios_lpp_set.add(srv)
                    record_lpp = Anexo5Record(
                        operador=operador_clean,
                        Servicio=srv,
                        Sentido=sen,
                        correlativo_pc=c_pc,
                        IPP_anterior=ipp_ant,
                        TPP=tpp_str,
                        IPP_posterior=ipp_post,
                        tipo_dia=td
                    )
                    records_lpp.append(record_lpp)

        # 3. Sobreescribir en la base de datos SQLite
        with Session(engine) as session:
            if records_pc:
                session.execute(text(f"DELETE FROM puntos_control_po WHERE LOWER(operador) = '{operador_clean}'"))
                session.add_all(records_pc)
            if records_lpp:
                session.execute(text(f"DELETE FROM anexo_5 WHERE LOWER(operador) = '{operador_clean}'"))
                session.add_all(records_lpp)
            session.commit()
            
        if not records_pc and not records_lpp:
            return {
                "success": False,
                "error": "No se encontraron tablas válidas de Puntos de Control ('PC') ni de Pasadas Programadas ('LPP') en el archivo."
            }

        return {
            "success": True,
            "tipo": "A5",
            "operador": operador_clean,
            "total_puntos_control": len(records_pc),
            "total_pasadas_lpp": len(records_lpp),
            "servicios_actualizados": len(servicios_pc_set.union(servicios_lpp_set))
        }
        
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
