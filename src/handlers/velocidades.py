import os
import re
import datetime
import pandas as pd
import openpyxl
import holidays
from typing import Dict, Tuple, List, Optional, Any

# Mapeo de meses en español
MESES_MAP = {
    "enero": (1, "ene", "Enero"),
    "febrero": (2, "feb", "Febrero"),
    "marzo": (3, "mar", "Marzo"),
    "abril": (4, "abr", "Abril"),
    "mayo": (5, "may", "Mayo"),
    "junio": (6, "jun", "Junio"),
    "julio": (7, "jul", "Julio"),
    "agosto": (8, "ago", "Agosto"),
    "septiembre": (9, "sep", "Septiembre"),
    "setiembre": (9, "sep", "Septiembre"),
    "octubre": (10, "oct", "Octubre"),
    "noviembre": (11, "nov", "Noviembre"),
    "diciembre": (12, "dic", "Diciembre")
}

DEFAULT_PO_A5_PATH = os.path.join("data", "tasacop", "PO_XIII_RM_SURPONIENTE_UN790_NORMAL_2029_4_A5__4.xlsx")


def get_base_service(variant: str) -> str:
    """Remueve sufijos como _I, _R, _R_, _I_ al final del nombre de la variante."""
    return re.sub(r'_(I|R|R_)$', '', str(variant)).strip()


def load_sheet_with_dynamic_header(file_path: str, sheet_name: str) -> pd.DataFrame:
    """Carga una hoja de Excel buscando dinámicamente la fila de cabecera que contiene 'servicio'."""
    wb = openpyxl.load_workbook(file_path, read_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"No se encontró la hoja '{sheet_name}' en el archivo {file_path}")
        
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    
    header_row_idx = 0
    for idx, r in enumerate(rows):
        if r and any(x is not None and str(x).strip().lower() == 'servicio' for x in r):
            header_row_idx = idx
            break
            
    headers = [str(h).strip() if h is not None else f"Col{i}" for i, h in enumerate(rows[header_row_idx])]
    data = rows[header_row_idx+1:]
    return pd.DataFrame(data, columns=headers)


def parse_time_to_seconds(val: Any) -> Optional[int]:
    """Convierte un valor de hora (time o string 'HH:MM:SS') a segundos continuos del día."""
    if pd.isnull(val):
        return None
    if isinstance(val, datetime.time):
        return val.hour * 3600 + val.minute * 60 + val.second
    if isinstance(val, str) and val.strip() != "":
        try:
            parts = val.strip().split(':')
            if len(parts) == 3:
                h, m, s = map(int, parts)
                return h * 3600 + m * 60 + s
            elif len(parts) == 2:
                h, m = map(int, parts)
                return h * 3600 + m * 60
        except Exception:
            return None
    return None


def _construir_mapas_desde_df_pc(df_pc: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Tuple[List[float], float, int]], Dict[str, int]]:
    """Construye las estructuras de parámetros, distancias e índices de PO a partir de df_pc."""
    df_pc['Distancia al origen'] = pd.to_numeric(df_pc['Distancia al origen'], errors='coerce')
    df_pc['Correlativo Punto de Control'] = pd.to_numeric(df_pc['Correlativo Punto de Control'], errors='coerce')
    df_pc['Sentido'] = pd.to_numeric(df_pc['Sentido'], errors='coerce')
    
    distances_map = {}
    po_lookup = {}
    param_rows = []
    
    servicios = df_pc['Servicio'].dropna().unique()
    for serv in servicios:
        serv_str = str(serv).strip()
        for sentido_num, sentido_str in [(0, "Ida"), (1, "Reg"), (1, "Regreso")]:
            ss_key = f"{serv_str}_{sentido_str}"
            df_filtered = df_pc[(df_pc['Servicio'] == serv_str) & (df_pc['Sentido'] == (0 if sentido_num == 0 else 1))]
            if df_filtered.empty:
                continue
                
            df_filtered = df_filtered.sort_values('Correlativo Punto de Control')
            ptos_ctrl_fin = int(df_filtered['Correlativo Punto de Control'].max())
            po_lookup[ss_key] = ptos_ctrl_fin
            po_lookup[f"{serv_str}_I_{sentido_str}"] = ptos_ctrl_fin
            po_lookup[f"{serv_str}_R_{sentido_str}"] = ptos_ctrl_fin
            po_lookup[f"{serv_str}_R__{sentido_str}"] = ptos_ctrl_fin
            
            pci_dist = df_filtered.iloc[0]['Distancia al origen']
            pcf_dist = df_filtered.iloc[-1]['Distancia al origen']
            dist_pci_pcf = (pcf_dist - pci_dist) / 1000.0
            
            param_rows.append({
                'SS': ss_key,
                'dist/exp': round(dist_pci_pcf, 2),
                'ptos_ctrl_fin': ptos_ctrl_fin,
                'dist_pci_pcf': dist_pci_pcf
            })
            
            df_opp = df_pc[(df_pc['Servicio'] == serv_str) & (df_pc['Sentido'] == (1 if sentido_num == 0 else 0))].copy()
            if not df_opp.empty:
                df_opp = df_opp.sort_values('Correlativo Punto de Control')
                d_opp_1 = float(df_opp.iloc[0]['Distancia al origen'])
            else:
                d_opp_1 = 0.0
                
            d_curr = [float(x) for x in df_filtered['Distancia al origen'].tolist()]
            N = len(d_curr)
            distances_map[ss_key] = (d_curr, d_opp_1, N)
            distances_map[f"{serv_str}_I_{sentido_str}"] = (d_curr, d_opp_1, N)
            distances_map[f"{serv_str}_R_{sentido_str}"] = (d_curr, d_opp_1, N)
            distances_map[f"{serv_str}_R__{sentido_str}"] = (d_curr, d_opp_1, N)

    df_params = pd.DataFrame(param_rows).drop_duplicates(subset=['SS'])
    return df_params, distances_map, po_lookup


def cargar_parametros_po(po_a5_path: Optional[str] = None, operador: str = "tasacop") -> Tuple[pd.DataFrame, Dict[str, Tuple[List[float], float, int]], Dict[str, int]]:
    """
    Carga los puntos de control del PO A5.
    1. Intenta consultar directamente la tabla 'puntos_control_po' en SQLite (prioritario para VM sin Excel).
    2. Si no hay datos en SQLite, carga desde el archivo Excel po_a5_path como fallback.
    """
    # 1. Intentar desde SQLite
    try:
        from sqlmodel import Session, select
        from database import engine, PuntoControlPO
        with Session(engine) as session:
            stmt = select(PuntoControlPO).where(PuntoControlPO.operador == operador)
            records = session.exec(stmt).all()
            if records:
                data = [{
                    'Servicio': r.servicio,
                    'Sentido': r.sentido,
                    'Correlativo Punto de Control': r.correlativo,
                    'Distancia al origen': r.distancia_origen
                } for r in records]
                df_pc = pd.DataFrame(data)
                return _construir_mapas_desde_df_pc(df_pc)
    except Exception as e:
        print(f"⚠️ Aviso al consultar puntos_control_po en SQLite: {e}")

    # 2. Fallback a archivo Excel
    if po_a5_path and os.path.exists(po_a5_path):
        df_pc = load_sheet_with_dynamic_header(po_a5_path, 'PC')
        return _construir_mapas_desde_df_pc(df_pc)
        
    raise ValueError("No se encontraron puntos de control del PO A5 ni en la base de datos SQLite ni en disco.")


def obtener_expediciones_tasacop_mes_db(mes_num: int, anio: int) -> Optional[pd.DataFrame]:
    """Extrae las expediciones de Tasacop del mes y año indicados directamente desde la tabla expediciones en SQLite."""
    try:
        from sqlmodel import Session, select, or_
        from database import engine, Expedicion
        with Session(engine) as session:
            mes_str_2d = f"{mes_num:02d}"
            patron_iso = f"{anio}-{mes_str_2d}%"
            patron_dmy = f"%/{mes_str_2d}/{anio}%"
            patron_dmy_dash = f"%-{mes_str_2d}-{anio}%"
            
            stmt = select(Expedicion).where(
                Expedicion.operador == 'tasacop',
                or_(
                    Expedicion.Fecha.like(patron_iso),
                    Expedicion.Fecha.like(patron_dmy),
                    Expedicion.Fecha.like(patron_dmy_dash)
                )
            )
            records = session.exec(stmt).all()
            if not records:
                return None
                
            rows = []
            for r in records:
                row_dict = {
                    'Fecha': r.Fecha,
                    'Bus': r.Bus,
                    'Conductor': r.Chofer,
                    'Con Despacho Asociado': None,
                    'Variante': r.Variante,
                    'Período': r.Periodo,
                    'Dirección': r.Sentido,
                    'Estado': r.Estado,
                    'Causa': r.Causa,
                    'Tipo Demanda': r.Tipo_demanda,
                    'Frecuencia Exigida': r.Frecuencia,
                }
                for i in range(1, 23):
                    p_val = getattr(r, f"p{i}", None)
                    row_dict[f"{i:02d}"] = p_val
                rows.append(row_dict)
                
            df = pd.DataFrame(rows)
            return df
    except Exception as e:
        print(f"⚠️ Aviso al consultar expediciones en SQLite: {e}")
        return None


def buscar_archivo_expediciones_tasacop(mes: str, anio: int = 2026) -> Optional[str]:
    """Busca el archivo de expediciones de Tasacop en data/tasacop para el mes y año indicados."""
    mes_info = MESES_MAP.get(mes.lower().strip())
    if not mes_info:
        return None
        
    mes_num, mes_abbr, mes_nombre = mes_info
    yy = str(anio)[-2:]
    
    tasacop_dir = os.path.join("data", "tasacop")
    if not os.path.exists(tasacop_dir):
        return None
        
    posibles_carpetas = [
        f"{mes_nombre}{yy}",
        f"{mes.capitalize()}{yy}",
        f"{mes.lower()}{yy}",
        mes_nombre,
        mes.lower()
    ]
    
    for c in posibles_carpetas:
        c_path = os.path.join(tasacop_dir, c)
        if os.path.isdir(c_path):
            archivos = [
                os.path.join(c_path, f) for f in os.listdir(c_path)
                if (f.startswith("expediciones_") or "tasaco" in f.lower()) and f.endswith(".xlsx")
            ]
            if archivos:
                return archivos[0]
                
    # Búsqueda recursiva si no se encontró en carpetas directas
    for root, _, files in os.walk(tasacop_dir):
        for f in files:
            f_lower = f.lower()
            if (mes_nombre.lower() in f_lower or mes_abbr in f_lower) and f_lower.endswith(".xlsx") and "expediciones" in f_lower:
                return os.path.join(root, f)
                
    return None


def procesar_modelo_cinematico(
    df_in: pd.DataFrame,
    po_a5_path: Optional[str] = DEFAULT_PO_A5_PATH,
    operador: str = "tasacop"
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    """
    Ejecuta la formulación matemática y cinemática completa:
    1. Limpieza de columnas y recálculo de Tipo de Día con feriados chilenos.
    2. Carga de parámetros de distancias del PO A5 (desde SQLite prioritariamente o Excel fallback).
    3. Construcción de la matriz Datos (51 columnas).
    4. Descomposición por micro-tramos en desref con interpolación proporcional y velocidades.
    5. Retorna (df_datos, df_desref, df_params, metricas_resumen).
    """

    # Estandarización de nombres de columnas de entrada
    col_renames = {}
    for col in df_in.columns:
        c_clean = str(col).strip()
        if c_clean.lower() in ['tipo de dia', 'tipo de día', 'tipo dia']:
            col_renames[col] = 'Tipo de Día'
        elif c_clean.lower() in ['direccion', 'dirección']:
            col_renames[col] = 'Dirección'
        elif c_clean.lower() in ['periodo', 'período']:
            col_renames[col] = 'Período'
        elif c_clean.lower() in ['frecuencia', 'frecuencia exigida']:
            col_renames[col] = 'Frecuencia Exigida'
        elif c_clean.lower() in ['tipo demanda', 'tipo de demanda']:
            col_renames[col] = 'Tipo Demanda'
        elif c_clean.lower() in ['con despacho asociado']:
            col_renames[col] = 'Con Despacho Asociado'
        else:
            col_renames[col] = c_clean
            
    df_in = df_in.rename(columns=col_renames).copy()

    # Normalizar checkpoints 01 a 22
    for i in range(1, 23):
        candidates = [str(i), f"{i:02d}", f"p{i}"]
        found = False
        for c in candidates:
            if c in df_in.columns:
                if c != f"{i:02d}":
                    df_in[f"{i:02d}"] = df_in[c]
                found = True
                break
        if not found:
            df_in[f"{i:02d}"] = None

    # Normalizar columnas esenciales faltantes
    for col in ['Bus', 'Conductor', 'Con Despacho Asociado', 'Causa', 'Tipo Demanda', 'Frecuencia Exigida']:
        if col not in df_in.columns:
            df_in[col] = None

    # Normalizar Estado
    if 'Estado' in df_in.columns:
        df_in['Estado'] = df_in['Estado'].astype(str).apply(
            lambda x: "Válida" if "v" in x.lower() and "inv" not in x.lower() else ("Inválida" if "inv" in x.lower() else x)
        )
    else:
        df_in['Estado'] = 'Válida'

    # Normalizar Dirección (Ida / Reg)
    if 'Dirección' in df_in.columns:
        df_in['Dirección'] = df_in['Dirección'].astype(str).apply(
            lambda x: "Ida" if "ida" in x.lower() or x.strip() == "0" else "Reg"
        )
    else:
        df_in['Dirección'] = "Ida"

    # Recalcular Tipo de Día con festivos chilenos
    df_in['Fecha'] = pd.to_datetime(df_in['Fecha'], errors='coerce')
    unique_years = [y for y in df_in['Fecha'].dt.year.dropna().unique().tolist() if y > 2000]
    if not unique_years:
        unique_years = [2026]
    chile_holidays = holidays.Chile(years=unique_years)

    def calculate_tipo_dia(row):
        if pd.isnull(row['Fecha']):
            return 'DLN'
        dt = row['Fecha'].date()
        if dt in chile_holidays or dt.weekday() == 6:  # Domingo es 6
            return 'DOM'
        elif dt.weekday() == 5:  # Sábado es 5
            return 'SAB'
        else:
            return 'DLN'

    df_in['Tipo de Día'] = df_in.apply(calculate_tipo_dia, axis=1)

    # Cargar parámetros PO
    df_params, distances_map, po_lookup = cargar_parametros_po(po_a5_path)

    df_in['SS_temp'] = df_in['Variante'].astype(str) + "_" + df_in['Dirección'].astype(str)

    chk_cols = [f"{i:02d}" for i in range(1, 23)]
    df_in_chks_sec = {col: df_in[col].apply(parse_time_to_seconds) for col in chk_cols}

    datos_rows_to_write = []
    desref_rows_to_write = []

    for idx, row in df_in.iterrows():
        row_id = idx + 1

        times_sec = []
        times_obj = []
        for col in chk_cols:
            sec = df_in_chks_sec[col].iloc[idx]
            times_sec.append(sec)
            val = row[col]
            if pd.isnull(val):
                times_obj.append(None)
            else:
                if isinstance(val, datetime.time):
                    times_obj.append(val)
                elif isinstance(val, str) and val.strip() != "":
                    try:
                        h, m, s = map(int, val.split(':'))
                        times_obj.append(datetime.time(h, m, s))
                    except Exception:
                        times_obj.append(None)
                else:
                    times_obj.append(None)

        # Último punto de control registrado físicamente
        last_chk_idx = 0
        for i in range(22, 0, -1):
            if times_obj[i - 1] is not None:
                last_chk_idx = i
                break
        puntos_control_OP = last_chk_idx

        ss = row['SS_temp']
        base_ss = f"{get_base_service(row['Variante'])}_{row['Dirección']}"
        puntos_control_PO = po_lookup.get(ss) or po_lookup.get(base_ss)

        # Criterio de muestra: Estado Válida y llegó al último punto exigido por PO
        muestra = 0
        if row['Estado'] == 'Válida' and puntos_control_PO is not None and puntos_control_OP == puntos_control_PO:
            muestra = 1

        fecha_val = row['Fecha']
        if pd.notnull(fecha_val):
            año = fecha_val.year
            mes = fecha_val.month
            día = fecha_val.day
            fecha_dt = fecha_val.date()
        else:
            año, mes, día = 2026, 1, 1
            fecha_dt = datetime.date(2026, 1, 1)

        tipo_dia_periodo = f"{row['Tipo de Día']}_{row['Período']}"
        ss_tipo_dia = f"{ss}_{row['Tipo de Día']}"
        ss_tipo_dia_periodo = f"{ss_tipo_dia}_{row['Período']}"

        # Hora de paso final PC_F
        pc_f_val = None
        if muestra == 1 and puntos_control_PO is not None:
            po_limit = int(puntos_control_PO)
            if 1 <= po_limit <= 22:
                pc_f_val = times_obj[po_limit - 1]

        datos_row_vals = [
            fecha_dt,
            row['Bus'],
            row['Conductor'],
            row['Con Despacho Asociado'],
            row['Variante'],
            row['Período'],
            row['Tipo de Día'],
            row['Dirección'],
            row['Estado'],
            row['Causa'] if pd.notnull(row['Causa']) else None,
            row['Tipo Demanda'] if pd.notnull(row['Tipo Demanda']) else None,
            row['Frecuencia Exigida']
        ] + times_obj + [
            row_id,
            año,
            mes,
            día,
            ss,
            puntos_control_OP,
            puntos_control_PO,
            muestra,
            tipo_dia_periodo,
            ss_tipo_dia,
            row['Período'],
            row['Tipo de Día'],
            ss_tipo_dia_periodo,
            row['Variante'],
            row['Dirección'],
            row['Tipo de Día'],
            pc_f_val
        ]
        datos_rows_to_write.append(datos_row_vals)

        # Distancias y cálculo cinemático de desref
        lookup_dist = distances_map.get(ss) or distances_map.get(base_ss)
        if lookup_dist:
            d_curr, d_opp_1, N = lookup_dist
        else:
            d_curr, d_opp_1, N = ([], 0.0, 0)

        step_metrics = {c: {'dist_km': 0.0, 'delta_TV': None, 'delta_TV_min': None, 'velocidad': None} for c in range(1, 23)}

        if N >= 1 and d_curr:
            seg_dists_m = {1: d_curr[0]}
            for c in range(2, N + 1):
                seg_dists_m[c] = d_curr[c - 1] - d_curr[c - 2]
            if N + 1 <= 22:
                seg_dists_m[N + 1] = d_opp_1
            for c in range(N + 2, 23):
                seg_dists_m[c] = 0.0

            for c in range(1, 23):
                step_metrics[c]['dist_km'] = seg_dists_m.get(c, 0.0) / 1000.0

            recorded = [c for c in range(1, N + 1) if times_sec[c - 1] is not None]

            # Interpolación entre puntos de control registrados
            for i in range(len(recorded) - 1):
                c_a = recorded[i]
                c_b = recorded[i + 1]
                t_a = times_sec[c_a - 1]
                t_b = times_sec[c_b - 1]

                diff_sec = t_b - t_a
                if diff_sec >= 0:
                    dist_ab_m = d_curr[c_b - 1] - (d_curr[c_a - 2] if c_a > 1 else 0.0)
                    if dist_ab_m > 0 and diff_sec > 0:
                        v_ab = (dist_ab_m / 1000.0) / (diff_sec / 3600.0)
                    else:
                        v_ab = None

                    for c in range(c_a + 1, c_b + 1):
                        seg_m = seg_dists_m.get(c, 0.0)
                        if dist_ab_m > 0 and diff_sec > 0:
                            t_seg_sec = diff_sec * (seg_m / dist_ab_m)
                            step_metrics[c]['delta_TV'] = t_seg_sec / 86400.0
                            step_metrics[c]['delta_TV_min'] = round(t_seg_sec / 60.0) / 1440.0
                            step_metrics[c]['velocidad'] = v_ab
                        elif diff_sec == 0:
                            step_metrics[c]['delta_TV'] = 0.0
                            step_metrics[c]['delta_TV_min'] = 0.0
                            step_metrics[c]['velocidad'] = None

            # Tramo 1: Origen -> PC 1
            first_op_speed = None
            for c in range(2, N + 1):
                if step_metrics[c]['velocidad'] is not None and step_metrics[c]['velocidad'] > 0:
                    first_op_speed = step_metrics[c]['velocidad']
                    break

            d1_km = step_metrics[1]['dist_km']
            if first_op_speed is not None and first_op_speed > 0 and d1_km > 0:
                step_metrics[1]['velocidad'] = first_op_speed
                t1_sec = (d1_km / first_op_speed) * 3600.0
                step_metrics[1]['delta_TV'] = t1_sec / 86400.0
                step_metrics[1]['delta_TV_min'] = round(t1_sec / 60.0) / 1440.0

            # Tramo Final: PC N -> Destino (c = N + 1)
            last_op_speed = None
            for c in range(N, 1, -1):
                if step_metrics[c]['velocidad'] is not None and step_metrics[c]['velocidad'] > 0:
                    last_op_speed = step_metrics[c]['velocidad']
                    break

            if N + 1 <= 22:
                d_final_km = step_metrics[N + 1]['dist_km']
                if last_op_speed is not None and last_op_speed > 0 and d_final_km > 0:
                    step_metrics[N + 1]['velocidad'] = last_op_speed
                    t_final_sec = (d_final_km / last_op_speed) * 3600.0
                    step_metrics[N + 1]['delta_TV'] = t_final_sec / 86400.0
                    step_metrics[N + 1]['delta_TV_min'] = round(t_final_sec / 60.0) / 1440.0

        for c in range(1, 23):
            hora_paso = times_obj[c - 1]
            m = step_metrics.get(c, {})
            desref_row_vals = [
                row_id,
                c,
                f"{c:02d}",
                hora_paso,
                m.get('delta_TV'),
                muestra,
                año,
                mes,
                día,
                ss,
                puntos_control_OP,
                puntos_control_PO,
                muestra,
                tipo_dia_periodo,
                ss_tipo_dia,
                row['Período'],
                row['Tipo de Día'],
                ss_tipo_dia_periodo,
                row['Variante'],
                row['Dirección'],
                m.get('delta_TV_min'),
                m.get('velocidad')
            ]
            desref_rows_to_write.append(desref_row_vals)

    headers_datos = [
        'Fecha', 'Bus', 'Conductor', 'Con Despacho Asociado', 'Variante', 'Período', 'Tipo de Día', 'Dirección', 'Estado',
        'Causa', 'Tipo Demanda', 'Frecuencia Exigida', '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11',
        '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', 'ID', 'año', 'mes', 'día', 'SS', 'puntos_control_OP',
        'puntos_control_PO', 'muestra', 'Tipo de Día_Período', 'SS_Tipo de Día', 'Período.1', 'Tipo de Día.1', 'SS_Tipo de Día_Período',
        'Variante.1', 'Dirección.1', 'Tipo Día', 'PC_F'
    ]

    headers_desref = [
        'fila_ID', 'columna', 'pto_ctrl', 'hora_paso', 'delta_TV', 'muestra', 'año', 'mes', 'día', 'SS',
        'puntos_control_OP', 'puntos_control_PO', 'muestra.1', 'Tipo de Día_Período', 'SS_Tipo de Día',
        'Período', 'Tipo de Día', 'SS_Tipo de Día_Período', 'Variante', 'Dirección', 'delta_TV_min', 'velocidad'
    ]

    df_datos_out = pd.DataFrame(datos_rows_to_write, columns=headers_datos)
    df_desref_out = pd.DataFrame(desref_rows_to_write, columns=headers_desref)

    # Calcular estadísticas de resumen
    total_exp = len(df_datos_out)
    total_muestra = int((df_datos_out['muestra'] == 1).sum())
    pct_muestra = (total_muestra / total_exp * 100.0) if total_exp > 0 else 0.0
    
    vel_series = df_desref_out[df_desref_out['muestra'] == 1]['velocidad'].dropna()
    vel_promedio = float(vel_series.mean()) if not vel_series.empty else 0.0

    metricas_resumen = {
        'total_expediciones': total_exp,
        'total_muestra': total_muestra,
        'pct_muestra': pct_muestra,
        'vel_promedio': vel_promedio
    }

    return df_datos_out, df_desref_out, df_params, metricas_resumen
