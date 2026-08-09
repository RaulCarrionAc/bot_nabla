"""
Módulo de cálculo del Indicador de Puntualidad (IP) para el Bot Nabla.

Calcula el IP de acuerdo a la metodología normativa:
- Cruza la Lista de Pasadas Programadas (LPP) del Anexo 5 con la
  Lista de Pasadas Observadas (LPO) de las expediciones telemáticas.
- Asigna tolerancias según intervalos anterior (IPP_ant) y posterior (IPP_post).
- Genera libro Excel con 'Resumen por Servicio' y 'Reporte' de detalle.
"""

import datetime
import io
import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

import holidays
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# Feriados de Chile
FERIADOS_CL = holidays.Chile()

# Mapeo de sentido
MAPEO_SENTIDO = {"Ida": 0, "Reg": 1, "I": 0, "V": 1, "Regreso": 1}

# Columnas del Anexo 5 (LPP)
CAMPOS_A5 = {
    "Correlativo Punto\nde Control": "correlativo_pc",
    "Intervalo Anterior\n(IPPdk-1)": "IPP_anterior",
    "Hora de Pasada Programada\n(TPPdk)": "TPP",
    "Intervalo Posterior\n(IPPdk)": "IPP_posterior",
    "Tipo de Día": "tipo_dia",
}

_PALABRAS_LABEL = (
    "FECHA",
    "TIPO",
    "REGIÓN",
    "REGION",
    "ZONA",
    "UNIDAD",
    "RES N",
    "CON VERSIONES",
    "ESTACIONALIDAD",
)

_ESTACIONALIDADES_CONOCIDAS = ("NORMAL", "ESTIVAL", "A1", "A2")

MESES_MAP: Dict[str, Tuple[int, str, str]] = {
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
    "diciembre": (12, "dic", "Diciembre"),
}

ESTACIONALIDADES_VALIDAS_IP = ("NORMAL", "ESTIVAL")


# --- UTILIDADES DE TIEMPO ---

def a_segundos(valor) -> float:
    """Convierte un valor de hora a segundos continuos del día."""
    if pd.isna(valor):
        return np.nan
    if isinstance(valor, (int, float, np.integer, np.floating)):
        return float(valor)
    if isinstance(valor, pd.Timedelta):
        return valor.total_seconds()
    if hasattr(valor, "hour"):
        seg = getattr(valor, "second", 0)
        micro = getattr(valor, "microsecond", 0) / 1e6
        return valor.hour * 3600 + valor.minute * 60 + seg + micro

    val_str = str(valor).strip()
    partes = val_str.split(":")
    try:
        if len(partes) == 3:
            return int(partes[0]) * 3600 + int(partes[1]) * 60 + float(partes[2])
        elif len(partes) == 2:
            return int(partes[0]) * 3600 + float(partes[1]) * 60
    except (ValueError, TypeError):
        return np.nan

    return np.nan


def clasificar_tipo_dia(fecha) -> str:
    """Clasifica una fecha en 'DL' (Laboral), 'DS' (Sábado) o 'DF' (Domingo/Festivo)."""
    fecha_dt = pd.Timestamp(fecha).date()
    if fecha_dt in FERIADOS_CL:
        return "DF"

    dow = pd.Timestamp(fecha_dt).dayofweek
    if dow == 6:
        return "DF"
    elif dow == 5:
        return "DS"
    else:
        return "DL"


# --- CARGA Y VIGENCIA DE ANEXO 5 (LPP) ---

def _parece_label(valor) -> bool:
    if not isinstance(valor, str):
        return False
    v = valor.strip().upper()
    return any(p in v for p in _PALABRAS_LABEL)


def _buscar_etiqueta(raw: pd.DataFrame, *etiquetas: str, max_filas: int = 25):
    limite = min(max_filas, raw.shape[0])
    objetivo = {e.strip().upper() for e in etiquetas}
    for i in range(limite):
        for j in range(raw.shape[1]):
            val = raw.iat[i, j]
            if isinstance(val, str) and val.strip().upper() in objetivo:
                return i, j
    return None


def _valor_asociado(raw: pd.DataFrame, fila: int, col: int):
    if fila + 1 < raw.shape[0]:
        val = raw.iat[fila + 1, col]
        if pd.notna(val) and not _parece_label(val):
            return val

    for c in range(col + 1, raw.shape[1]):
        val = raw.iat[fila, c]
        if pd.notna(val) and not _parece_label(val):
            return val

    return None


def _inferir_estacionalidad_de_nombre(nombre_archivo: str) -> str:
    nombre = nombre_archivo.upper()
    for clave in _ESTACIONALIDADES_CONOCIDAS:
        if clave in nombre:
            return clave
    return "DESCONOCIDA"


def leer_vigencia_a5(a5_path: Path, hoja: str = "LPP") -> Tuple[str, datetime.date, datetime.date]:
    """Lee estacionalidad, fecha_inicio y fecha_fin de vigencia del archivo A5."""
    a5_path = Path(a5_path)
    raw = pd.read_excel(a5_path, sheet_name=hoja, header=None)

    pos_inicio = _buscar_etiqueta(raw, "FECHA INICIO A5", "FECHA INICIO")
    pos_fin = _buscar_etiqueta(raw, "FECHA FIN A5", "FECHA FIN")
    if pos_inicio is None or pos_fin is None:
        raise ValueError(
            f"No se encontraron las etiquetas de vigencia (FECHA INICIO/FIN) en {a5_path.name}"
        )

    val_ini = _valor_asociado(raw, *pos_inicio)
    val_fin = _valor_asociado(raw, *pos_fin)
    if val_ini is None or val_fin is None:
        raise ValueError(
            f"No se pudieron leer los valores de vigencia en {a5_path.name}"
        )

    fecha_inicio = pd.to_datetime(val_ini, dayfirst=True).date()
    fecha_fin = pd.to_datetime(val_fin, dayfirst=True).date()

    pos_est = _buscar_etiqueta(raw, "Estacionalidad")
    if pos_est is not None:
        valor_est = _valor_asociado(raw, *pos_est)
        estacionalidad = str(valor_est).strip() if valor_est is not None else None
    else:
        estacionalidad = None

    if not estacionalidad:
        estacionalidad = _inferir_estacionalidad_de_nombre(a5_path.name)

    return estacionalidad, fecha_inicio, fecha_fin


def buscar_a5_para_fecha(
    empresa_dir: Path, fecha: datetime.date, estacionalidades_validas: Tuple[str, ...] = ESTACIONALIDADES_VALIDAS_IP
) -> Path:
    """Busca el archivo A5 oficial vigente para la fecha y empresa especificadas."""
    empresa_dir = Path(empresa_dir)
    fecha_dt = pd.Timestamp(fecha).date()
    validas = {e.upper() for e in estacionalidades_validas}

    archivos_a5 = sorted(empresa_dir.glob("*.xlsx"))
    candidatos_validos = []

    for a5_path in archivos_a5:
        # Ignorar archivos que sean A1 o de salida
        if "_A1_" in a5_path.name.upper() or a5_path.name.startswith("reporte_"):
            continue
        try:
            est, f_ini, f_fin = leer_vigencia_a5(a5_path)
        except Exception:
            continue

        if est.upper() in validas and f_ini <= fecha_dt <= f_fin:
            candidatos_validos.append((a5_path, f_ini, f_fin))

    if candidatos_validos:
        # Retorna el más específico
        return candidatos_validos[0][0]

    # Fallback: buscar archivos con 'A5' en el nombre
    for a5_path in archivos_a5:
        if "A5" in a5_path.name.upper():
            return a5_path

    raise FileNotFoundError(
        f"No se encontró un archivo A5 vigente para la fecha {fecha_dt} en {empresa_dir}"
    )


def cargar_lpp(a5_path: Path, hoja: str = "LPP", skiprows: int = 10) -> pd.DataFrame:
    """Carga y procesa la hoja LPP del archivo A5."""
    df_a5 = pd.read_excel(a5_path, sheet_name=hoja, skiprows=skiprows, header=0)
    df_a5 = df_a5.rename(columns=CAMPOS_A5)

    for col in ["IPP_anterior", "TPP", "IPP_posterior"]:
        if col in df_a5.columns:
            df_a5[col + "_seg"] = df_a5[col].apply(a_segundos)

    return df_a5


# --- CARGA Y LIMPIEZA DE EXPEDICIONES (LPO) ---

def _limpiar_expediciones(df_expediciones: pd.DataFrame) -> pd.DataFrame:
    df = df_expediciones.copy()

    renombres = {
        "Variante": "Servicio",
        "Dirección": "Sentido",
        "Direccion": "Sentido",
        "Fecha": "Inicio Expedicion",
        "Inicio Expedición": "Inicio Expedicion",
        "Período": "Periodo",
    }
    for col_orig, col_dest in renombres.items():
        if col_orig in df.columns and col_dest not in df.columns:
            df = df.rename(columns={col_orig: col_dest})

    for col_req in ["Servicio", "Sentido", "Inicio Expedicion", "Estado"]:
        if col_req not in df.columns:
            raise KeyError(f"No se encontró la columna requerida '{col_req}' en las expediciones.")

    if "Periodo" not in df.columns:
        df["Periodo"] = None
    if "Bus" not in df.columns:
        df["Bus"] = None

    df["Sentido"] = df["Sentido"].replace(MAPEO_SENTIDO)

    df["Estado_clean"] = (
        df["Estado"]
        .astype(str)
        .str.strip()
        .str.lower()
        .str.replace("á", "a")
        .str.replace("é", "e")
        .str.replace("í", "i")
        .str.replace("ó", "o")
        .str.replace("ú", "u")
    )
    df = df[df["Estado_clean"].str.startswith("val")].copy()
    df["Estado"] = "valida"

    df["fecha"] = pd.to_datetime(df["Inicio Expedicion"]).dt.date
    df["tipo_dia"] = df["fecha"].apply(clasificar_tipo_dia)

    return df


def cargar_lpo(expediciones_path: Path, formato: str = "auto") -> pd.DataFrame:
    """Carga el archivo de expediciones observadas."""
    expediciones_path = Path(expediciones_path)
    ext = expediciones_path.suffix.lower()

    if formato == "auto":
        if ext == ".xlsx":
            df_expediciones = pd.read_excel(expediciones_path)
        elif ext == ".csv":
            df_expediciones = pd.read_csv(expediciones_path)
        else:
            try:
                df_expediciones = pd.read_html(expediciones_path)[0]
            except Exception:
                df_expediciones = pd.read_excel(expediciones_path)
    elif formato == "html":
        df_expediciones = pd.read_html(expediciones_path)[0]
    elif formato in ("excel", "xlsx"):
        df_expediciones = pd.read_excel(expediciones_path)
    elif formato == "csv":
        df_expediciones = pd.read_csv(expediciones_path)
    else:
        raise ValueError(f"Formato de expediciones no soportado: {formato}")

    return _limpiar_expediciones(df_expediciones)


def construir_lpo_largo(df_expediciones: pd.DataFrame) -> pd.DataFrame:
    """Convierte las columnas de horas por punto de control en formato largo."""
    columnas_pc = [c for c in df_expediciones.columns if str(c).strip().isdigit()]

    if not columnas_pc:
        return pd.DataFrame(
            columns=[
                "Servicio",
                "Sentido",
                "fecha",
                "Inicio Expedicion",
                "Estado",
                "Bus",
                "Periodo",
                "correlativo_pc",
                "TPO",
                "TPO_seg",
            ]
        )

    df_lpo = df_expediciones.melt(
        id_vars=["Servicio", "Sentido", "fecha", "Inicio Expedicion", "Estado", "Bus", "Periodo"],
        value_vars=columnas_pc,
        var_name="correlativo_pc",
        value_name="TPO",
    )

    df_lpo["correlativo_pc"] = df_lpo["correlativo_pc"].astype(int)
    df_lpo["TPO_seg"] = df_lpo["TPO"].apply(a_segundos)
    df_lpo = df_lpo[(df_lpo["Estado"] == "valida") & df_lpo["TPO_seg"].notna()]

    return df_lpo


def buscar_archivo_expediciones(empresa_dir: Path, mes_num: int, anio_num: int) -> Optional[Path]:
    """Busca el archivo de expediciones para un mes y año determinado."""
    empresa_dir = Path(empresa_dir)
    anio_2d = anio_num % 100
    mes_nombre = None
    for k, v in MESES_MAP.items():
        if v[0] == mes_num:
            mes_nombre = v[2]
            break

    patrones_carpeta = [
        f"{mes_nombre}{anio_2d}",
        f"{mes_nombre}{anio_num}",
        f"{mes_nombre}_{anio_2d}",
        f"{mes_nombre}_{anio_num}",
        f"{mes_num:02d}{anio_2d}",
        f"{mes_num:02d}{anio_num}",
    ]

    for carpeta in empresa_dir.iterdir():
        if not carpeta.is_dir():
            continue
        c_name = carpeta.name.strip()
        # Coincidencia directa o parcial
        for pat in patrones_carpeta:
            if pat.lower() in c_name.lower():
                candidatos = sorted(carpeta.glob("*.xls")) + sorted(carpeta.glob("*.xlsx")) + sorted(carpeta.glob("*.csv"))
                if candidatos:
                    return candidatos[0]

    # Búsqueda directa en la raíz de empresa_dir
    candidatos_raiz = sorted(empresa_dir.glob(f"*{mes_nombre.lower()}*.xlsx")) if mes_nombre else []
    if candidatos_raiz:
        return candidatos_raiz[0]

    return None


# --- MOTOR DE CÁLCULO DE IP ---

def calcular_ip_un_tpp(
    tpp: float,
    ipp_ant: float,
    ipp_post: float,
    lpo_disponible: List[dict],
) -> Tuple[float, Optional[dict]]:
    """Calcula el IP de un TPP según las bandas de tolerancia reglamentarias."""
    niveles = [
        (tpp - ipp_ant / 12, tpp + ipp_post / 6, 1.0),
        (tpp - ipp_ant / 6, tpp - ipp_ant / 12, 0.75),
        (tpp + ipp_post / 6, tpp + ipp_post / 3, 0.75),
        (tpp - ipp_ant / 4, tpp - ipp_ant / 6, 0.50),
        (tpp + ipp_post / 3, tpp + ipp_post / 2, 0.50),
        (tpp - ipp_ant / 3, tpp - ipp_ant / 4, 0.25),
        (tpp + ipp_post / 2, tpp + (2 / 3) * ipp_post, 0.25),
    ]

    for lo, hi, valor in niveles:
        for bus in lpo_disponible:
            tpo = bus["TPO_seg"]
            if lo <= tpo <= hi:
                return valor, bus

    return 0.0, None


def calcular_resultados(df_a5: pd.DataFrame, df_lpo: pd.DataFrame, fechas: List) -> pd.DataFrame:
    """Cruza todas las TPP del A5 contra las LPO observadas."""
    servicios_sentido_a5 = list(
        df_a5[["Servicio", "Sentido"]].drop_duplicates().itertuples(index=False, name=None)
    )

    resultados = []

    for fecha in sorted(fechas):
        tipo_dia_actual = clasificar_tipo_dia(fecha)
        lpo_del_dia = df_lpo[df_lpo["fecha"] == fecha]

        for servicio, sentido in servicios_sentido_a5:
            lpp_grupo = df_a5[
                (df_a5.Servicio == servicio)
                & (df_a5.Sentido == sentido)
                & (df_a5.tipo_dia == tipo_dia_actual)
            ].sort_values("correlativo_pc")

            if lpp_grupo.empty:
                continue

            for pc in lpp_grupo["correlativo_pc"].unique():
                lpp_pc = lpp_grupo[lpp_grupo.correlativo_pc == pc].sort_values("TPP_seg")

                lpo_pc = lpo_del_dia[
                    (lpo_del_dia.Servicio == servicio)
                    & (lpo_del_dia.Sentido == sentido)
                    & (lpo_del_dia.correlativo_pc == pc)
                ][["TPO_seg", "Periodo"]].to_dict("records")
                lpo_pc.sort(key=lambda x: x["TPO_seg"])

                for _, fila in lpp_pc.iterrows():
                    ip_valor, bus_usado = calcular_ip_un_tpp(
                        fila["TPP_seg"], fila["IPP_anterior_seg"], fila["IPP_posterior_seg"], lpo_pc
                    )

                    periodo_usado = None
                    if bus_usado is not None:
                        periodo_usado = bus_usado["Periodo"]
                        lpo_pc.remove(bus_usado)

                    resultados.append(
                        {
                            "fecha": fecha,
                            "tipo_dia": tipo_dia_actual,
                            "Servicio": servicio,
                            "Sentido": sentido,
                            "correlativo_pc": pc,
                            "TPP": fila["TPP"],
                            "Periodo": periodo_usado,
                            "IP": ip_valor,
                        }
                    )

    return pd.DataFrame(resultados)


def _aplicar_topes_ip(ip_val: float) -> float:
    """Aplica topes normativos: < 0.50 -> 0.50, > 0.90 -> 1.00."""
    if pd.isna(ip_val):
        return np.nan
    ip_red = round(ip_val, 2)
    if ip_red < 0.50:
        return 0.50
    elif ip_red > 0.90:
        return 1.00
    return ip_red


def calcular_ip_mensual(df_resultados: pd.DataFrame) -> Tuple[float, float]:
    """Calcula IP promedio (IP_M') e IP final (IP_M) con topes."""
    if df_resultados.empty:
        return 0.0, 0.50
    ip_promedio = round(float(df_resultados["IP"].mean()), 2)
    ip_final = _aplicar_topes_ip(ip_promedio)
    return ip_promedio, ip_final


def construir_resumen_por_servicio(df_resultados: pd.DataFrame) -> pd.DataFrame:
    """Construye la tabla resumen agregada por Servicio."""
    filas = []
    servicios = sorted(df_resultados["Servicio"].dropna().unique())

    for s in servicios:
        sub = df_resultados[df_resultados["Servicio"] == s]
        total_tpp = len(sub)
        p_100 = int((sub["IP"] == 1.0).sum())
        p_075 = int((sub["IP"] == 0.75).sum())
        p_050 = int((sub["IP"] == 0.50).sum())
        p_025 = int((sub["IP"] == 0.25).sum())
        p_000 = int((sub["IP"] == 0.0).sum())
        pct_cumple = round(((p_100 + p_075) / total_tpp) * 100, 2) if total_tpp > 0 else 0.0
        ip_prom = round(float(sub["IP"].mean()), 4) if total_tpp > 0 else 0.0
        ip_final = _aplicar_topes_ip(ip_prom)

        filas.append(
            {
                "Servicio": s,
                "Total TPP": total_tpp,
                "Puntaje 1.0": p_100,
                "Puntaje 0.75": p_075,
                "Puntaje 0.50": p_050,
                "Puntaje 0.25": p_025,
                "Puntaje 0.0": p_000,
                "% Cumplimiento (>=0.75)": pct_cumple,
                "IP Promedio (IP')": round(ip_prom, 2),
                "IP Final (IP_M)": ip_final,
            }
        )

    total_gral = len(df_resultados)
    if total_gral > 0:
        tot_100 = int((df_resultados["IP"] == 1.0).sum())
        tot_075 = int((df_resultados["IP"] == 0.75).sum())
        tot_050 = int((df_resultados["IP"] == 0.50).sum())
        tot_025 = int((df_resultados["IP"] == 0.25).sum())
        tot_000 = int((df_resultados["IP"] == 0.0).sum())
        pct_tot_cumple = round(((tot_100 + tot_075) / total_gral) * 100, 2)
        ip_prom_gral = round(float(df_resultados["IP"].mean()), 4)
        ip_final_gral = _aplicar_topes_ip(ip_prom_gral)
    else:
        tot_100 = tot_075 = tot_050 = tot_025 = tot_000 = 0
        pct_tot_cumple = ip_prom_gral = ip_final_gral = 0.0

    filas.append(
        {
            "Servicio": "TOTAL / PROMEDIO GENERAL",
            "Total TPP": total_gral,
            "Puntaje 1.0": tot_100,
            "Puntaje 0.75": tot_075,
            "Puntaje 0.50": tot_050,
            "Puntaje 0.25": tot_025,
            "Puntaje 0.0": tot_000,
            "% Cumplimiento (>=0.75)": pct_tot_cumple,
            "IP Promedio (IP')": round(ip_prom_gral, 2),
            "IP Final (IP_M)": ip_final_gral,
        }
    )

    return pd.DataFrame(filas)


# --- EXPORTACIÓN DE EXCEL Y RESUMEN ---

def generar_excel_bytes_ip(df_resultados: pd.DataFrame) -> bytes:
    """Genera el contenido binario del libro Excel de IP con estilos."""
    df_resumen = construir_resumen_por_servicio(df_resultados)
    ip_promedio = round(float(df_resultados["IP"].mean()), 4) if not df_resultados.empty else 0.0

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resumen.to_excel(writer, sheet_name="Resumen por Servicio", index=False)
        ws_resumen = writer.sheets["Resumen por Servicio"]

        df_resultados.to_excel(writer, sheet_name="Reporte", index=False)
        ws_reporte = writer.sheets["Reporte"]

        # Estilos
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        center_align = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )
        total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        total_font = Font(name="Calibri", size=11, bold=True, color="000000")
        double_bottom_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="double", color="000000"),
        )

        # Formatear Hoja 'Resumen por Servicio'
        for col_num in range(1, len(df_resumen.columns) + 1):
            cell = ws_resumen.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        n_filas_resumen = len(df_resumen)
        for row_num in range(2, n_filas_resumen + 2):
            is_total_row = row_num == n_filas_resumen + 1
            for col_num in range(1, len(df_resumen.columns) + 1):
                cell = ws_resumen.cell(row=row_num, column=col_num)
                cell.alignment = center_align
                if is_total_row:
                    cell.fill = total_fill
                    cell.font = total_font
                    cell.border = double_bottom_border
                else:
                    cell.border = thin_border

        for col in ws_resumen.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_resumen.column_dimensions[col_letter].width = max(max_len + 4, 12)

        # Formatear Hoja 'Reporte'
        for col_num in range(1, len(df_resultados.columns) + 1):
            cell = ws_reporte.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        fila_salto = len(df_resultados) + 3
        ws_reporte.cell(row=fila_salto, column=1, value="Promedio IP:")
        ws_reporte.cell(row=fila_salto, column=2, value=round(ip_promedio, 2))
        ws_reporte.cell(row=fila_salto, column=1).font = Font(bold=True)
        ws_reporte.cell(row=fila_salto, column=2).font = Font(bold=True)

        for col in ws_reporte.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_reporte.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output.seek(0)
    return output.getvalue()


def generar_resumen_whatsapp(
    empresa: str,
    mes_nombre: str,
    anio: int,
    a5_nombre: str,
    df_resultados: pd.DataFrame,
    df_resumen: pd.DataFrame,
) -> str:
    """Genera el mensaje resumen formateado para WhatsApp."""
    total_tpp = len(df_resultados)
    ip_promedio, ip_final = calcular_ip_mensual(df_resultados)

    tot_100 = int((df_resultados["IP"] == 1.0).sum())
    tot_075 = int((df_resultados["IP"] == 0.75).sum())
    pct_cumple = round(((tot_100 + tot_075) / total_tpp) * 100, 1) if total_tpp > 0 else 0.0

    msg = (
        f"📊 *Reporte del Indicador de Puntualidad (IP)*\n\n"
        f"🏢 *Operador*: {empresa.upper()}\n"
        f"📅 *Período*: {mes_nombre} {anio}\n"
        f"📑 *A5 Vigente*: `{a5_nombre}`\n"
        f"🚌 *Total Pasadas Programadas (TPP)*: {total_tpp:,}\n"
        f"🎯 *Cumplimiento (≥ 0.75)*: {pct_cumple}%\n\n"
        f"📈 *IP Promedio (IP'_M)*: *{ip_promedio:.2f}*\n"
        f"🏆 *IP Final (IP_M)*: *{ip_final:.2f}*\n\n"
    )

    # Detalle por servicio
    servicios_rows = df_resumen[df_resumen["Servicio"] != "TOTAL / PROMEDIO GENERAL"]
    if not servicios_rows.empty:
        msg += "📋 *Desglose por Servicio*:\n"
        for _, r in servicios_rows.iterrows():
            srv = r["Servicio"]
            tpp_s = r["Total TPP"]
            cumple_s = r["% Cumplimiento (>=0.75)"]
            ip_fin_s = r["IP Final (IP_M)"]
            msg += f"• *{srv}*: {cumple_s:.1f}% cumple | IP={ip_fin_s:.2f} ({tpp_s:,} TPP)\n"
        msg += "\n"

    msg += "📎 _Adjunto reporte detallado en Excel con hojas 'Resumen por Servicio' y 'Reporte'._"
    return msg


# --- ORQUESTADOR PRINCIPAL DEL CÁLCULO ---

def ejecutar_calculo_ip(
    empresa: str,
    anio: int,
    mes: int,
    data_dir: Optional[Union[str, Path]] = None,
    salida_dir: Optional[Union[str, Path]] = None,
) -> Tuple[bytes, str, str]:
    """
    Ejecuta el cálculo completo del Indicador de Puntualidad (IP) para una empresa, año y mes.

    Retorna:
    --------
    (excel_bytes, resumen_txt, filename)
    """
    empresa_clean = empresa.lower().strip()
    if "tasa" in empresa_clean:
        empresa_clean = "tasacop"

    if data_dir is None:
        data_dir = Path("data")
    else:
        data_dir = Path(data_dir)

    empresa_dir = data_dir / empresa_clean
    if not empresa_dir.exists() or not empresa_dir.is_dir():
        raise FileNotFoundError(f"No se encontró el directorio de datos para la empresa '{empresa_clean}' en {data_dir}")

    mes_nombre = None
    mes_abbr = f"{mes:02d}"
    for k, v in MESES_MAP.items():
        if v[0] == mes:
            mes_abbr = v[1]
            mes_nombre = v[2]
            break

    if mes_nombre is None:
        mes_nombre = f"Mes{mes:02d}"

    # 1. Buscar A5 vigente para la fecha consultada
    anchor_fecha = datetime.date(anio, mes, 1)
    a5_path = buscar_a5_para_fecha(
        empresa_dir, anchor_fecha, estacionalidades_validas=ESTACIONALIDADES_VALIDAS_IP
    )

    # 2. Buscar archivo de expediciones
    expediciones_path = buscar_archivo_expediciones(empresa_dir, mes_num=mes, anio_num=anio)
    if not expediciones_path:
        raise FileNotFoundError(
            f"No se encontró archivo de expediciones para {empresa_clean.upper()} en {mes_nombre} {anio} dentro de {empresa_dir}"
        )

    # 3. Cargar datos
    df_a5 = cargar_lpp(a5_path)
    df_expediciones = cargar_lpo(expediciones_path, formato="auto")
    df_lpo = construir_lpo_largo(df_expediciones)

    # 4. Calcular resultados
    fechas = df_expediciones["fecha"].unique()
    df_resultados = calcular_resultados(df_a5, df_lpo, fechas)

    if df_resultados.empty:
        raise ValueError(
            f"No se generaron resultados para {empresa_clean.upper()} en {mes_nombre} {anio}. "
            f"Verifica el cruce de Servicios, Sentidos y Puntos de Control entre el A5 y las expediciones."
        )

    # 5. Construir resumen y generar Excel
    df_resumen = construir_resumen_por_servicio(df_resultados)
    excel_bytes = generar_excel_bytes_ip(df_resultados)

    # 6. Guardar copia en disco en carpeta salidas
    if salida_dir is None:
        salida_dir = Path("salidas")
    else:
        salida_dir = Path(salida_dir)
    salida_dir.mkdir(parents=True, exist_ok=True)

    anio_2d = anio % 100
    filename = f"reporte_IP_{empresa_clean.upper()}_{mes_nombre}{anio_2d}.xlsx"
    out_path = salida_dir / filename
    with open(out_path, "wb") as f:
        f.write(excel_bytes)

    resumen_txt = generar_resumen_whatsapp(
        empresa=empresa_clean,
        mes_nombre=mes_nombre,
        anio=anio,
        a5_nombre=a5_path.name,
        df_resultados=df_resultados,
        df_resumen=df_resumen,
    )

    return excel_bytes, resumen_txt, filename
