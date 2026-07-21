import requests
import pandas as pd
import io
import os
import logging
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import sqlite3
from sqlmodel import Session
from sqlalchemy import text

# Cargar variables de entorno
load_dotenv()

# --- CONFIGURACIÓN ---
DIAS_RECONCILIACION = int(os.getenv("DIAS_RECONCILIACION", "7"))

EMPRESAS = {
    "lider": {
        "metodo": "requests",
        "base_url": "https://lider.transidea.cl/sitio",
        "usuario_env": "user",
        "password_env": "pass",
    },
    "toptur": {
        "metodo": "requests",
        "base_url": "https://toptur.transidea.cl/sitio",
        "usuario_env": "user",
        "password_env": "pass",
    },
    "tasacop": {  # Normalizado a 'tasacop' con una 'o' como en bot_nabla
        "metodo": "playwright",
        "base_url": "https://saef.citymovil.cl",
        "usuario_env": "TASACOOP_USER",
        "password_env": "TASACOOP_PASS",
    },
}

logging.basicConfig(
    filename="log_actualizacion.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# --- UTILS ---

def limpiar_nan(val):
    if pd.isna(val) or val is None or str(val).strip().lower() in ("nan", "nat", ""):
        return None
    return val

# --- MÉTODO 1: Requests (Lider, Toptur) ---

def login_requests(session, base_url, usuario, password):
    session.get(f"{base_url}/login/")
    resp = session.post(
        f"{base_url}/login/ajax/valida.php",
        headers={"x-requested-with": "XMLHttpRequest"},
        data={"modal": "true", "usuario_login": usuario, "pass_login": password},
    )
    data = resp.json()
    if not data.get("ok"):
        raise Exception(f"Login falló: {data}")
    return session


def descargar_expediciones_requests(session, base_url, fecha_inicio, fecha_fin):
    visor_url = f"{base_url}/sistemas/msgcr/modulos/reporteria_star/expediciones_visor.php"
    controller_url = f"{base_url}/sistemas/msgcr/modulos/reporteria_star/ajax/expediciones_controller.php"

    session.get(visor_url)
    body = {
        "rangofecha": f"{fecha_inicio} 00:00:00 - {fecha_fin} 23:59:59",
        "estado": 0,
        "motivos[]": [1, 3, 5, 4, 2, 8, 6, 7, 9],
        "maquina": 0,
        "servicio": 0,
        "recorrido": 0,
        "porcentaje": 0,
        "operacional": "",
        "periodo": "",
        "star_bd": "",
    }
    resp = session.post(
        f"{controller_url}?tipo_peticion=data_excel",
        headers={"x-requested-with": "XMLHttpRequest"},
        data=body,
    )
    resp.raise_for_status()
    return resp.text


def parsear_html_a_dataframe(html_text):
    tablas = pd.read_html(io.StringIO(html_text))
    df = tablas[0]
    df.columns = [str(c).strip() for c in df.columns]
    return df


def obtener_datos_requests(config, fecha_inicio, fecha_fin):
    session = requests.Session()
    login_requests(
        session,
        config["base_url"],
        os.getenv(config["usuario_env"]),
        os.getenv(config["password_env"]),
    )
    html = descargar_expediciones_requests(session, config["base_url"], fecha_inicio, fecha_fin)
    return parsear_html_a_dataframe(html)

# --- MÉTODO 2: Playwright (Tasacop) ---

def set_datepicker_value(page, selector, fecha):
    page.evaluate(f"""
        () => {{
            const $el = $('{selector}');
            $el.datepicker('setDate', '{fecha}');
            const inst = $el.data('datepicker');
            if (inst && inst.settings && typeof inst.settings.onSelect === 'function') {{
                const dateStr = $el.val();
                inst.settings.onSelect.call($el[0], dateStr, inst);
            }}
        }}
    """)


def obtener_datos_playwright(config, fecha_inicio, fecha_fin):
    usuario = os.getenv(config["usuario_env"])
    password = os.getenv(config["password_env"])
    base_url = config["base_url"]

    ruta_temporal = Path(f"_tmp_descarga_tasacop.xlsx")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()

            page.goto(f"{base_url}/SaefWeb/index.zul")
            page.fill("input[name='username']", usuario)
            page.fill("input[name='password']", password)
            page.click("#loginBtn")
            page.wait_for_load_state("networkidle")

            page.goto(f"{base_url}/SaefWeb/component/indicador/expeditions.zul")
            page.wait_for_load_state("networkidle")

            set_datepicker_value(page, "#myPicker", fecha_inicio)
            page.wait_for_timeout(500)
            set_datepicker_value(page, "#myPickerTo", fecha_fin)
            page.wait_for_timeout(500)

            page.click("#btnAccept")
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)

            with page.expect_download(timeout=30000) as download_info:
                page.click("#btnExport")
            download = download_info.value
            download.save_as(ruta_temporal)

            browser.close()

        df = pd.read_excel(ruta_temporal)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    finally:
        if ruta_temporal.exists():
            try:
                ruta_temporal.unlink()
            except Exception as e:
                print(f"⚠️ No se pudo eliminar el archivo temporal: {e}", flush=True)

# --- INTEGRACIÓN CON LA BASE DE DATOS SQLITE ---

def guardar_expediciones_en_db(df: pd.DataFrame, operador: str, fecha_inicio: str, fecha_fin: str):
    """
    Guarda las expediciones del DataFrame en la base de datos SQLite.
    Elimina los registros existentes para el operador y rango de fechas antes de insertar.
    """
    from database import engine, Expedicion
    
    # Convertir fechas de dd-mm-yyyy a yyyy-mm-dd
    f_ini_dt = datetime.strptime(fecha_inicio, "%d-%m-%Y")
    f_fin_dt = datetime.strptime(fecha_fin, "%d-%m-%Y")
    f_ini_str = f_ini_dt.strftime("%Y-%m-%d")
    f_fin_str = f_fin_dt.strftime("%Y-%m-%d")
    
    # 1. Limpiar registros existentes para este operador y rango de fechas
    with Session(engine) as session:
        session.execute(
            text("DELETE FROM expediciones WHERE operador = :op AND Fecha BETWEEN :ini AND :fin"),
            {"op": operador, "ini": f_ini_str, "fin": f_fin_str}
        )
        session.commit()
    print(f"🧹 Registros anteriores de '{operador}' entre {f_ini_str} y {f_fin_str} eliminados de la DB.")
    
    # 2. Mapear filas a objetos Expedicion
    records = []
    is_transidea = (operador in ("lider", "toptur"))
    
    for idx, row in df.iterrows():
        try:
            fecha_val = str(row.get('Fecha', ''))
            try:
                if "-" in fecha_val:
                    partes = fecha_val.split("-")
                    if len(partes[0]) == 4:
                        fecha_db = fecha_val
                    else:
                        fecha_db = datetime.strptime(fecha_val, "%d-%m-%Y").strftime("%Y-%m-%d")
                elif "/" in fecha_val:
                    partes = fecha_val.split("/")
                    if len(partes[0]) == 4:
                        fecha_db = fecha_val
                    else:
                        fecha_db = datetime.strptime(fecha_val, "%d/%m/%Y").strftime("%Y-%m-%d")
                else:
                    fecha_db = fecha_val
            except Exception:
                fecha_db = fecha_val

            if is_transidea:
                chofer = limpiar_nan(row.get('Chofer'))
                sentido = str(row.get('Sentido', ''))
                periodo = int(row.get('Periodo', 0)) if not pd.isna(row.get('Periodo')) else 0
                tipo_demanda = limpiar_nan(row.get('Tipo demanda'))
                frecuencia = float(row.get('Frecuencia')) if not pd.isna(row.get('Frecuencia')) else None
                inicio_exp = str(row.get('Inicio Expedicion', ''))
                fin_exp = str(row.get('Fin Expedicion', ''))
                servicio = str(row.get('Servicio', ''))
                propietario = limpiar_nan(row.get('Propietario'))
                
                pois = {}
                for p_num in range(1, 23):
                    pois[f"p{p_num}"] = limpiar_nan(row.get(str(p_num)))
            else:
                chofer = limpiar_nan(row.get('Conductor'))
                sentido = str(row.get('Dirección', ''))
                periodo = int(row.get('Período', 0)) if not pd.isna(row.get('Período')) else 0
                tipo_demanda = limpiar_nan(row.get('Tipo Demanda'))
                frecuencia = float(row.get('Frecuencia Exigida')) if not pd.isna(row.get('Frecuencia Exigida')) else None
                inicio_exp = ""
                fin_exp = ""
                servicio = str(row.get('Variante', '')).strip()
                if operador.lower() == "tasacop":
                    import re
                    servicio = re.sub(r"_(I|R)$", "", servicio)
                propietario = None
                
                pois = {}
                for p_num in range(1, 23):
                    pois[f"p{p_num}"] = limpiar_nan(row.get(f"{p_num:02d}"))

            exp_data = {
                'operador': operador,
                'Fecha': fecha_db,
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
            exp_data.update(pois)
            
            record = Expedicion(**exp_data)
            records.append(record)
            
        except Exception as e:
            print(f"⚠️ Error procesando fila {idx} de {operador}: {e}")
            
    # 3. Insertar nuevos registros en lotes
    if records:
        batch_size = 1000
        total_records = len(records)
        with Session(engine) as session:
            for i in range(0, total_records, batch_size):
                batch = records[i:i+batch_size]
                session.add_all(batch)
                session.commit()
        print(f"✅ Se importaron {total_records} nuevas expediciones de '{operador}' en SQLite.")
        logging.info(f"[{operador}] OK - Insertados: {total_records}")
        return total_records
    else:
        print(f"⚠️ No se encontraron expediciones para guardar para '{operador}'.")
        logging.info(f"[{operador}] Sin registros nuevos.")
        return 0

# --- ORQUESTADOR ---

def procesar_empresa(nombre, config) -> int:
    try:
        hoy = datetime.now()
        if nombre == "tasacop":
            # 1 día de desfase: descargar solo el día anterior (fecha_inicio = fecha_fin = ayer)
            ayer = hoy - timedelta(days=1)
            fecha_inicio = ayer.strftime("%d-%m-%Y")
            fecha_fin = ayer.strftime("%d-%m-%Y")
        else:
            # Transidea: 5 días de reconciliación
            fecha_inicio = (hoy - timedelta(days=5)).strftime("%d-%m-%Y")
            fecha_fin = hoy.strftime("%d-%m-%Y")

        print(f"\n🔄 Iniciando descarga para '{nombre}' ({fecha_inicio} al {fecha_fin})...")
        if config["metodo"] == "requests":
            df_nuevo = obtener_datos_requests(config, fecha_inicio, fecha_fin)
        elif config["metodo"] == "playwright":
            df_nuevo = obtener_datos_playwright(config, fecha_inicio, fecha_fin)
        else:
            raise Exception(f"Método desconocido: {config['metodo']}")

        print(f"📥 Descargados {len(df_nuevo)} registros para '{nombre}'.")
        
        # Guardar en SQLite y retornar cantidad
        total = guardar_expediciones_en_db(df_nuevo, nombre, fecha_inicio, fecha_fin)
        return total

    except Exception as e:
        logging.error(f"[{nombre}] FALLÓ: {e}")
        print(f"❌ [{nombre}] Error: {e}")
        raise e


def descargar_todos() -> dict:
    from database import init_db
    init_db()
    
    resumen = {}
    errores = []
    
    for nombre, config in EMPRESAS.items():
        try:
            total_insertado = procesar_empresa(nombre, config)
            resumen[nombre] = total_insertado
        except Exception as e:
            resumen[nombre] = 0
            errores.append(f"{nombre}: {str(e)}")
            
    return {
        "success": len(errores) == 0,
        "resumen": resumen,
        "errores": errores
    }


if __name__ == "__main__":
    resultado = descargar_todos()
    print("\n🏁 Proceso de descarga finalizado:", resultado)

