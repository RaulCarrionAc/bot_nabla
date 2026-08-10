from fastapi import FastAPI, Request, BackgroundTasks
from typing import Optional, List, Dict, Union, Tuple
import uvicorn
import base64
import os
import requests as http_requests
from dotenv import load_dotenv
from database import (
    init_db, obtener_estado, guardar_estado, eliminar_estado, DATABASE_URL,
    es_usuario_permitido, agregar_usuario_permitido, quitar_usuario_permitido, listar_usuarios_permitidos
)
from handlers.puntualidad import procesar_puntualidad_desde_db
from handlers.icf import ejecutar_calculo_icf
from handlers.ip import ejecutar_calculo_ip
from datetime import datetime

load_dotenv()

ADMIN_CELLPHONES = [num.strip() for num in os.getenv("ADMIN_CELLPHONE", "56989158197").split(",") if num.strip()]

# Lista global para registrar los últimos intentos denegados
INTENTOS_DENEGADOS = []

def registrar_intento_denegado(sender: str):
    global INTENTOS_DENEGADOS
    ahora = datetime.now()
    # Evitar duplicados recientes
    INTENTOS_DENEGADOS = [x for x in INTENTOS_DENEGADOS if x["sender"] != sender]
    INTENTOS_DENEGADOS.append({"sender": sender, "time": ahora})
    # Mantener solo los últimos 5 intentos
    if len(INTENTOS_DENEGADOS) > 5:
        INTENTOS_DENEGADOS.pop(0)


def obtener_remitente_limpio(data: dict, session_id: str = None) -> str:
    """Extrae y resuelve el número de celular limpio del remitente (soportando LID)."""
    sender_id = data.get("author") or data.get("sender", {}).get("id") or data.get("from") or ""
    
    # 1. Si ya viene resuelto en el payload
    if data.get("senderPhone"):
        raw_phone = str(data.get("senderPhone")).split('@')[0]
        return "".join(filter(str.isdigit, raw_phone))
        
    # 2. Si es un LID y tenemos la sesión, intentar resolver a JID/Teléfono
    if str(sender_id).endswith("@lid") and session_id:
        # Método A: Consultar los detalles del contacto directamente
        try:
            r = http_requests.get(
                f"{OPENWA_URL}/api/sessions/{session_id}/contacts/{sender_id}",
                headers={"x-api-key": OPENWA_KEY}
            )
            if r.status_code in (200, 201):
                c_data = r.json()
                contact_obj = c_data.get("data") if isinstance(c_data, dict) and "data" in c_data else c_data
                if isinstance(contact_obj, dict):
                    # 'number' es el número telefónico real devuelto en whatsapp-web.js
                    phone = contact_obj.get("number") or contact_obj.get("phone")
                    if phone:
                        return "".join(filter(str.isdigit, str(phone).split('@')[0]))
        except Exception as e:
            print(f"⚠️ Error al resolver LID {sender_id} mediante /contacts: {e}", flush=True)

        # Método B: Fallback a /phone
        try:
            r = http_requests.get(
                f"{OPENWA_URL}/api/sessions/{session_id}/contacts/{sender_id}/phone",
                headers={"x-api-key": OPENWA_KEY}
            )
            if r.status_code in (200, 201):
                c_data = r.json()
                phone = c_data.get("phone") if isinstance(c_data, dict) else None
                if phone:
                    return "".join(filter(str.isdigit, str(phone).split('@')[0]))
        except Exception as e:
            print(f"⚠️ Error al resolver LID {sender_id} mediante /phone: {e}", flush=True)
            
    # 3. Fallback
    raw_num = str(sender_id).split('@')[0]
    return "".join(filter(str.isdigit, raw_num))

app = FastAPI()

@app.get("/health")
def health_check():
    return {"status": "ok"}


@app.get("/download/{filename}")
def download_archivo(filename: str):
    from fastapi.responses import FileResponse
    file_path = os.path.join("salidas", filename)
    if os.path.exists(file_path):
        return FileResponse(
            file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return {"error": f"Archivo '{filename}' no encontrado"}


OPENWA_URL = os.getenv("URL", "http://openwa-api:2785")
OPENWA_KEY = os.getenv("API-KEY")

def obtener_session_id_activo() -> str:
    """Consulta a la API de OpenWA para obtener el ID de la sesión 'nabla-bot' activa."""
    try:
        r = http_requests.get(
            f"{OPENWA_URL}/api/sessions",
            headers={"x-api-key": OPENWA_KEY},
            timeout=5
        )
        if r.status_code in (200, 201):
            sesiones = r.json()
            if isinstance(sesiones, list) and sesiones:
                for s in sesiones:
                    if s.get("name") == "nabla-bot" and s.get("status") in ("ready", "WORKING", "STARTING", "authenticated"):
                        return s.get("id") or s.get("name") or "nabla-bot"
                # Si hay alguna sesión activa
                return sesiones[0].get("id") or sesiones[0].get("name") or "nabla-bot"
    except Exception as e:
        print(f"⚠️ Error al obtener session_id activo: {e}", flush=True)
    return "nabla-bot"


def tarea_actualizar_datos(session_id: str, chat_id: str):
    from descargar_datos import descargar_todos
    try:
        res = descargar_todos()
        
        # Crear un resumen legible
        detalles = ""
        for op, cant in res.get("resumen", {}).items():
            detalles += f"- *{op.upper()}*: {cant} registros nuevos\n"
            
        if res.get("success"):
            msg = f"✅ *Descarga e Importación Diaria Completada*\n\n{detalles}"
        else:
            msg = f"⚠️ *Descarga Finalizada con Errores*\n\n{detalles}\n❌ *Errores detectados*:\n"
            for err in res.get("errores", []):
                msg += f"- {err}\n"
                
        enviar_mensaje(session_id, chat_id, msg)
    except Exception as e:
        enviar_mensaje(
            session_id, 
            chat_id, 
            f"❌ *Error crítico durante la actualización de datos*: {e}"
        )


def tarea_calcular_puntualidad(session_id: str, chat_id: str):
    try:
        db_path = DATABASE_URL.replace("sqlite:///", "")
        excel_bytes, resumen = procesar_puntualidad_desde_db(db_path)
        
        enviar_mensaje(session_id, chat_id, resumen)
        enviar_documento(session_id, chat_id, excel_bytes, "puntualidad_reporte.xlsx")
        print("✅ [DEBUG] Proceso completo desde base de datos", flush=True)
    except Exception as e:
        print(f"❌ [DEBUG] Error procesando desde base de datos: {e}", flush=True)
        enviar_mensaje(session_id, chat_id, "❌ Error al procesar los datos. Por favor verifica que las tablas 'operaciones' y 'anexo_5' en SQLite contengan registros válidos.")


def tarea_calcular_icf(session_id: str, chat_id: str, operador: str, anio: int, mes: int):
    try:
        db_path = DATABASE_URL.replace("sqlite:///", "")
        reporte_bytes, proy_bytes, resumen_txt = ejecutar_calculo_icf(operador, anio, mes, db_path)
        
        # Determinar si es mes pasado o vigente (mes actual o futuro)
        hoy = datetime.now()
        es_mes_pasado = (anio < hoy.year) or (anio == hoy.year and mes < hoy.month)
        
        enviar_mensaje(session_id, chat_id, resumen_txt)
        
        if es_mes_pasado:
            enviar_documento(session_id, chat_id, reporte_bytes, f"reporte_{operador}_{mes:02d}_{anio}.xlsx")
            print("✅ [DEBUG] Reporte ICF de mes pasado enviado con éxito", flush=True)
        else:
            enviar_documento(session_id, chat_id, proy_bytes, f"reporte_proyeccion_{operador}_{mes:02d}_{anio}.xlsx")
            print("✅ [DEBUG] Reporte Simulación ICF de mes vigente enviado con éxito", flush=True)
            
    except Exception as e:
        print(f"❌ [DEBUG] Error calculando ICF: {e}", flush=True)
        enviar_mensaje(session_id, chat_id, f"❌ Error calculando ICF: Asegúrate de que las frecuencias y expediciones estén cargadas para la fecha indicada.")


def tarea_calcular_ip(session_id: str, chat_id: str, empresa: str, anio: int, mes: int):
    try:
        session_id = session_id or obtener_session_id_activo()
        excel_bytes, resumen_txt, filename = ejecutar_calculo_ip(empresa, anio, mes)
        
        enviar_mensaje(session_id, chat_id, resumen_txt)
        enviar_documento(session_id, chat_id, excel_bytes, filename)
        print(f"✅ [DEBUG] Reporte IP ({empresa.upper()} {mes:02d}/{anio}) enviado con éxito", flush=True)
    except Exception as e:
        print(f"❌ [DEBUG] Error calculando IP ({empresa} {mes}/{anio}): {e}", flush=True)
        if session_id:
            enviar_mensaje(session_id, chat_id, f"❌ Error calculando Indicador de Puntualidad (IP):\n\n{e}")



def tarea_generar_reporte_tv(session_id: str, chat_id: str, mes_str: str, anio: int):
    try:
        from handlers.velocidades import (
            procesar_modelo_cinematico, 
            buscar_archivo_expediciones_tasacop, 
            MESES_MAP,
            DEFAULT_PO_A5_PATH
        )
        from generators.excel_speeds import generar_libro_excel_velocidades
        import pandas as pd

        session_id = session_id or obtener_session_id_activo()
        mes_info = MESES_MAP.get(mes_str.lower().strip())
        if not mes_info:
            if session_id:
                enviar_mensaje(session_id, chat_id, f"❌ Mes '{mes_str}' no reconocido. Ejemplos válidos: `mayo`, `junio`, `julio`.")
            return

        mes_num, mes_abbr, mes_nombre = mes_info
        print(f"⏳ [DEBUG] Iniciando generación de reporte cinemático para Tasacop: {mes_nombre} {anio} (session: {session_id})", flush=True)

        # 1. Intentar cargar expediciones directamente desde la base de datos SQLite
        from handlers.velocidades import obtener_expediciones_tasacop_mes_db
        df_in = obtener_expediciones_tasacop_mes_db(mes_num, anio)

        # 2. Si no está en SQLite, intentar fallback a archivo Excel en disco
        if df_in is None or df_in.empty:
            in_file = buscar_archivo_expediciones_tasacop(mes_str, anio)
            if in_file and os.path.exists(in_file):
                df_in = pd.read_excel(in_file)
            else:
                enviar_mensaje(
                    session_id, 
                    chat_id, 
                    f"⚠️ No se encontraron expediciones registradas para Tasacoop en *{mes_nombre} {anio}* (ni en SQLite ni en disco)."
                )
                return

        df_datos, df_desref, df_params, metrics = procesar_modelo_cinematico(
            df_in, 
            po_a5_path=None,
            operador="tasacop"
        )

        out_filename = f"reporte_velocidades_tasacop_{mes_abbr}_{anio}.xlsx"
        out_dir = os.path.join("salidas")
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, out_filename)

        _, raw_bytes = generar_libro_excel_velocidades(df_datos, df_desref, df_params, out_path)

        resumen_txt = (
            f"📊 *Reporte Cinemático de Velocidades y Tiempos de Viaje (PO A5)*\n\n"
            f"🏢 *Operador*: Tasacoop\n"
            f"📅 *Período*: {mes_nombre} {anio}\n"
            f"🚌 *Total Expediciones*: {metrics['total_expediciones']:,}\n"
            f"✅ *Expediciones en Muestra*: {metrics['total_muestra']:,} ({metrics['pct_muestra']:.1f}%)\n"
            f"⚡ *Velocidad Promedio Global*: {metrics['vel_promedio']:.2f} km/h\n\n"
            f"📎 _Adjunto libro Excel con Tablas Dinámicas y Tramos Descompuestos (desref)._"
        )

        enviar_mensaje(session_id, chat_id, resumen_txt)
        enviar_documento(session_id, chat_id, raw_bytes, out_filename)
        print(f"✅ [DEBUG] Reporte cinemático enviado exitosamente a {chat_id}", flush=True)

    except Exception as e:
        print(f"❌ [DEBUG] Error generando reporte de velocidades: {e}", flush=True)
        enviar_mensaje(session_id, chat_id, f"❌ Ocurrió un error al generar el reporte de velocidades: {e}")


def extraer_archivo_bytes(session_id: str, chat_id: str, message_id: str, data: dict) -> Optional[bytes]:
    """Extrae el contenido binario de un archivo adjunto enviado por WhatsApp (compatible con Baileys y whatsapp-web.js)."""
    # 1. Si viene en el objeto 'media' (estándar en OpenWA / Baileys para documentos)
    media_obj = data.get("media")
    if isinstance(media_obj, dict):
        # 1a. Base64 en media['data'] o media['base64']
        media_data = media_obj.get("data") or media_obj.get("base64")
        if media_data and isinstance(media_data, str) and len(media_data) > 50:
            try:
                if "base64," in media_data:
                    media_data = media_data.split("base64,")[1]
                return base64.b64decode(media_data)
            except Exception as e:
                print(f"⚠️ Error decodificando Base64 de media.data: {e}", flush=True)
                
        # 1b. URL en media['url'] o media['mediaUrl']
        m_url = media_obj.get("url") or media_obj.get("mediaUrl")
        if m_url and isinstance(m_url, str) and m_url.startswith("http"):
            try:
                r = http_requests.get(m_url, timeout=60)
                if r.status_code == 200:
                    return r.content
            except Exception as e:
                print(f"⚠️ Error descargando media.url {m_url}: {e}", flush=True)

    # 2. Si viene como Base64 en body
    body_val = data.get("body")
    if body_val and isinstance(body_val, str):
        if "base64," in body_val:
            try:
                b64_clean = body_val.split("base64,")[1]
                return base64.b64decode(b64_clean)
            except Exception:
                pass
        elif len(body_val) > 200 and not body_val.startswith("!"):
            try:
                return base64.b64decode(body_val)
            except Exception:
                pass
                
    # 3. Si viene mediaUrl o url en la raíz de data
    media_url = data.get("mediaUrl") or data.get("url")
    if media_url and isinstance(media_url, str) and media_url.startswith("http"):
        try:
            r = http_requests.get(media_url, timeout=60)
            if r.status_code == 200:
                return r.content
        except Exception as e:
            print(f"⚠️ Error descargando mediaUrl raíz {media_url}: {e}", flush=True)

    # 4. Intentar consultar endpoint /media de OpenWA REST API
    if session_id and chat_id and message_id:
        clean_msg_id = str(message_id).replace("/", "%2F")
        clean_chat_id = str(chat_id).replace("/", "%2F")
        try:
            r = http_requests.get(
                f"{OPENWA_URL}/api/sessions/{session_id}/messages/{clean_chat_id}/{clean_msg_id}/media",
                headers={"x-api-key": OPENWA_KEY},
                timeout=60
            )
            if r.status_code == 200:
                return r.content
        except Exception:
            pass

        try:
            r = http_requests.get(
                f"{OPENWA_URL}/api/sessions/{session_id}/messages/{clean_msg_id}/media",
                headers={"x-api-key": OPENWA_KEY},
                timeout=60
            )
            if r.status_code == 200:
                return r.content
        except Exception:
            pass

    return None


def tarea_actualizar_anexo(session_id: str, chat_id: str, tipo_anexo: str, empresa: str, archivo_bytes: bytes):
    """Ejecuta en segundo plano la actualización y sobreescritura de Anexos A1 o A5 en SQLite."""
    try:
        session_id = session_id or obtener_session_id_activo()
        from handlers.actualizar_anexos import actualizar_anexo_1_desde_bytes, actualizar_anexo_5_desde_bytes
        
        tipo_clean = tipo_anexo.lower().strip()
        empresa_clean = empresa.lower().strip()
        
        if tipo_clean in ("a1", "anexo1", "anexo_1", "1"):
            res = actualizar_anexo_1_desde_bytes(archivo_bytes, empresa_clean)
            if res.get("success"):
                msg = (
                    f"✅ *Actualización Exitosa de Anexo A1*\n\n"
                    f"🏢 *Empresa*: {empresa_clean.upper()}\n"
                    f"📊 *Registros de Frecuencia Insertados*: {res['total_registros']}\n"
                    f"🚌 *Servicios/Variantes Actualizados*: {res['servicios_actualizados']}\n\n"
                    f"💾 _Base de datos SQLite actualizada y sobreescrita correctamente._"
                )
            else:
                msg = f"❌ *Error al actualizar Anexo A1 ({empresa_clean.upper()})*:\n\n{res.get('error')}"
                
        elif tipo_clean in ("a5", "anexo5", "anexo_5", "5"):
            res = actualizar_anexo_5_desde_bytes(archivo_bytes, empresa_clean)
            if res.get("success"):
                pts = res.get('total_puntos_control', 0)
                lpp = res.get('total_pasadas_lpp', 0)
                serv = res.get('servicios_actualizados') or res.get('variantes_actualizadas', 0)
                msg = (
                    f"✅ *Actualización Exitosa de Anexo A5 (PO)*\n\n"
                    f"🏢 *Empresa*: {empresa_clean.upper()}\n"
                    f"📍 *Puntos de Control Oficiales (PC)*: {pts}\n"
                    f"⏱️ *Pasadas Programadas (LPP)*: {lpp}\n"
                    f"🚌 *Servicios/Variantes*: {serv}\n\n"
                    f"💾 _Datos del Anexo 5 en SQLite sobreescritos con éxito._"
                )
            else:
                msg = f"❌ *Error al actualizar Anexo A5 ({empresa_clean.upper()})*:\n\n{res.get('error')}"
        else:
            msg = f"❌ Tipo de anexo '{tipo_anexo}' no reconocido. Usa `A1` (frecuencias) o `A5` (puntos de control)."
            
        enviar_mensaje(session_id, chat_id, msg)
        
    except Exception as e:
        print(f"❌ Error actualizando anexo en segundo plano: {e}", flush=True)
        if session_id:
            enviar_mensaje(session_id, chat_id, f"❌ Ocurrió un error al actualizar el anexo: {e}")


async def planificador_descargas():
    """Planifica y ejecuta la descarga diaria de datos a las 06:30 AM hora de Chile."""
    import asyncio
    from zoneinfo import ZoneInfo
    from datetime import datetime, timedelta
    
    chile_tz = ZoneInfo("America/Santiago")
    
    # Esperar a que la base de datos y la API estén estables al arrancar
    await asyncio.sleep(30)
    
    while True:
        try:
            ahora = datetime.now(chile_tz)
            proxima = ahora.replace(hour=6, minute=30, second=0, microsecond=0)
            if ahora >= proxima:
                proxima += timedelta(days=1)
                
            segundos_espera = (proxima - ahora).total_seconds()
            print(f"⏰ Planificador: Próxima descarga programada para {proxima} Chile. Esperando {segundos_espera:.1f} seg.", flush=True)
            
            await asyncio.sleep(segundos_espera)
            
            # Ejecutar la descarga automática
            print("⏰ Planificador: Iniciando descarga automática diaria de las 06:30 AM...", flush=True)
            
            # Obtener la sesión activa para notificar a los administradores
            session_id = obtener_session_id_activo()
            
            from descargar_datos import descargar_todos
            res = await asyncio.to_thread(descargar_todos)
            
            if session_id:
                # Crear mensaje de reporte
                detalles = ""
                for op, cant in res.get("resumen", {}).items():
                    detalles += f"- *{op.upper()}*: {cant} registros nuevos\n"
                    
                if res.get("success"):
                    msg = f"⏰ *Planificador Automático (06:30 AM)*\n\n✅ *Actualización exitosa*\n\n{detalles}"
                else:
                    msg = f"⏰ *Planificador Automático (06:30 AM)*\n\n⚠️ *Actualización con Errores*\n\n{detalles}\n❌ *Errores detectados*:\n"
                    for err in res.get("errores", []):
                        msg += f"- {err}\n"
                        
                for admin in ADMIN_CELLPHONES:
                    chat_id = f"{admin}@c.us"
                    enviar_mensaje(session_id, chat_id, msg)
            else:
                print("⚠️ Planificador: No se pudo enviar el reporte por WhatsApp porque no se encontró ninguna sesión activa 'nabla-bot'.", flush=True)
                
        except Exception as e:
            print(f"❌ Error en el planificador de descargas: {e}", flush=True)
            # En caso de error, esperar 5 minutos antes de reintentar el cálculo
            await asyncio.sleep(300)


@app.on_event("startup")
def on_startup():
    init_db()
    import asyncio
    asyncio.create_task(planificador_descargas())




def enviar_mensaje(session_id: str, chat_id: str, texto: str) -> bool:
    try:
        r = http_requests.post(
            f"{OPENWA_URL}/api/sessions/{session_id}/messages/send-text",
            json={"chatId": chat_id, "text": texto},
            headers={"x-api-key": OPENWA_KEY},
            timeout=30
        )
        if r.status_code in (200, 201):
            return True
        print(f"❌ Error enviando mensaje. Status: {r.status_code} - {r.text}", flush=True)
        return False
    except Exception as e:
        print(f"❌ Excepción enviando mensaje: {e}", flush=True)
        return False


def enviar_documento(session_id: str, chat_id: str, archivo_bytes: bytes, filename: str) -> bool:
    print(f"📤 Enviando documento '{filename}' ({len(archivo_bytes)/1024/1024:.2f} MB) a {chat_id}...", flush=True)
    mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    
    candidate_sessions = list(dict.fromkeys([s for s in [session_id, obtener_session_id_activo(), "nabla-bot"] if s]))
    download_url = f"http://nabla-webhook:8000/download/{filename}"
    
    # 1. Intentar streaming vía URL interna (evita pasar megabytes de Base64 por JSON)
    for sess in candidate_sessions:
        for endpoint in ["send-file", "send-document"]:
            url_target = f"{OPENWA_URL}/api/sessions/{sess}/messages/{endpoint}"
            try:
                # Esquema estándar A: file string directo con URL
                p1 = {"chatId": chat_id, "file": download_url, "filename": filename, "mimetype": mimetype}
                r = http_requests.post(url_target, json=p1, headers={"x-api-key": OPENWA_KEY}, timeout=120)
                if r.status_code in (200, 201):
                    print(f"✅ Documento '{filename}' enviado con éxito vía streaming URL ({endpoint}) a {chat_id}.", flush=True)
                    return True
                
                # Esquema estándar B: file object
                p2 = {"chatId": chat_id, "file": {"url": download_url, "filename": filename, "mimetype": mimetype}, "filename": filename}
                r = http_requests.post(url_target, json=p2, headers={"x-api-key": OPENWA_KEY}, timeout=120)
                if r.status_code in (200, 201):
                    print(f"✅ Documento '{filename}' enviado con éxito vía streaming URL objeto ({endpoint}) a {chat_id}.", flush=True)
                    return True
            except Exception as e:
                print(f"⚠️ Error intentando streaming URL en {endpoint}: {e}", flush=True)

    # 2. Fallback a Base64 estructurado
    b64 = base64.b64encode(archivo_bytes).decode()
    data_uri = f"data:{mimetype};base64,{b64}"
    
    for sess in candidate_sessions:
        for endpoint in ["send-document", "send-file"]:
            url_target = f"{OPENWA_URL}/api/sessions/{sess}/messages/{endpoint}"
            try:
                # Esquema C: file object con data
                p3 = {
                    "chatId": chat_id,
                    "file": {
                        "mimetype": mimetype,
                        "filename": filename,
                        "data": b64
                    },
                    "filename": filename
                }
                r = http_requests.post(url_target, json=p3, headers={"x-api-key": OPENWA_KEY}, timeout=180)
                if r.status_code in (200, 201):
                    print(f"✅ Documento '{filename}' enviado con éxito vía Base64 objeto ({endpoint}) a {chat_id}.", flush=True)
                    return True
                    
                # Esquema D: data URI string
                p4 = {"chatId": chat_id, "file": data_uri, "filename": filename}
                r = http_requests.post(url_target, json=p4, headers={"x-api-key": OPENWA_KEY}, timeout=180)
                if r.status_code in (200, 201):
                    print(f"✅ Documento '{filename}' enviado con éxito vía Data URI ({endpoint}) a {chat_id}.", flush=True)
                    return True
                else:
                    print(f"⚠️ Endpoint {url_target} retornó status {r.status_code}: {r.text}", flush=True)
            except Exception as e:
                print(f"⚠️ Excepción en Base64 {endpoint}: {e}", flush=True)

    print(f"❌ No fue posible despachar el documento '{filename}' tras agotar todos los métodos de OpenWA.", flush=True)
    return False


@app.post("/webhook")
async def recibir_evento(request: Request, background_tasks: BackgroundTasks):
    evento = await request.json()
    print(f"\n🔔 [DEBUG] Evento: {evento.get('event')}", flush=True)
    print(f"🔔 [DEBUG] Data completa: {evento.get('data')}", flush=True)

    data = evento.get("data", {})
    session_id = evento.get("sessionId") or data.get("sessionId") or obtener_session_id_activo()
    chat_id = data.get("chatId")
    from_me = data.get("fromMe", False)
    cuerpo = str(data.get("body") or "").strip().lower()

    if from_me:
        return {"status": "ok"}

    # Extraer y limpiar número de teléfono del remitente
    sender_clean = obtener_remitente_limpio(data, session_id)
    
    # Redirigir JID de destino si logramos resolver el número real del remitente
    # para evitar enviar mensajes a direcciones LID (@lid) que congelan/crashean a whatsapp-web.js
    if chat_id.endswith("@lid") and len(sender_clean) <= 12:
        chat_id = f"{sender_clean}@c.us"
        
    admins_clean = ["".join(filter(str.isdigit, admin)) for admin in ADMIN_CELLPHONES]
    es_admin = (sender_clean in admins_clean)
    es_permitido = es_admin or es_usuario_permitido(sender_clean)

    # Validar si el mensaje es un comando del bot
    comandos_validos = ("!actualizar", "!puntualidad", "!icf", "!ip", "!permisos", "!reporte_tv", "!reporte_velocidad", "!velocidades")
    es_comando = any(cuerpo.startswith(cmd) for cmd in comandos_validos)

    if es_comando:
        if not es_permitido:
            print(f"❌ Acceso denegado para remitente: {sender_clean}", flush=True)
            registrar_intento_denegado(sender_clean)
            # Ignorar silenciosamente sin enviar mensaje de WhatsApp al LID no autorizado.
            # Esto previene que el motor de whatsapp-web.js se caiga al intentar enviar mensajes a IDs de privacidad.
            return {"status": "ok"}

    # Comando !permisos (Solo Administrador)
    if cuerpo.startswith("!permisos"):
        if not es_admin:
            enviar_mensaje(
                session_id, 
                chat_id, 
                "❌ *Acceso Denegado*. Solo el administrador principal puede gestionar permisos."
            )
            return {"status": "ok"}
            
        import re
        parts = re.split(r'\s+', cuerpo)
        
        # !permisos quitar <numero>
        if len(parts) >= 3 and parts[1] == "quitar":
            target = "".join(filter(str.isdigit, parts[2]))
            if not target:
                enviar_mensaje(session_id, chat_id, "❌ Indica un número válido. Ejemplo: `!permisos quitar 56912345678`.")
            else:
                exito = quitar_usuario_permitido(target)
                
                # Intentar resolver y quitar JID/LID asociado
                resolved_id = None
                try:
                    r = http_requests.get(
                        f"{OPENWA_URL}/api/sessions/{session_id}/contacts/check/{target}",
                        headers={"x-api-key": OPENWA_KEY}
                    )
                    if r.status_code in (200, 201):
                        jid = r.json().get("data")
                        if jid and isinstance(jid, str):
                            clean_jid = "".join(filter(str.isdigit, jid.split('@')[0]))
                            if clean_jid and clean_jid != target:
                                resolved_id = clean_jid
                except Exception as e:
                    print(f"⚠️ Error al verificar número {target} en OpenWA para quitar: {e}", flush=True)
                
                if resolved_id:
                    quitar_usuario_permitido(resolved_id)
                
                if exito:
                    enviar_mensaje(session_id, chat_id, f"✅ El número *{target}* fue eliminado de la lista de autorizados.")
                else:
                    enviar_mensaje(session_id, chat_id, f"⚠️ El número *{target}* no estaba en la lista de autorizados.")
            return {"status": "ok"}
            
        # !permisos <numero>
        elif len(parts) >= 2 and parts[1] != "":
            target = "".join(filter(str.isdigit, parts[1]))
            if not target:
                enviar_mensaje(session_id, chat_id, "❌ Indica un número válido. Ejemplo: `!permisos 56912345678`.")
            else:
                exito = agregar_usuario_permitido(target)
                
                # Intentar consultar en OpenWA si tiene un JID/LID asociado
                resolved_id = None
                try:
                    r = http_requests.get(
                        f"{OPENWA_URL}/api/sessions/{session_id}/contacts/check/{target}",
                        headers={"x-api-key": OPENWA_KEY}
                    )
                    if r.status_code in (200, 201):
                        res_data = r.json()
                        jid = res_data.get("data")
                        if jid and isinstance(jid, str):
                            clean_jid = "".join(filter(str.isdigit, jid.split('@')[0]))
                            if clean_jid and clean_jid != target:
                                resolved_id = clean_jid
                except Exception as e:
                    print(f"⚠️ Error al verificar número {target} en OpenWA: {e}", flush=True)
                
                # Si se obtuvo un ID diferente (ej. un LID), agregarlo también
                msg_add = ""
                if resolved_id:
                    agregar_usuario_permitido(resolved_id)
                    msg_add = f" (ID de privacidad asociado: *{resolved_id}*)"
                
                if exito:
                    enviar_mensaje(session_id, chat_id, f"✅ El número *{target}* ahora está autorizado para usar el bot.{msg_add}")
                else:
                    enviar_mensaje(session_id, chat_id, f"⚠️ El número *{target}* ya se encuentra autorizado.{msg_add}")
            return {"status": "ok"}
            
        # !permisos (listar)
        else:
            autorizados = listar_usuarios_permitidos()
            msg = "👥 *Usuarios Autorizados en el Bot*:\n\n"
            msg += "👑 *Administrador(es)*:\n"
            for admin in ADMIN_CELLPHONES:
                msg += f"- {admin}\n"
            msg += "\n"
            msg += "👤 *Celulares Permitidos*:\n"
            if autorizados:
                for cel in autorizados:
                    msg += f"- {cel}\n"
            else:
                msg += "_(Ninguno además del administrador)_\n"
                
            # Mostrar intentos denegados recientes para facilitar copia
            if INTENTOS_DENEGADOS:
                msg += "\n🚫 *Intentos denegados recientes*:\n"
                for intento in reversed(INTENTOS_DENEGADOS):
                    diff = datetime.now() - intento["time"]
                    minutos = int(diff.total_seconds() / 60)
                    hace = f"hace {minutos} min" if minutos > 0 else "hace instantes"
                    msg += f"- `{intento['sender']}` ({hace}) _-> Para autorizar: `!permisos {intento['sender']}`_\n"
                    
            msg += "\n💡 _Para agregar: `!permisos <numero>`_\n_Para quitar: `!permisos quitar <numero>`_"
            enviar_mensaje(session_id, chat_id, msg)
            return {"status": "ok"}

    # Comando !actualizar
    if cuerpo.startswith("!actualizar"):
        import re
        parts = re.split(r'\s+', cuerpo)
        
        tipo_anexo = None
        empresa = None
        
        for p in parts[1:]:
            p_clean = p.strip().lower()
            if p_clean in ("a1", "anexo1", "anexo_1", "1"):
                tipo_anexo = "A1"
            elif p_clean in ("a5", "anexo5", "anexo_5", "5"):
                tipo_anexo = "A5"
            elif p_clean in ("tasacop", "tasacoop", "lider", "toptur"):
                empresa = "tasacop" if "tasa" in p_clean else p_clean

        # Caso A: Actualización de Anexo (A1 o A5)
        if tipo_anexo or empresa:
            if not tipo_anexo or not empresa:
                enviar_mensaje(
                    session_id, 
                    chat_id, 
                    "❌ Por favor especifica el tipo de anexo (`A1` o `A5`) y la empresa (`tasacop`, `lider`, `toptur`).\n\n"
                    "💡 *Ejemplo*: `!actualizar A1 tasacop` o `!actualizar A5 lider` (adjuntando el archivo Excel o enviándolo a continuación)."
                )
                return {"status": "ok"}
                
            # Intentar extraer archivo adjunto en este mismo mensaje
            message_id = data.get("id")
            archivo_bytes = extraer_archivo_bytes(session_id, chat_id, message_id, data)
            
            if archivo_bytes:
                enviar_mensaje(
                    session_id, 
                    chat_id, 
                    f"⏳ Procesando y sobreescribiendo *Anexo {tipo_anexo}* para *{empresa.upper()}* en SQLite..."
                )
                background_tasks.add_task(tarea_actualizar_anexo, session_id, chat_id, tipo_anexo, empresa, archivo_bytes)
            else:
                guardar_estado(chat_id, estado="esperando_anexo", nombre_operacion=f"{tipo_anexo}_{empresa}")
                enviar_mensaje(
                    session_id, 
                    chat_id, 
                    f"📥 *Listo para actualizar Anexo {tipo_anexo} ({empresa.upper()})*\n\n"
                    f"📎 Por favor envía el archivo Excel (`.xlsx` o `.xls`) a continuación en este chat para sobreescribir la base de datos."
                )
            return {"status": "ok"}

        # Caso B: Actualización automática periódica de expediciones telemáticas
        print("✅ ¡Comando !actualizar detectado!", flush=True)
        enviar_mensaje(session_id, chat_id, "⏳ Iniciando descarga y actualización de expediciones para todos los operadores en segundo plano. Te avisaré cuando termine...")
        background_tasks.add_task(tarea_actualizar_datos, session_id, chat_id)
        return {"status": "ok"}

    # Comando !puntualidad
    if "!puntualidad" in cuerpo:
        print("✅ ¡Comando detectado!", flush=True)
        enviar_mensaje(session_id, chat_id, "⏳ Procesando reporte de puntualidad desde la base de datos SQLite...")
        background_tasks.add_task(tarea_calcular_puntualidad, session_id, chat_id)
        return {"status": "ok"}

    # Comando !icf
    if cuerpo.startswith("!icf"):
        import re
        parts = re.split(r'\s+', cuerpo)
        
        meses_map = {
            "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
            "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
            "noviembre": 11, "diciembre": 12
        }
        
        operador = None
        mes = None
        anio = None
        
        for part in parts[1:]:
            part_clean = part.strip().lower()
            if part_clean in ("lider", "tasacop", "toptur"):
                operador = part_clean
            elif part_clean in meses_map:
                mes = meses_map[part_clean]
            elif part_clean.isdigit():
                val = int(part_clean)
                if 1 <= val <= 12 and mes is None:
                    mes = val
                elif 2000 <= val <= 2100:
                    anio = val
                    
        hoy = datetime.now()
        if anio is None:
            anio = hoy.year
        if mes is None:
            mes = hoy.month
        if operador is None:
            enviar_mensaje(session_id, chat_id, "❌ Por favor especifica el operador. Ejemplo: `!icf lider mayo 2026`.\n\nOperadores disponibles: *lider*, *tasacop*, *toptur*.")
            return {"status": "ok"}
            
        print(f"✅ Comando !icf detectado: operador={operador}, mes={mes}, anio={anio}", flush=True)
        enviar_mensaje(session_id, chat_id, f"⏳ Calculando reporte ICF para *{operador.upper()}* ({mes:02d}/{anio}). Espera un momento...")
        background_tasks.add_task(tarea_calcular_icf, session_id, chat_id, operador, anio, mes)
        return {"status": "ok"}

    # Comando !ip (Indicador de Puntualidad)
    if cuerpo.startswith("!ip"):
        import re
        from handlers.ip import MESES_MAP
        parts = re.split(r'\s+', cuerpo)
        
        empresa = None
        mes = None
        anio = None
        
        for part in parts[1:]:
            part_clean = part.strip().lower()
            if part_clean in ("lider", "tasacop", "tasacoop", "toptur"):
                empresa = "tasacop" if "tasa" in part_clean else part_clean
            elif part_clean in MESES_MAP:
                mes = MESES_MAP[part_clean][0]
            elif part_clean.isdigit():
                val = int(part_clean)
                if 1 <= val <= 12 and mes is None:
                    mes = val
                elif 2000 <= val <= 2100:
                    anio = val
                elif 20 <= val <= 99:
                    anio = 2000 + val
            else:
                # Comprobar si viene compuesto (ej. mayo26, mayo2026, 0526)
                m_match = re.match(r"([a-z]+)(\d+)", part_clean)
                if m_match:
                    m_txt, a_txt = m_match.group(1), m_match.group(2)
                    if m_txt in MESES_MAP and mes is None:
                        mes = MESES_MAP[m_txt][0]
                    if a_txt.isdigit() and anio is None:
                        a_val = int(a_txt)
                        anio = 2000 + a_val if a_val < 100 else a_val
                        
        hoy = datetime.now()
        if anio is None:
            anio = hoy.year
        if mes is None:
            mes = hoy.month
        if empresa is None:
            enviar_mensaje(
                session_id, 
                chat_id, 
                "❌ *Parámetros incompletos para !ip*.\n\n"
                "📌 *Sintaxis*: `!ip <empresa> <mes> <anio>`\n"
                "🏢 *Empresas*: `tasacop`, `lider`, `toptur`\n"
                "💡 *Ejemplo*: `!ip tasacop mayo 2026` o `!ip lider 5 2026`"
            )
            return {"status": "ok"}
            
        mes_nombre = f"{mes:02d}"
        for k, v in MESES_MAP.items():
            if v[0] == mes:
                mes_nombre = v[2]
                break
                
        print(f"✅ Comando !ip detectado: empresa={empresa}, mes={mes} ({mes_nombre}), anio={anio}", flush=True)
        enviar_mensaje(
            session_id, 
            chat_id, 
            f"⏳ Calculando *Indicador de Puntualidad (IP)* para *{empresa.upper()}* ({mes_nombre} {anio})... Por favor espera un momento."
        )
        background_tasks.add_task(tarea_calcular_ip, session_id, chat_id, empresa, anio, mes)
        return {"status": "ok"}


    # Comando !reporte_tv / !reporte_velocidad (Tasacoop PO A5)
    if any(cuerpo.startswith(prefix) for prefix in ("!reporte_tv", "!reporte_velocidad", "!velocidades")):
        import re
        from handlers.velocidades import MESES_MAP

        parts = re.split(r'\s+', cuerpo)
        mes_str = None
        anio = None

        for part in parts[1:]:
            p_clean = part.strip().lower()
            if p_clean in MESES_MAP:
                mes_str = p_clean
            elif p_clean.isdigit():
                val = int(p_clean)
                if 1 <= val <= 12 and mes_str is None:
                    # Encontrar nombre de mes por número
                    for k, v in MESES_MAP.items():
                        if v[0] == val:
                            mes_str = k
                            break
                elif 2000 <= val <= 2100:
                    anio = val

        if anio is None:
            anio = datetime.now().year

        if not mes_str:
            enviar_mensaje(
                session_id, 
                chat_id, 
                "❌ Por favor especifica el mes a consultar.\n\nEjemplo: `!reporte_tv mayo` o `!reporte_tv junio 2026`.\n\n💡 _Este reporte genera el análisis cinemático de velocidades y tiempos de viaje (PO A5) para Tasacoop._"
            )
            return {"status": "ok"}

        mes_nombre = MESES_MAP[mes_str][2]
        print(f"✅ Comando !reporte_tv detectado: Tasacop, mes={mes_str}, anio={anio}", flush=True)
        enviar_mensaje(
            session_id, 
            chat_id, 
            f"⏳ Generando reporte cinemático de velocidades y tiempos de viaje para *Tasacoop* ({mes_nombre} {anio})... Esto tomará unos momentos."
        )
        background_tasks.add_task(tarea_generar_reporte_tv, session_id, chat_id, mes_str, anio)
        return {"status": "ok"}

    # Manejo de estados conversacionales (ej: recepción de archivo de Anexo pendiente)
    estado = obtener_estado(chat_id)
    if estado and estado.estado == "esperando_anexo":
        message_id = data.get("id")
        archivo_bytes = extraer_archivo_bytes(session_id, chat_id, message_id, data)
        
        if archivo_bytes:
            op_parts = estado.nombre_operacion.split("_")
            tipo_anexo = op_parts[0] if len(op_parts) > 0 else "A1"
            empresa = op_parts[1] if len(op_parts) > 1 else "tasacop"
            eliminar_estado(chat_id)
            
            enviar_mensaje(
                session_id, 
                chat_id, 
                f"⏳ Archivo recibido. Procesando y sobreescribiendo *Anexo {tipo_anexo.upper()}* para *{empresa.upper()}* en SQLite..."
            )
            background_tasks.add_task(tarea_actualizar_anexo, session_id, chat_id, tipo_anexo, empresa, archivo_bytes)
            return {"status": "ok"}
        else:
            print("⚠️ Mensaje recibido en estado 'esperando_anexo' sin archivo válido adjunto.", flush=True)

    return {"status": "ok"}


if __name__ == "__main__":
    uvicorn.run("webhook:app", host="0.0.0.0", port=8000)
