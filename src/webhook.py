from fastapi import FastAPI, Request, BackgroundTasks
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
from datetime import datetime

load_dotenv()

ADMIN_CELLPHONES = [num.strip() for num in os.getenv("ADMIN_CELLPHONE", "56989158197").split(",") if num.strip()]

def obtener_remitente_limpio(data: dict, session_id: str = None) -> str:
    """Extrae y resuelve el número de celular limpio del remitente (soportando LID)."""
    sender_id = data.get("author") or data.get("sender", {}).get("id") or data.get("from") or ""
    
    # 1. Si ya viene resuelto en el payload
    if data.get("senderPhone"):
        raw_phone = str(data.get("senderPhone")).split('@')[0]
        return "".join(filter(str.isdigit, raw_phone))
        
    # 2. Si es un LID y tenemos la sesión, intentar resolver a JID/Teléfono
    if str(sender_id).endswith("@lid") and session_id:
        try:
            r = http_requests.get(
                f"{OPENWA_URL}/api/sessions/{session_id}/contacts/{sender_id}/phone",
                headers={"x-api-key": OPENWA_KEY}
            )
            if r.status_code in (200, 201):
                phone = r.json().get("phone")
                if phone:
                    return "".join(filter(str.isdigit, str(phone).split('@')[0]))
        except Exception as e:
            print(f"⚠️ Error al resolver LID {sender_id} a teléfono: {e}", flush=True)
            
    # 3. Fallback
    raw_num = str(sender_id).split('@')[0]
    return "".join(filter(str.isdigit, raw_num))

app = FastAPI()

@app.get("/health")
def health_check():
    return {"status": "ok"}


def tarea_actualizar_datos(session_id: str, chat_id: str):
    from descargar_datos import descargar_todos
    try:
        descargar_todos()
        enviar_mensaje(
            session_id, 
            chat_id, 
            "✅ *Actualización exitosa*. Se descargaron e insertaron en la base de datos las expediciones de los últimos 7 días para todos los operadores."
        )
    except Exception as e:
        enviar_mensaje(
            session_id, 
            chat_id, 
            f"❌ *Error durante la actualización*: {e}"
        )

@app.on_event("startup")
def on_startup():
    init_db()

OPENWA_URL = os.getenv("URL", "http://openwa-api:2785")
OPENWA_KEY = os.getenv("API-KEY")


def enviar_mensaje(session_id: str, chat_id: str, texto: str) -> bool:
    try:
        r = http_requests.post(
            f"{OPENWA_URL}/api/sessions/{session_id}/messages/send-text",
            json={"chatId": chat_id, "text": texto},
            headers={"x-api-key": OPENWA_KEY}
        )
        if r.status_code in (200, 201):
            return True
        print(f"❌ Error enviando mensaje. Status: {r.status_code} - {r.text}", flush=True)
        return False
    except Exception as e:
        print(f"❌ Excepción enviando mensaje: {e}", flush=True)
        return False


def enviar_documento(session_id: str, chat_id: str, archivo_bytes: bytes, filename: str) -> bool:
    b64 = base64.b64encode(archivo_bytes).decode()
    try:
        r = http_requests.post(
            f"{OPENWA_URL}/api/sessions/{session_id}/messages/send-document",
            json={
                "chatId": chat_id,
                "base64": b64,
                "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "filename": filename,
            },
            headers={"x-api-key": OPENWA_KEY}
        )
        if r.status_code in (200, 201):
            return True
        print(f"❌ Error enviando documento. Status: {r.status_code} - {r.text}", flush=True)
        return False
    except Exception as e:
        print(f"❌ Excepción enviando documento: {e}", flush=True)
        return False


@app.post("/webhook")
async def recibir_evento(request: Request, background_tasks: BackgroundTasks):
    evento = await request.json()
    print(f"\n🔔 [DEBUG] Evento: {evento.get('event')}", flush=True)
    print(f"🔔 [DEBUG] Data completa: {evento.get('data')}", flush=True)

    data = evento.get("data", {})
    session_id = evento.get("sessionId")
    chat_id = data.get("chatId")
    from_me = data.get("fromMe", False)
    cuerpo = str(data.get("body") or "").strip().lower()

    if from_me:
        return {"status": "ok"}

    # Extraer y limpiar número de teléfono del remitente
    sender_clean = obtener_remitente_limpio(data, session_id)
    admins_clean = ["".join(filter(str.isdigit, admin)) for admin in ADMIN_CELLPHONES]
    es_admin = (sender_clean in admins_clean)
    es_permitido = es_admin or es_usuario_permitido(sender_clean)

    # Validar si el mensaje es un comando del bot
    comandos_validos = ("!actualizar", "!puntualidad", "!icf", "!permisos")
    es_comando = any(cuerpo.startswith(cmd) for cmd in comandos_validos)

    if es_comando:
        if not es_permitido:
            print(f"❌ Acceso denegado para remitente: {sender_clean}", flush=True)
            enviar_mensaje(
                session_id, 
                chat_id, 
                "❌ *Acceso Denegado*. Tu número de teléfono no está autorizado para ejecutar comandos en este bot. Contacta al administrador."
            )
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
                if exito:
                    enviar_mensaje(session_id, chat_id, f"✅ El número *{target}* ahora está autorizado para usar el bot.")
                else:
                    enviar_mensaje(session_id, chat_id, f"⚠️ El número *{target}* ya se encuentra autorizado.")
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
            msg += "\n💡 _Para agregar: `!permisos <numero>`_\n_Para quitar: `!permisos quitar <numero>`_"
            enviar_mensaje(session_id, chat_id, msg)
            return {"status": "ok"}

    # Comando !actualizar
    if "!actualizar" in cuerpo:
        print("✅ ¡Comando !actualizar detectado!", flush=True)
        enviar_mensaje(session_id, chat_id, "⏳ Iniciando descarga y actualización de expediciones para todos los operadores en segundo plano. Te avisaré cuando termine...")
        background_tasks.add_task(tarea_actualizar_datos, session_id, chat_id)
        return {"status": "ok"}

    # Comando !puntualidad
    if "!puntualidad" in cuerpo:
        print("✅ ¡Comando detectado!", flush=True)
        enviar_mensaje(session_id, chat_id, "⏳ Procesando reporte de puntualidad desde la base de datos SQLite...")
        
        try:
            db_path = DATABASE_URL.replace("sqlite:///", "")
            excel_bytes, resumen = procesar_puntualidad_desde_db(db_path)
            
            enviar_mensaje(session_id, chat_id, resumen)
            enviar_documento(session_id, chat_id, excel_bytes, "puntualidad_reporte.xlsx")
            print("✅ [DEBUG] Proceso completo desde base de datos", flush=True)
        except Exception as e:
            print(f"❌ [DEBUG] Error procesando desde base de datos: {e}", flush=True)
            enviar_mensaje(session_id, chat_id, "❌ Error al procesar los datos. Por favor verifica que las tablas 'operaciones' y 'anexo_5' en SQLite contengan registros válidos.")
        
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
        
        try:
            db_path = DATABASE_URL.replace("sqlite:///", "")
            reporte_bytes, proy_bytes, resumen_txt = ejecutar_calculo_icf(operador, anio, mes, db_path)
            
            enviar_mensaje(session_id, chat_id, resumen_txt)
            enviar_documento(session_id, chat_id, reporte_bytes, f"reporte_{operador}_{mes:02d}_{anio}.xlsx")
            enviar_documento(session_id, chat_id, proy_bytes, f"reporte_proyeccion_{operador}_{mes:02d}_{anio}.xlsx")
            print("✅ [DEBUG] Reportes ICF enviados con éxito", flush=True)
        except Exception as e:
            print(f"❌ [DEBUG] Error calculando ICF: {e}", flush=True)
            enviar_mensaje(session_id, chat_id, f"❌ Error calculando ICF: Asegúrate de que las frecuencias y expediciones estén cargadas para la fecha indicada.")
            
        return {"status": "ok"}

    # Ejemplo de uso de persistencia de estado para otros comandos futuros
    estado = obtener_estado(chat_id)
    if estado is None:
        return {"status": "ok"}

    return {"status": "ok"}


if __name__ == "__main__":
    uvicorn.run("webhook:app", host="0.0.0.0", port=8000)
