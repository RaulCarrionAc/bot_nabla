from fastapi import FastAPI, Request
import uvicorn
import base64
import os
import requests as http_requests
from dotenv import load_dotenv
from estados import EstadoConversacion
from handlers.puntualidad import procesar_puntualidad

load_dotenv()

app = FastAPI()
estado_chats: dict[str, EstadoConversacion] = {}

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
async def recibir_evento(request: Request):
    evento = await request.json()
    print(f"\n🔔 [DEBUG] Evento: {evento.get('event')}", flush=True)
    print(f"🔔 [DEBUG] Data completa: {evento.get('data')}", flush=True)

    data = evento.get("data", {})
    tipo_msg = data.get("type")
    cuerpo = str(data.get("body") or "").strip().lower()
    session_id = evento.get("sessionId")
    chat_id = data.get("chatId")
    from_me = data.get("fromMe", False)

    if from_me:
        return {"status": "ok"}

    # Comando !puntualidad
    if "!puntualidad" in cuerpo:
        print("✅ ¡Comando detectado!", flush=True)
        estado_chats[chat_id] = EstadoConversacion(chat_id)
        ok = enviar_mensaje(session_id, chat_id,
            "📊 *Análisis de Puntualidad*\n\nEnvía el archivo *Anexo 5* (A5) en formato .xlsx")
        print(f"📤 [DEBUG] Resultado envío mensaje: {ok}", flush=True)
        return {"status": "ok"}

    estado = estado_chats.get(chat_id)
    if estado is None:
        return {"status": "ok"}

    # Recibir documento
    if tipo_msg == "document":
        media = data.get("media", {})
        b64_data = media.get("data", "")
        filename = media.get("filename", "")
        print(f"📎 [DEBUG] Documento recibido: {filename}", flush=True)

        if not b64_data:
            enviar_mensaje(session_id, chat_id, "❌ No se pudo leer el archivo.")
            return {"status": "ok"}

        archivo_bytes = base64.b64decode(b64_data)

        if not estado.tiene_a5:
            estado.bytes_a5 = archivo_bytes
            estado.tiene_a5 = True
            print("✅ [DEBUG] Anexo 5 guardado", flush=True)
            enviar_mensaje(session_id, chat_id,
                "✅ Anexo 5 recibido.\n\nAhora envía el archivo de *Operación* en formato .xlsx")
            return {"status": "ok"}

        if not estado.tiene_operacion:
            estado.bytes_operacion = archivo_bytes
            estado.nombre_operacion = filename.replace(".xlsx", "")
            estado.tiene_operacion = True
            print("✅ [DEBUG] Operación guardada, procesando...", flush=True)
            enviar_mensaje(session_id, chat_id, "⏳ Procesando datos, espera un momento...")

            try:
                excel_bytes, resumen = procesar_puntualidad(
                    estado.bytes_operacion,
                    estado.bytes_a5,
                    estado.nombre_operacion
                )
                enviar_mensaje(session_id, chat_id, resumen)
                enviar_documento(session_id, chat_id, excel_bytes,
                    f"puntualidad_{estado.nombre_operacion}.xlsx")
                print("✅ [DEBUG] Proceso completo", flush=True)
            except Exception as e:
                print(f"❌ [DEBUG] Error procesando: {e}", flush=True)
                enviar_mensaje(session_id, chat_id, f"❌ Error procesando: {e}")
            finally:
                del estado_chats[chat_id]

    return {"status": "ok"}


if __name__ == "__main__":
    uvicorn.run("webhook:app", host="0.0.0.0", port=8000)
