from fastapi import FastAPI, Request, HTTPException
from pprint import pprint
import uvicorn
import base64
import requests as http_requests
from estados import EstadoConversacion
from handlers.puntualidad import procesar_puntualidad

app = FastAPI()
estado_chats: dict[str, EstadoConversacion] = {}

OPENWA_URL = "http://openwa-api:2785"
OPENWA_KEY = "dev-admin-key"

def enviar_mensaje(session_id: str, chat_id: str, texto: str):
    http_requests.post(
        f"{OPENWA_URL}/api/sessions/{session_id}/messages/send-text",
        json={"chatId": chat_id, "text": texto},
        headers={"x-api-key": OPENWA_KEY}
    )

def enviar_documento(session_id: str, chat_id: str, excel_bytes: bytes, filename: str):
    b64 = base64.b64encode(excel_bytes).decode()
    http_requests.post(
        f"{OPENWA_URL}/api/sessions/{session_id}/messages/send-document",
        json={
            "chatId": chat_id,
            "base64": b64,
            "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "filename": filename,
        },
        headers={"x-api-key": OPENWA_KEY}
    )

@app.post("/webhook")
async def recibir_evento(request: Request):
    evento = await request.json()
    tipo_evento = evento.get("event")
    data = evento.get("data", {})
    session_id = evento.get("sessionId")

    if tipo_evento != "message.received":
        return {"status": "ok"}

    chat_id = data.get("chatId")
    cuerpo = data.get("body", "").strip()
    tipo_msg = data.get("type")
    from_me = data.get("fromMe", False)

    if from_me or tipo_msg == "e2e_notification":
        return {"status": "ok"}

    estado = estado_chats.get(chat_id)

    # Comando !puntualidad
    if cuerpo.lower() == "!puntualidad":
        estado_chats[chat_id] = EstadoConversacion(chat_id)
        enviar_mensaje(session_id, chat_id,
            "📊 *Análisis de Puntualidad*\n\nEnvía el archivo *Anexo 5* (A5) en formato .xlsx")
        return {"status": "ok"}

    if estado is None:
        return {"status": "ok"}

    # Recibir documento
    if tipo_msg == "document":
        media = data.get("media", {})
        b64_data = media.get("data", "")
        filename = media.get("filename", "")

        if not b64_data:
            enviar_mensaje(session_id, chat_id, "❌ No se pudo leer el archivo.")
            return {"status": "ok"}

        archivo_bytes = base64.b64decode(b64_data)

        if not estado.tiene_a5:
            estado.bytes_a5 = archivo_bytes
            estado.tiene_a5 = True
            enviar_mensaje(session_id, chat_id,
                "✅ Anexo 5 recibido.\n\nAhora envía el archivo de *Operación* en formato .xlsx")
            return {"status": "ok"}

        if not estado.tiene_operacion:
            estado.bytes_operacion = archivo_bytes
            estado.nombre_operacion = filename.replace(".xlsx","")
            estado.tiene_operacion = True
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
            except Exception as e:
                enviar_mensaje(session_id, chat_id, f"❌ Error procesando: {e}")
            finally:
                del estado_chats[chat_id]

    return {"status": "ok"}

if __name__ == "__main__":
    uvicorn.run("webhook:app", host="0.0.0.0", port=8000, reload=True)