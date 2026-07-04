import requests
import time
import os
from typing import Dict, Any

def limpiar_webhooks(url: str, headers: dict, session_id: str):
    """Elimina todos los webhooks activos de la sesion de forma limpia"""
    endpoint = f"{url}/api/sessions/{session_id}/webhooks"
    try:
        r = requests.get(endpoint, headers=headers)
        
        if r.status_code != 200:
            return
            
        webhooks = r.json()
        if not isinstance(webhooks, list):
            return

        print(f"🧹 Limpiando webhooks antiguos atascados: {len(webhooks)}")
        for wh in webhooks:
            if isinstance(wh, dict) and "id" in wh:
                wh_id = wh.get("id")
                rd = requests.delete(f"{endpoint}/{wh_id}", headers=headers)
                print(f"   ↳ Eliminado antiguo ID {wh_id}: Status {rd.status_code}")
                time.sleep(0.2)
    except Exception as e:
        print(f"⚠️ Aviso al limpiar webhooks: {e}")

def asegurar_webhook(url: str, headers: dict, session_id: str) -> Dict[str, Any] | None:
    """Verifica si el webhook ya existe. Si no, limpia y crea uno nuevo."""
    webhook_url = os.getenv("WEBHOOK_URL", "http://webhook.nabla.net:8000/webhook")
    endpoint = f"{url}/api/sessions/{session_id}/webhooks"

    print("🔎 Verificando registros de webhooks en la API...")
    try:
        r = requests.get(endpoint, headers=headers)
        
        if r.status_code == 200:
            webhooks_actuales = r.json()
            
            if isinstance(webhooks_actuales, list):
                for wh in webhooks_actuales:
                    if isinstance(wh, dict) and wh.get("url") == webhook_url:
                        print(f"✅ Webhook ya configurado y activo apuntando a: {webhook_url}")
                        return wh
        
        print(f"⚙️ Webhook no encontrado o desincronizado. Configurando hacia: {webhook_url}")
        limpiar_webhooks(url, headers, session_id)
        
        payload = {
            "url": webhook_url,
            "events": [
                "message.received",
                "session.status"
            ]
        }

        r_post = requests.post(endpoint, json=payload, headers=headers)
        respuesta = r_post.json()

        if r_post.status_code in [200, 201]:
            print("✅ Webhook definitivo creado exitosamente")
            return respuesta
        else:
            print("❌ Error al crear el webhook:")
            print(respuesta)
            return None

    except Exception as e:
        print(f"❌ Error de red conectando con la API en asegurar_webhook: {e}")
        return None
