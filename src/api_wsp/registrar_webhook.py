import requests

def registrar_webhook(url: str, headers:dict, sesion_id:str, webhook_url: str)-> dict|None:
    """Registra el webhook en OpenWA para la sesion"""
    payload = {
        "url": webhook_url,
        "events": [
            "message.received",
            "message.sent",
            "session.status",
            "session.authenticated",
            "session.disconnected"
        ]
    }
    try:
        r = requests.post(f"{url}/api/sessions/{sesion_id}/webhooks", json = payload, headers = headers)
        if r.status_code in (200, 201):
            print(f"webhook registrado: {webhook_url}")
            return r.json()
        
        print(f"Error registrando webhook. status: {r.status_code} - {r.text}")
        return None
    except Exception as e:
        print(f"Error registrando webhook: {e}")
        return None

def listar_webhooks(url: str, headers: dict, sesion_id: str) -> list:
    """Lista webhooks registrados para la sesion."""
    try:
        r = requests.get(f"{url}/api/sessions/{sesion_id}/webhooks", headers=headers)
        return r.json() if r.status_code == 200 else []
    except Exception as e:
        print(f"Error listando webhooks: {e}")
        return []