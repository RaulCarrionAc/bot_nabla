from dotenv import load_dotenv
import os

# --- modulos propios
from api_wsp.logging import listar_sesiones, crear_sesion, iniciar_sesion, verificar_estado, generar_qr, detener_sesion
from api_wsp.registrar_webhook import registrar_webhook, listar_webhooks
# ---

from pprint import pprint
import time

load_dotenv()


API_KEY = os.getenv("API-KEY")
URL = os.getenv("URL")
HEADERS = {
    "Content-Type": "application/json",
    "x-api-key": f"{API_KEY}",
}

WEBHOOK_URL = os.getenv("WEBHOOK_URL", "http://host.docker.internal:8000/webhook")

SESSION_NAME = "nabla-bot"
MAX_INTENTOS = 30
SLEEP_ENTRE_INTENTOS = 5
MAX_INTENTOS_AUTHENTICATING = 6  # 30s antes de hacer stop+restart

def obtener_o_crear_sesion() -> str | None:
    """Retorna el ID de la sesion 'nabla-bot', creandola si no existe"""
    sesiones = listar_sesiones(URL, HEADERS)

    if not sesiones:
        print("No se encontraron sesiones. Creando una nueva...")
        sesion = crear_sesion(URL, HEADERS)
        if not sesion:
            print("Error: no se pudo crear la sesion")
            return None
        pprint(sesion)
        return sesion.get("id")
    
    for sesion in sesiones:
        if sesion.get("name") == SESSION_NAME:
            print(f"Sesion '{SESSION_NAME}' encontrada")
            return sesion.get("id")
        
    print(f"No se encontro ninguna sesion con el nombre '{SESSION_NAME}'")
    return None

def esperar_sesion_lista(sesion_id: str) -> bool:
    """
    Maneja el ciclo de vida de la sesion:
      - created        → iniciar sesion
      - initializing   → esperando QR
      - qr_ready       → mostrar QR y esperar escaneo
      - authenticating → esperando confirmacion, si se traba: stop + restart
      - disconnected   → reiniciar sesion
      - ready          → autenticado y listo
    """
    qr_mostrado = False
    qr_escaneado = False  # ← nuevo flag
    intentos_authenticating = 0

    for intento in range(1, MAX_INTENTOS + 1):
        print(f"\nIntento {intento}/{MAX_INTENTOS}...")
        estado = verificar_estado(URL, HEADERS, sesion_id)

        if estado == "ready":
            print("Sesion lista y autenticada.")
            return True

        elif estado in ("initializing", "qr_ready"):
            if not qr_mostrado:
                print("QR disponible. Mostrando...")
                qr_mostrado = generar_qr(URL, HEADERS, sesion_id)
                if qr_mostrado:
                    print("Escanea el QR con WhatsApp para continuar...")
            else:
                print("Esperando que escanees el QR...")
            intentos_authenticating = 0

        elif estado == "authenticating":
            qr_escaneado = True  # si llegó aquí, el QR fue escaneado
            intentos_authenticating += 1
            print(f"Autenticando... ({intentos_authenticating}/{MAX_INTENTOS_AUTHENTICATING})")

            # Solo reiniciar si el QR nunca fue escaneado (sesion huerfana)
            if intentos_authenticating >= MAX_INTENTOS_AUTHENTICATING and not qr_escaneado:
                print("Autenticacion trabada sin scan. Reiniciando...")
                detener_sesion(URL, HEADERS, sesion_id)
                qr_mostrado = False
                qr_escaneado = False
                intentos_authenticating = 0

        elif estado in ("created", "disconnected"):
            print(f"Sesion en estado '{estado}'. Iniciando...")
            iniciar_sesion(URL, HEADERS, sesion_id)
            qr_mostrado = False
            intentos_authenticating = 0

        else:
            print(f"Estado desconocido: {estado}")

        time.sleep(SLEEP_ENTRE_INTENTOS)

    print("No se pudo completar la autenticacion tras los intentos.")
    return False
    
def main():
    sesion_id = obtener_o_crear_sesion()
    if not sesion_id:
        return

    print(f"ID de la sesion: {sesion_id}")
    if not esperar_sesion_lista(sesion_id):
        return
    
    #registrar el webhook solo si no existe ya
    webhooks = listar_webhooks(URL, HEADERS, sesion_id)
    if not any(w.get("url") == WEBHOOK_URL for w in webhooks):
        registrar_webhook(URL, HEADERS, sesion_id, WEBHOOK_URL)
    else:
        print("Webhook ya registrado")

if __name__ == "__main__":
    main()