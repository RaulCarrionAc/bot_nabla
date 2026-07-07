from dotenv import load_dotenv
import os
import time
from api_wsp.logging import listar_sesiones, crear_sesion, iniciar_sesion, verificar_estado, generar_qr, detener_sesion
from api_wsp.registrar_webhook import asegurar_webhook

load_dotenv()
API_KEY = os.getenv("API-KEY") 
URL = os.getenv("URL")
HEADERS = {"Content-Type": "application/json", "x-api-key": f"{API_KEY}"}
SESSION_NAME = "nabla-bot"

def obtener_o_crear_sesion() -> str | None:
    try:
        sesiones = listar_sesiones(URL, HEADERS)
        if isinstance(sesiones, list):
            for s in sesiones:
                if s.get("name") == SESSION_NAME: return s.get("id")
        nueva_sesion = crear_sesion(URL, HEADERS)
        return nueva_sesion.get("id") if isinstance(nueva_sesion, dict) else None
    except: return None

def main():
    sesion_id = obtener_o_crear_sesion()
    if not sesion_id: return
    
    print(f"✅ Sesion ID detectada: {sesion_id}. Esperando a la API...")
    
    # Bucle de espera inteligente
    for _ in range(60): # Intentar por 20 minutos
        estado = verificar_estado(URL, HEADERS, sesion_id)
        print(f"Estado actual: {estado}")
        
        if estado == "ready":
            print("✅ ¡Sesión lista!")
            asegurar_webhook(URL, HEADERS, sesion_id)
            break
        elif estado in ("disconnected", "created"):
            print("⏳ Iniciando la sesión...")
            iniciar_sesion(URL, HEADERS, sesion_id)
            time.sleep(10)
        elif estado == "initializing":
            print("⏳ La API está cargando el navegador... esperando 20 segundos.")
            time.sleep(20)
        elif estado == "qr_ready":
            print("📷 Escanea el siguiente código QR para conectar WhatsApp:")
            generar_qr(URL, HEADERS, sesion_id)
            time.sleep(10)
        else:
            time.sleep(10)

    print("\n🟢 Bot en línea. Esperando eventos...")
    while True: time.sleep(60)

if __name__ == "__main__":
    main()
