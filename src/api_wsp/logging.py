import requests
from pprint import pprint
import base64, subprocess

def listar_sesiones(url:str, headers:dict)->list:
    """Listar sesiones"""
    r = requests.get(f"{url}/api/sessions", headers = headers)
    return r.json()

def crear_sesion(url:str, headers:dict):
    """Crear sesion en api wsp"""
    payload = {"name": "nabla-bot"}
    try:
        r = requests.post(f"{url}/api/sessions", json=payload, headers=headers)
        if r.status_code in (200, 201):
            return r.json()
        else:
            print(f"Error en la petición. Código de estado: {r.status_code}")
    except Exception as e:
        print(f"La solicitud de creacion de sesion fallo: {e}")

def iniciar_sesion(url:str, headers:dict, sesion_id:str):
    try:
        r = requests.post(f"{url}/api/sessions/{sesion_id}/start", headers=headers)
        if r.status_code in (200, 201):
            print("Sesion iniciada.")
        else:
            print(f"No se pudo iniciar la sesion. Status: {r.status_code}")
        return r.json()
    except Exception as e:
        print(f"Error: {e}")

def verificar_estado(url:str, headers:dict, sesion_id:str) -> str | None:
    """Retorna el status de la sesion como string, o None si falla"""
    try:
        r = requests.get(f"{url}/api/sessions/{sesion_id}", headers=headers)
        if r.status_code in (200, 201):
            salida = r.json()
            pprint(salida)
            return salida.get("status")
        return None
    except Exception as e:
        print(f"No se pudo verificar el estado: {e}")
        return None
    
def generar_qr(url:str, headers:dict, sesion_id:str) -> bool:
    """Genera y muestra QR en terminal."""
    try:
        r = requests.get(f"{url}/api/sessions/{sesion_id}/qr", headers=headers)
        if r.status_code in (200, 201):
            data = r.json().get("qrCode", "")
            if not data:
                print("QR no disponible aún.")
                return False
            
            # Decodificar el contenido del QR desde la imagen
            b64 = data.split(",")[1]
            img_bytes = base64.b64decode(b64)
            
            # Leer el contenido del QR con pyzbar
            from PIL import Image
            from pyzbar.pyzbar import decode
            import io
            
            img = Image.open(io.BytesIO(img_bytes))
            decoded = decode(img)
            
            if not decoded:
                print("No se pudo decodificar el QR.")
                return False
            
            qr_content = decoded[0].data.decode("utf-8")
            
            # Redibujar el QR en terminal con qrcode
            import qrcode
            qr = qrcode.QRCode(border = 1)
            qr.add_data(qr_content)
            qr.make()
            qr.print_ascii(invert=True)
            return True
        else:
            print(f"No se pudo obtener el QR. Status: {r.status_code}")
            return False
    except Exception as e:
        print(f"Error generando QR: {e}")
        return False

def detener_sesion(url:str, headers:dict, sesion_id:str) -> bool:
    try:
        r = requests.post(f"{url}/api/sessions/{sesion_id}/stop", headers=headers)
        if r.status_code in (200, 201):
            print("Sesion detenida.")
            return True
        print(f"No se pudo detener la sesion. Status: {r.status_code}")
        return False
    except Exception as e:
        print(f"Error deteniendo sesion: {e}")
        return False