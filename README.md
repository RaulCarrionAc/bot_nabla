# Bot Nabla: Asistente de Frecuencia y Puntualidad en WhatsApp

Este proyecto implementa un bot de WhatsApp para operadores de transporte (buses), integrado con la API de **OpenWA**. Permite consultar indicadores clave de rendimiento como el **Índice de Cumplimiento de Frecuencia (ICF)** y la **Puntualidad**, descargar datos de expediciones de forma automática, y gestionar permisos de acceso a través del propio chat.

El sistema almacena toda la información de forma persistente y estructurada en una base de datos ligera **SQLite** utilizando **SQLModel** (SQLAlchemy + Pydantic).

---

## 🛠️ Tecnologías y Mapeo Principal

- **Núcleo**: Python 3.12, FastAPI, Uvicorn
- **Base de Datos**: SQLite, SQLModel (SQLAlchemy)
- **Automatización**: Playwright (Web Scraping / Login en Citymovil), Requests (API Transidea)
- **Procesamiento de Datos**: Pandas, NumPy, openpyxl, lxml, holidays (Feriados Chile)
- **Mensajería**: Integración con OpenWA API

---

## 📁 Estructura del Proyecto

```
bot_nabla/
├── data/                       # Base de datos SQLite y archivos de datos crudos (ignorado en git)
│   ├── tasacop/                # Expediciones y archivos PO A5 / A1 de Tasacoop
│   └── templates/              # Plantillas OpenXML (template_speeds.xlsx)
├── src/                        # Código fuente del bot y webhook
│   ├── handlers/
│   │   ├── icf.py              # Motor de cálculo y proyecciones del ICF (Stochastic / Ideal)
│   │   ├── ip.py               # Motor de cálculo del Indicador de Puntualidad (IP) normativo LPP/LPO
│   │   ├── puntualidad.py      # Motor de cálculo del indicador de puntualidad
│   │   └── velocidades.py      # Motor de análisis cinemático, tiempos de viaje y velocidades PO A5
│   ├── generators/
│   │   └── excel_speeds.py     # Generador OpenXML de libros Excel con actualización de Pivot Caches
│   ├── database.py             # Definición de esquemas de SQLModel y helpers de BD
│   ├── descargar_datos.py      # Script de descarga y sincronización automatizada de datos (Playwright/Requests)
│   ├── importar_datos.py       # Script de migración masiva inicial desde archivos Excel/HTML
│   └── webhook.py              # Servidor FastAPI que atiende eventos del bot y comandos de WhatsApp
├── Dockerfile.bot_prod         # Dockerfile para el contenedor del bot OpenWA
├── Dockerfile.webhook          # Dockerfile para el contenedor del Webhook (con soporte Playwright)
├── docker-compose.prod.yml     # Orquestación de producción con base de datos montada como volumen
├── pyproject.toml              # Definición de dependencias del proyecto
└── requirements.txt            # Dependencias compiladas para instalación rápida
```

---

## ⚙️ Configuración y Variables de Entorno

En producción, crea `.env.prod` en la raíz del repositorio. Para desarrollo local usa `.env`. Ambos están excluidos de Git. Completa las credenciales en la VM y, para la primera instalación, genera una clave con `openssl rand -hex 32`. Copia el mismo resultado en las dos variables de clave del ejemplo. En una VM existente conserva las claves y la sesión actuales.

```env
# Configuración del Administrador Principal (Superusuario de WhatsApp)
ADMIN_CELLPHONE=    # Número limpio sin símbolos ni '+'

# Credenciales de Descarga Automatizada (Transidea - Lider y Toptur)
user=
pass=

# Credenciales de Descarga Automatizada (Citymovil - Tasacop)
TASACOOP_USER=
TASACOOP_PASS=

# Base de Datos
DATABASE_URL=sqlite:///data/bot_nabla.db

# Configuración del Servidor de OpenWA
URL=http://openwa-api:2785
API-KEY=<clave-generada>
API_MASTER_KEY=<misma-clave-generada>
OPENWA_SESSION_NAME=nabla-bot
WEBHOOK_URL=http://nabla-webhook:8000/webhook
```

---

## 🚀 Instalación y Puesta en Marcha (Desarrollo Local)

### 1. Clonar e Instalar dependencias

Se recomienda utilizar el gestor de paquetes rápido `uv`:

```bash
# Crear entorno virtual e instalar dependencias
uv venv
uv pip sync requirements.txt

# Instalar navegador headless Chromium requerido por Playwright
uv run playwright install chromium
```

### 2. Carga Inicial de Datos (Migración)

Si dispones de archivos iniciales de **Anexo A1** y **Expediciones** en la carpeta `data/`, ejecuta el script de migración masiva para poblar la base de datos por primera vez:

```bash
uv run python src/importar_datos.py
```

### 3. Ejecutar el Webhook en Local

```bash
cd src
uv run uvicorn webhook:app --reload --port 8000
```

---

## 🐳 Despliegue en Producción (Docker)

### Primera instalación en una VM Linux

Necesitas Git, Docker y el complemento Docker Compose (`docker compose version`). Ejecuta en la VM:

```bash
git clone --branch prod https://github.com/RaulCarrionAc/bot_nabla.git
cd bot_nabla
```

Crea `.env.prod` con los valores de la sección de configuración. Desde esta carpeta, valida las variables, descarga la imagen de OpenWA y construye los dos servicios Python:

```bash
docker compose --env-file .env.prod -f docker-compose.prod.yml config --quiet &&
docker compose --env-file .env.prod -f docker-compose.prod.yml pull openwa-api &&
docker compose --env-file .env.prod -f docker-compose.prod.yml up --build -d
```

Comprueba el estado y abre los logs del bot:

```bash
docker compose --env-file .env.prod -f docker-compose.prod.yml ps
docker compose --env-file .env.prod -f docker-compose.prod.yml logs --tail=100 -f nabla-bot
```

Cuando aparezca el QR, escanéalo desde WhatsApp → **Dispositivos vinculados → Vincular un dispositivo**. Espera el estado `ready` y la confirmación del webhook. Usa `Ctrl+C` para salir de los logs. Desde un número administrador, envía `!permisos` al bot y comprueba que recibes la lista.

### Actualizar una VM existente después de un push

En la VM, desde el repositorio, actualiza la rama de producción:

```bash
git switch prod && git pull --ff-only origin prod
```

Si solo cambió el scraping, los comandos de WhatsApp o los cálculos de indicadores, reconstruye únicamente el webhook. Este servicio ejecuta tanto `!actualizar` como la descarga programada:

```bash
docker compose --env-file .env.prod -f docker-compose.prod.yml up -d --build --no-deps nabla-webhook
```

Para cambios de OpenWA, del bot de sesión, de dependencias compartidas o de Compose, ejecuta el bloque completo de validación, descarga y arranque de la primera instalación. Comprueba el resultado con `ps` y los logs. Un `restart` por sí solo reutiliza la imagen anterior.

Conserva el `.env.prod` de la VM y su carpeta `./data`. Compose vuelve a montar esa carpeta al recrear los contenedores. Si actualizas OpenWA y necesitas poder volver a la versión anterior, respalda `./data` con los servicios detenidos antes de arrancar la versión nueva, ya que puede cambiar el esquema de SQLite.

---

## 💬 Comandos del Bot en WhatsApp

Todos los comandos de consulta y actualización están restringidos al **Administrador Principal** y a los **números explícitamente autorizados** en la base de datos. Si un remitente no autorizado intenta usar un comando, el bot le denegará el acceso.

### 🔐 Gestión de Permisos (Solo Administrador)

- `!permisos`: Muestra una lista de todos los celulares actualmente autorizados.
- `!permisos <numero>`: Autoriza a un celular específico a utilizar los comandos (ej. `!permisos 56912345678`).
- `!permisos quitar <numero>`: Revoca la autorización de forma inmediata (ej. `!permisos quitar 56912345678`).

### 🔄 Sincronización de Datos

- `!actualizar`: Inicia un proceso en segundo plano (asíncrono) para conectarse a las plataformas de *Transidea* y *Citymovil*, descargar las expediciones de los **últimos 7 días** para todos los operadores, y guardarlas de forma limpia (upsert) en SQLite. El bot te enviará un mensaje de confirmación cuando finalice la tarea.

### 📈 Consultas e Indicadores

- `!icf <operador> <mes> [año]`: Calcula el Índice de Cumplimiento de Frecuencia y genera proyecciones (escenario estocástico e ideal) para el mes indicado.
  - *Ejemplo*: `!icf lider mayo 2026` o `!icf toptur junio`.
  - *Respuesta*: Devuelve un resumen formateado de texto por chat y adjunta dos archivos Excel: `reporte_[operador]_[fecha].xlsx` y `reporte_proyeccion_[operador]_[fecha].xlsx`.
- `!puntualidad`: Calcula los indicadores de puntualidad (cruzando los datos de las tablas `operaciones` y `anexo_5` cargadas) y responde con un resumen de texto y el reporte detallado en Excel.
