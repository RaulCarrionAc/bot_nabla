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

Crea un archivo `.env` en la raíz del proyecto (basándote en `.env.prod`) con la siguiente estructura:

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
API-KEY=<genera-una-key-segura>
API_MASTER_KEY=dev-admin-key
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

El proyecto está configurado para desplegarse mediante Docker Compose. El webhook se basa en una imagen Debian Slim a la que se le inyecta Chromium y las librerías de sistema de Playwright.

Para desplegar los servicios en segundo plano:

```bash
docker-compose -f docker-compose.prod.yml up --build -d
```

> ⚠️ **Importante**: La base de datos se almacena en el volumen local `./data` para evitar pérdidas de información al reiniciar o recrear los contenedores.

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
