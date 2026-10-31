# Validación local de OpenWA

Fecha: 2026-10-30. Versión: 0.26.0.

GHCR rechazó la descarga con `denied`. Se verificó en el workflow oficial
de la etiqueta v0.26.0 que Docker Hub también es un destino de publicación.
Se descargó y arrancó `docker.io/rmyndharis/openwa:0.26.0` correctamente.
Digest: `sha256:46e3d3d01b1948c96a0bea94f6650aa772370f40dd3956052d2da24fe4ee2acf`.

Pruebas realizadas sobre un contenedor aislado, con SQLite y clave nuevas:

- Health HTTP 200 y rechazo HTTP 401 sin autenticación.
- Listado y creación de sesiones con `src/api_wsp/logging.py`.
- Registro y reutilización del mismo webhook con `src/api_wsp/registrar_webhook.py`.
- Endpoints send-text y send-document: HTTP 400 indicando sesión inactiva.
  Esto confirma acceso a las rutas, pero no la entrega ni toda la validación del payload.
- Configuración Compose de producción válida con `--env-file .env.prod`.
- Las claves locales del bot y API coinciden y cumplen el mínimo de 32 caracteres.

El contenedor `nabla-openwa-smoke` quedó detenido. No se modificaron datos
ni sesiones existentes. Pendiente: vincular una cuenta de prueba para comprobar
QR, entrega de mensajes/documentos y recepción real de eventos en el webhook.
