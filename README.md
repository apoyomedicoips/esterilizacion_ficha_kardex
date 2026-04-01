# Esterilizacion Ficha Kardex (Web App)

## Implementacion recomendada: Google Apps Script

La version solicitada con Apps Script esta en:

- `apps_script/`

Guia directa:

- [apps_script/README.md](apps_script/README.md)

Resumen:

1. Crear proyecto web app con `clasp`.
2. Subir codigo de `apps_script/src`.
3. Ejecutar `setupDefaultAdmin`.
4. Desplegar como Web App.
5. Inicializar hojas con boton `Init hojas`.

---
Aplicacion web segura para reemplazar planillas online en el proceso de kardex de esterilizacion.

Repositorio: `https://github.com/apoyomedicoips/esterilizacion_ficha_kardex.git`

## Seguridad incluida

- Login con usuarios y roles (`admin`, `operator`, `viewer`).
- Passwords con hash (`werkzeug.security`).
- Formularios con token CSRF.
- Auditoria inmutable de acciones en hoja `AUDIT_LOG`.
- Movimientos inmutables en hoja `MOVEMENTS`.

## Persistencia de datos (obligatoria)

Todos los datos se leen y escriben en:

- `https://docs.google.com/spreadsheets/d/161uPwejp2D8YVO7OVgbVyqeUiqzt0fYBzfZG5ERmLLg/edit?usp=sharing`

La app crea/usa estas hojas:

- `ITEMS`
- `SERVICES`
- `MOVEMENTS`
- `AUDIT_LOG`

## 1) Preparar credenciales Google

1. Crear un Service Account en Google Cloud.
2. Descargar credencial JSON.
3. Compartir la planilla con el email del Service Account como `Editor`.
4. Guardar JSON en `service-account.json` (no subir a git).

## 2) Configurar entorno

Copiar `.env.example` a `.env` y completar:

- `SECRET_KEY`
- `SPREADSHEET_ID` (ya viene con el ID solicitado)
- `GOOGLE_SERVICE_ACCOUNT_FILE=./service-account.json`
- `KARDEX_USERS_JSON` con hashes de password

Generar hash:

```bash
python scripts/hash_password.py
```

## 3) Ejecutar local

```bash
python -m venv .venv
.venv\\Scripts\\activate
pip install -r requirements.txt
python app.py
```

Abrir: `http://127.0.0.1:5050`

## 4) Inicializar esquema de hojas

Ingresar con usuario admin y en Dashboard usar boton `Inicializar hojas`.

## 5) Deploy

Incluye `Procfile` para plataformas compatibles con `gunicorn`.

## Usuario demo

Si no defines `KARDEX_USERS_JSON`, usa:

- usuario: `admin`
- password: `Cambiar123!`

Cambiarlo inmediatamente en produccion.

