# Google Apps Script - Kardex Seguro

Esta carpeta contiene la version web usando Google Apps Script y Google Sheets.

Planilla de datos:
- https://docs.google.com/spreadsheets/d/161uPwejp2D8YVO7OVgbVyqeUiqzt0fYBzfZG5ERmLLg/edit?usp=sharing

## Estructura
- `src/Code.gs`: backend (auth, roles, auditoria, CRUD y dashboard)
- `src/index.html`: frontend web app
- `src/appsscript.json`: manifest

## Seguridad incluida
- Login propio con hash SHA-256 + salt + pepper
- Roles: `admin`, `operator`, `viewer`
- Sesion con expiracion (8 horas) en `CacheService`
- Auditoria en hoja `AUDIT_LOG`
- Datos en hojas separadas: `ITEMS`, `SERVICES`, `MOVEMENTS`

## Publicacion con clasp

1. Instalar herramientas
```bash
npm i -g @google/clasp
clasp login
```

2. Crear proyecto Apps Script (una sola vez)
```bash
cd apps_script
clasp create --type webapp --title "Kardex Seguro Esterilizacion"
```

3. Subir codigo
```bash
clasp push
```

4. Inicializar usuario basico desde editor de Apps Script
- Abrir Apps Script
- Ejecutar funcion: `setupDefaultAdmin`
- Crea: `user / 123` (rol `admin`)

5. Desplegar web app
- Deploy > New deployment > Web app
- Execute as: `User accessing the web app`
- Who has access: segun politica (ideal: usuarios del dominio)

6. Primera configuracion en la web
- Iniciar sesion con `user`
- Boton `Init hojas` para crear estructura base

## Gestion de usuarios

En Apps Script, ejecutar manualmente:
```javascript
upsertUser('operador1', 'ClaveSegura123!', 'operator', true)
upsertUser('visual1', 'ClaveSegura123!', 'viewer', true)
upsertUser('admin2', 'ClaveSegura123!', 'admin', true)
```

## Recomendaciones
- Cambiar la clave de `user` inmediatamente.
- Limitar acceso del deploy a usuarios autenticados del dominio.
- No compartir el editor de Apps Script con usuarios no admin.
