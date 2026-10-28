# Reunion 123

Aplicación web para trabajar con el formato Excel y homologaciones de productos.

## Ejecución local

Requiere Node.js 20 o superior.

```bash
npm start
```

La aplicación queda disponible en `http://localhost:3070`.

## Persistencia

Si existe `DATABASE_URL`, el servidor usa Neon PostgreSQL y crea automáticamente las tablas `reunion123_workbook_state` y `reunion123_homologations` al iniciar. Sin esa variable, conserva el modo local con archivos JSON.

Para desarrollo local, copia `.env.example` como `.env` y agrega la cadena privada de Neon. Nunca subas `.env` al repositorio.

## Despliegue en Render

- Build command: `npm install`
- Start command: `npm start`

La aplicación usa la variable de entorno `PORT` que Render proporciona automáticamente.

> Actualmente los cambios se guardan en archivos JSON. El almacenamiento en disco de Render no es persistente; la migración a una base de datos se hará en la siguiente etapa.

## Actualizar Data local

El botón **Actualizar Data** ejecuta un proceso independiente del repositorio `BOT-123`:

1. Entra a POSCO con las credenciales locales.
2. Descarga el reporte de órdenes activas usando el mismo rango y filtros del bot Pegar Data.
3. Reordena y normaliza las columnas del reporte.
4. Actualiza únicamente la hoja web `PegarData`, incluyendo `Table2` y las fórmulas `S:W`.

Para habilitarlo localmente, copia `.env.example` a `.env` y completa:

```env
POSCO_USER=usuario@email.com
POSCO_PASSWORD=tu_contraseña
# Optional locally; Render uses the packaged Chromium binary when this is omitted.
# POSCO_BROWSER_CHANNEL=msedge
```

Después inicia el sistema con `npm start` y usa el botón **Actualizar Data**. En Render, Chromium viene empaquetado en la aplicación; localmente se puede usar Edge con `POSCO_BROWSER_CHANNEL=msedge` o indicar una ruta directa con `POSCO_BROWSER_EXECUTABLE_PATH`.

## Actualizar Data e Inventario con GitHub Actions

El workflow `.github/workflows/posco-updates.yml` ejecuta ambos procesos en un runner de GitHub y escribe los resultados en Neon. Para habilitar los botones desde Render, configura estas variables en Render:

```env
GITHUB_ACTIONS_TOKEN=token_con_permiso_para_disparar_workflows
REUNION_SYNC_TOKEN=token_privado_compartido_con_GitHub
GITHUB_REPOSITORY=0322203936-cmd/Reunion-123-WEB
```

Y estos secretos en GitHub Actions:

```text
DATABASE_URL
POSCO_USER
POSCO_PASSWORD
REUNION_CALLBACK_URL=https://reunion-123.onrender.com/api/posco-callback
REUNION_SYNC_TOKEN   (el mismo valor configurado en Render)
```

El botón **Actualizar Data** dispara la operación `data`; **Actualizar Inventario** dispara `inventario`. Cuando termina, GitHub avisa a Render, se limpian las cachés y las pestañas cargan los datos nuevos desde Neon.
