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
