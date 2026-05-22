# MBO CHSG3 - App Flask para nube

Este paquete está preparado para desplegar el aplicativo MBO en Render.

## Archivos incluidos

- app.py: aplicativo Flask completo.
- requirements.txt: librerías Python.
- Procfile: comando de arranque con Gunicorn.
- .gitignore: evita subir base SQLite y carpetas temporales.

## Archivo Excel base requerido

Debes agregar al repositorio privado el Excel base con este nombre exacto:

HGP-SG3-OP-FR-033 Mantenimiento Básico Operacional (MBO)_May_26.xlsx

El app lo busca junto a app.py y también en DATA_DIR.

## Variables de entorno recomendadas en Render

SECRET_KEY = una_clave_larga
APP_PASSWORD = clave_para_entrar_al_app
DATA_DIR = /var/data

## Disco persistente recomendado

En Render crea un Persistent Disk montado en:

/var/data

Ahí se guardarán:

- mbo.db
- uploads/
- exports/

## Build Command

pip install -r requirements.txt

## Start Command

gunicorn app:app --workers 1 --timeout 180

## Uso

Una vez desplegado, Render te entregará una URL.
Desde el celular abres esa URL, ingresas la clave, registras datos y descargas el Excel desde los botones de exportación.
