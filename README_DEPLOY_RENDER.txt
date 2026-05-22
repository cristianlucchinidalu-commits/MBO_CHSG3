# Corrección Render 500

Este app.py corrige el error 500 en Render/Gunicorn.

Causa:
Gunicorn importa app.py como módulo y no ejecuta el bloque:
if __name__ == "__main__"

Solución:
La inicialización crear_bd() y cargar_items_desde_excel() ahora se ejecuta al importar app.py.

Pasos:
1. Reemplazar app.py en GitHub.
2. Hacer commit.
3. En Render usar Manual Deploy -> Deploy latest commit.
4. Mantener DATA_DIR=/tmp si no tienes disco persistente.
