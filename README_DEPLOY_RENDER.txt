MBO_CHSG3 - Deploy Render

Archivos principales:
- app.py
- requirements.txt
- runtime.txt
- Procfile

Start Command recomendado en Render:
gunicorn app:app --workers 1 --timeout 180 --bind 0.0.0.0:$PORT

Notas:
1. Subir el archivo principal con nombre exacto app.py.
2. Mantener el Excel MBO en el repositorio o junto al app.py si se ejecuta localmente.
3. En Render usar Manual Deploy -> Clear build cache & deploy si no ves cambios.
4. Evitar trabajar dentro de OneDrive si GitHub Desktop no detecta cambios.
