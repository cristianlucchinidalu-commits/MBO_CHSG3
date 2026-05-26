MBO_CHSG3 corregido.

1) Reemplaza estos archivos en tu repositorio.
2) En GitHub Desktop: Commit to main y Push origin.
3) En Render: Manual Deploy -> Clear build cache & deploy.

Start Command recomendado:
gunicorn app:app --workers 1 --timeout 180 --bind 0.0.0.0:$PORT
