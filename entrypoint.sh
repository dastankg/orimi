#!/bin/sh


while ! nc -z db 5432; do
  sleep 0.2
done

python manage.py makemigrations
python manage.py migrate
python manage.py collectstatic --noinput

if [ -n "$PROMETHEUS_MULTIPROC_DIR" ]; then
  rm -rf "$PROMETHEUS_MULTIPROC_DIR"
  mkdir -p "$PROMETHEUS_MULTIPROC_DIR"
fi

gunicorn configs.wsgi:application \
  --config gunicorn.conf.py \
  --bind 0.0.0.0:8000 \
  --workers 4 \
  --timeout 120 \
  --access-logfile - \
  --error-logfile - \
  --log-level info \
  --capture-output