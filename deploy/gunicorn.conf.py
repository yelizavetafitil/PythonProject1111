# Gunicorn settings — workers=1 рекомендуется при SQLite (одна запись в БД).
import os

bind = os.environ.get("GUNICORN_BIND", "127.0.0.1:5004")
workers = int(os.environ.get("GUNICORN_WORKERS", "1"))
threads = int(os.environ.get("GUNICORN_THREADS", "4"))
timeout = int(os.environ.get("GUNICORN_TIMEOUT", "120"))
worker_class = "gthread"
accesslog = "-"
errorlog = "-"
capture_output = True
