# gunicorn.conf.py
workers = 1
worker_class = "sync"
timeout = 300          # 5 minutes max
keepalive = 5
max_requests = 1000
max_requests_jitter = 50
preload_app = True     # Load app before workers