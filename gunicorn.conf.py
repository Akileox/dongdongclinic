# Gunicorn configuration for Render.com (512MB RAM optimization)
import multiprocessing

# Workers: 2 workers for balanced throughput on Render Free/Starter Tier
workers = 2

# Threads: 4 threads per worker for I/O efficiency
threads = 4

# Timeout: Set to 10 minutes (600 seconds)
timeout = 600

# Worker class: gthread is suitable for I/O bound tasks like image generation
worker_class = 'gthread'

# Bind to the PORT environment variable provided by Render
import os
port = os.environ.get('PORT', '5000')
bind = f"0.0.0.0:{port}"

# Logging
accesslog = "-"
errorlog = "-"
loglevel = "info"
