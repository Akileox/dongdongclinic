# Gunicorn configuration for Render.com (512MB RAM optimization)
import multiprocessing

# Workers: 2 workers recommended for 1 CPU / 2GB RAM to balance context switching
workers = 2

# Threads: 4 threads per worker
threads = 4

# Timeout: Set to 1 hour (3600 seconds) for massive batches
timeout = 3600

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
