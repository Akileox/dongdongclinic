# Gunicorn configuration for Render.com (512MB RAM optimization)
import multiprocessing

# Workers: Reduced to 1 to stay within Render Free Tier memory limits (512MB)
workers = 1

# Threads: Reduced to 2 to minimize memory overhead
threads = 2

# Timeout: Set to 1 hour (3600 seconds) for massive batches (100-200 reports)
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
