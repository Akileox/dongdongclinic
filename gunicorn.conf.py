# Gunicorn configuration for Render.com (512MB RAM optimization)
import multiprocessing

# Workers: 2 workers (Reduced from default to balance parallel load and 512MB RAM limit)
# The user requested 'half' of original parallel processing.
workers = 2

# Threads: Use threads for slightly better concurrency within the single worker
threads = 4

# Timeout: Set to 10 minutes (600 seconds) to handle large batch processing (100+ reports)
# Default 30s is too short for long-running image generation tasks
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
