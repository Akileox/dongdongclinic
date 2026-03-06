# Use official Python runtime as a parent image
FROM python:3.12-slim

# Install necessary system dependencies for Playwright
RUN apt-get update && apt-get install -y \
    libnss3 \
    libnspr4 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdrm2 \
    libxkbcommon0 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxrandr2 \
    libgbm1 \
    libasound2 \
    libpango-1.0-0 \
    libcairo2 \
    && rm -rf /var/lib/apt/lists/*

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file into the container
COPY requirements.txt .

# Install dependencies and explicitly install gunicorn for production serving
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install --no-cache-dir gunicorn

# Install Playwright browsers (chromium only to save space)
RUN playwright install chromium

# Copy the rest of the application code
COPY . .

# Expose port (Render automatically uses PORT env var)
EXPOSE 5000

# Run the application using gunicorn
CMD gunicorn -b 0.0.0.0:$PORT app:app
