# syntax=docker/dockerfile:1
FROM python:3.11-slim

# Instalar dependencias del sistema para Playwright y Chromium
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
    libpango-1.0-0 \
    libcairo2 \
    libasound2 \
    libatspi2.0-0 \
    libxshmfence1 \
    fonts-liberation \
    libfontconfig1 \
    libfreetype6 \
    wget \
    gnupg \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

ENV PYTHONUNBUFFERED=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

WORKDIR /app

COPY requirements.txt .

# Instalar dependencias Python
RUN pip install --no-cache-dir -r requirements.txt

# Instalar Chromium en ruta fija (evita conflictos de versión en cache)
RUN python -m playwright install --with-deps chromium

COPY . .

RUN mkdir -p /app/downloads_8080

EXPOSE 8080

CMD ["python", "app.py"]
