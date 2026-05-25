# syntax=docker/dockerfile:1
FROM python:3.11-slim

# Dependencias del sistema para Chromium/Playwright
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
    libx11-6 \
    libxcb1 \
    libxext6 \
    fonts-liberation \
    fonts-unifont \
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

RUN pip install --no-cache-dir -r requirements.txt

# Sin --with-deps porque ya instalamos las dependencias arriba manualmente
RUN python -m playwright install chromium

COPY . .

RUN mkdir -p /app/downloads_8080

EXPOSE 8080

CMD ["python", "app.py"]
