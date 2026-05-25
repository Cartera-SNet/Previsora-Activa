# syntax=docker/dockerfile:1
FROM python:3.11-slim

# Instalar dependencias del sistema para Playwright y Chromium
RUN apt-get update && apt-get install -y \
    # Playwright dependencies
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
    # Fonts for PDF rendering
    fonts-liberation \
    libfontconfig1 \
    libfreetype6 \
    # Utilities
    wget \
    gnupg \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Establecer variables de entorno
ENV PYTHONUNBUFFERED=1 \
    PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=0 \
    PLAYWRIGHT_CHROMIUM_CHANNEL=chromium

# Configurar directorio de trabajo
WORKDIR /app

# Copiar archivos de dependencias primero (para caché de Docker)
COPY requirements.txt .

# Instalar dependencias de Python
RUN pip install --no-cache-dir -r requirements.txt

# Instalar Playwright browsers (Chromium solo)
RUN playwright install --with-deps chromium

# Copiar resto de la aplicación
COPY . .

# Crear directorio de descargas dinámico
RUN mkdir -p /app/downloads_5050

# Puerto de la aplicación (Previsora usa puerto 5050 por defecto)
EXPOSE 5050

# Comando de inicio
CMD ["python", "app.py"]
