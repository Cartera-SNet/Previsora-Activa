#!/bin/bash
# ============================================================
#  Activa Previsora - Script de Build para Railway
#  Este script se ejecuta durante el build del contenedor
# ============================================================

set -e

echo "============================================================"
echo "  Activa Previsora SOAT - Build para Railway"
echo "============================================================"
echo ""

# El build se maneja vía Dockerfile
# Este script留着 para documentación o pasos adicionales si son necesarios

echo "Instalando dependencias de Python..."
pip install --no-cache-dir -r requirements.txt

echo "Instalando navegador Chromium..."
playwright install --with-deps chromium

echo ""
echo "Build completado exitosamente!"
echo "La aplicación estará disponible en el puerto 5050"
