# ============================================================
#  Activa Previsora SOAT - README para Railway
# ============================================================

## Estructura del Proyecto

```
activa_previsora/
├── railway.json      # Configuración de Railway
├── Dockerfile       # Imagen de Docker con Playwright
├── requirements.txt # Dependencias Python
├── build.sh         # Script de build (opcional)
├── app.py           # Aplicación Flask principal
├── iniciar.sh       # Script de inicio local (ignorar en Railway)
├── iniciar.bat      # Script de inicio Windows local
└── templates/
    └── index.html   # Interfaz web
```

## Configuración en Railway

### Método 1: Deploy desde GitHub (Recomendado)

1. Sube este proyecto a un repositorio GitHub
2. Ve a [Railway Dashboard](https://railway.app)
3. Crea un nuevo proyecto → "Deploy from GitHub repo"
4. Selecciona el repositorio
5. Railway detectará automáticamente el Dockerfile

### Método 2: Deploy con CLI

```bash
# Instalar Railway CLI
npm install -g @railway/cli

# Login
railway login

# Ir al directorio del proyecto
cd activa_previsora

# Inicializar proyecto (si no existe ya)
railway init

# Deploy
railway up
```

### Método 3: Deploy directo desde ZIP

1. Descarga este folder completo como ZIP
2. Ve a Railway → New Project → "Deploy from local source"
3. Sube el ZIP

## Configuración de Variables de Entorno

Railway detectará automáticamente:
- Puerto: `5050` (definido en código)
- Build command: automático desde Dockerfile

Opcionalmente puedes configurar:
- `PORT` (default: 5050)
- `FLASK_ENV` (production/development)

## Verificación del Deploy

Después de desplegar, la URL será algo como:
`https://activa-previsora.up.railway.app`

La aplicación estará disponible en `/` (página principal)

## Notas Importantes

1. **Playwright + Chromium**: El Dockerfile instala Chromium automáticamente
2. **Descargas**: Los PDFs se guardan en `/app/downloads_5050` (persistido en Railway)
3. **Logs**: Los logs de ejecución están disponibles en el dashboard de Railway

## Solución de Problemas

### Error: Chromium no está instalado
```bash
# En el contenedor, ejecuta:
playwright install chromium
```

### Error: Permisos en directorio de descargas
```bash
# Asegúrate que existe y tiene permisos:
mkdir -p /app/downloads_5050
chmod 777 /app/downloads_5050
```

### Timeout en inicialización
Railway tiene límite de 120s para startup. Si tu app tarda más:
- Reduce Tiempo de espera en railway.json
- Optimiza el proceso de instalación de Playwright

---

** Autor**: Sistema Activa IT
** Versión**: 1.0
** Fecha**: 2026
