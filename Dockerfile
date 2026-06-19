# ══════════════════════════════════════════════════════════════════════════
# Dockerfile — conteo-inventario + wkhtmltopdf
# Render detecta este archivo y lo usa automáticamente en lugar del
# buildpack de Node. El resto del deploy (npm install, start) sigue igual.
# ══════════════════════════════════════════════════════════════════════════

# Base: Node LTS sobre Debian (necesario para apt-get)
FROM node:20-bookworm-slim

# Instalar wkhtmltopdf y sus dependencias de sistema
# wkhtmltopdf requiere libras de X11/Qt para renderizar correctamente.
# xvfb-run no es necesario porque usamos --headless implícitamente.
RUN apt-get update && apt-get install -y --no-install-recommends \
    wkhtmltopdf \
    libxrender1 \
    libxext6 \
    libx11-6 \
    libssl3 \
    ca-certificates \
    fontconfig \
    fonts-liberation \
 && rm -rf /var/lib/apt/lists/*

# Verificar que wkhtmltopdf quedó instalado y accesible
RUN wkhtmltopdf --version

# Directorio de trabajo
WORKDIR /app

# Copiar dependencias primero (cache de capas)
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

# Copiar el resto de la aplicación
COPY . .

# Puerto expuesto (Render lo asigna via $PORT)
EXPOSE 3000

# Comando de inicio — mismo que package.json pero respeta $PORT de Render
CMD ["node", "--max-old-space-size=1800", "server.js"]
