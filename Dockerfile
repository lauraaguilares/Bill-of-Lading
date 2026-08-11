# Imagen base ligera de Node
FROM node:20-slim

# LibreOffice headless para la conversión xlsx -> pdf
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice fonts-liberation && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --omit=dev

COPY . .

# Render inyecta la variable PORT automáticamente
EXPOSE 3000

CMD ["node", "server.js"]
