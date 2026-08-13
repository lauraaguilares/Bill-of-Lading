# Imagen base ligera de Node
FROM node:20-slim

# LibreOffice headless para la conversión xlsx -> pdf
# fonts-crosextra-carlito: sustituto de Calibri con las MISMAS métricas (ancho de letras),
# imprescindible porque la plantilla del BOL usa Calibri y sin esto LibreOffice cae a una
# fuente genérica que desalinea y corta texto (Calibri en sí no se puede redistribuir).
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice fonts-liberation fonts-crosextra-carlito && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --omit=dev

COPY . .

# Render inyecta la variable PORT automáticamente
EXPOSE 3000

CMD ["node", "--dns-result-order=ipv4first", "server.js"]
