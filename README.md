# BMM Document Generator

Backend único para generar documentos automáticamente a partir de datos (empieza con el
Bill of Lading; se agregarán más tipos de documento con el mismo patrón).

## Estructura

```
config/        reglas de negocio por tipo de documento (bolConfig.js, futuros...)
generators/    lógica que llena la plantilla (generateBOL.js, readShipmentFromSource.js, futuros...)
templates/     plantillas .xlsx base
public/        formulario web (servido directo por Express, sin build step)
server.js      Express, expone POST /generate/{tipo} y sirve el formulario en /
Dockerfile     imagen con Node + LibreOffice (para la conversión a PDF)
```

## Cómo agregar un nuevo tipo de documento

1. Crea `config/miDocumentoConfig.js` con sus reglas de negocio.
2. Crea `generators/generateMiDocumento.js` con su función `generate...`.
3. Sube su plantilla `.xlsx`/`.docx` a `templates/`.
4. Agrega la ruta en `server.js`: `app.post('/generate/mi-documento', ...)`.
5. Si aplica, agrega su selector al formulario en `public/index.html`.

## Reglas de negocio del BOL (referencia rápida)

- **Direcciones**: siempre "Address for truck" (NOB Address / SOB Address), nunca direcciones de Customs.
  - Southbound: Shipper = NOB Address · Consignee = Customs Broker
  - Northbound: Shipper = Customs Broker · Consignee = NOB Address
  - Bonded: Shipper = NOB Address · Consignee = SOB Address
- **Fechas**: el campo "Date:" y el de "Special Instructions" usan columnas distintas según el escenario:
  - Southbound / Bonded: Date = NOB Day for Truck to Arrive · Special Instructions = CUSTOMS Day
  - Northbound: Date = CUSTOMS Day for Truck to Arrive · Special Instructions = NOB Day
- **Order #**: normalmente `BMM-{INICIALES}{FECHA}-01`. Excepción: si el carrier es ESTES u Old Dominion, se usa el Order Number tal cual viene en el archivo de origen (columna G de "PASTE HERE"). El PO # siempre mantiene el formato generado.
- **Escenario**: viene de la columna TAGS ('Northbound' / 'Bonded' / 'LIVE' o vacío = Southbound).
- **Contacto dinámico**: Moises, luego el PMA del embarque, luego Hannia (si el PMA no es ya Hannia).
- **Peso**: `cubic ft × 6.5`, excepto si el carrier es "Ground Freight Solutions" (manual).

## Correr localmente

```bash
npm install
node server.js
```

Abre `http://localhost:3000` para usar el formulario web (sube el archivo maestro, elige el
cliente, descarga el BOL). El archivo maestro nunca se guarda permanentemente en el servidor.

## Desplegar — paso a paso

### 1. Subir a GitHub

```bash
git init
git add .
git commit -m "BMM Document Generator - BOL"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/bmm-document-generator.git
git push -u origin main
```

(Si no tienes el repo creado todavía, créalo primero vacío en github.com/new — sin README,
sin .gitignore, para que no choque con este que ya trae uno.)

### 2. Desplegar en Render.com (gratis)

1. Entra a https://render.com y crea una cuenta (puedes usar tu cuenta de GitHub para entrar,
   no pide tarjeta).
2. Click en **New +** → **Web Service**.
3. Conecta tu repo `bmm-document-generator`.
4. Render detecta el `Dockerfile` automáticamente — déjalo en **Environment: Docker**.
5. Plan: **Free**.
6. Click **Create Web Service**. El primer build tarda unos minutos porque instala LibreOffice.
7. Cuando termine, Render te da una URL pública tipo
   `https://bmm-document-generator.onrender.com` — ahí mismo vive el formulario web, listo
   para usarse (no hace falta desplegar nada aparte en Vercel).

### 3. Nota sobre el plan gratuito

El servicio se "duerme" después de 15 minutos sin uso. La primera llamada después de estar
dormido tarda ~30-50 segundos en responder (Render lo está "despertando"); las siguientes son
rápidas. Para uso interno bajo demanda esto no debería ser un problema.
