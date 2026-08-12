/**
 * server.js
 * Backend único para la generación de documentos BMM.
 * Cada tipo de documento vive en su propia ruta /generate/{tipo}.
 * Hoy solo tiene BOL; el checklist y futuros documentos se agregan igual.
 */

const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const { execFile } = require('child_process');
const { generateBOL } = require('./generators/generateBOL');
const { readShipmentFromSource, listClientsFromSource } = require('./generators/readShipmentFromSource');
const { checkUpcomingLoadsAndSend } = require('./generators/checkUpcomingLoads');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const OUTPUT_DIR = path.join(__dirname, 'output');
const UPLOADS_DIR = path.join(__dirname, 'uploads');
[OUTPUT_DIR, UPLOADS_DIR].forEach((dir) => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

const upload = multer({ dest: UPLOADS_DIR });

// Healthcheck (Render lo usa para saber si el servicio sigue vivo)
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

/**
 * POST /parse-source  (multipart, campo "archivo")
 * Sube el archivo maestro (BMM Weeklies) y devuelve la lista de clientes de "PASTE HERE"
 * para poblar el selector del formulario.
 */
app.post('/parse-source', upload.single('archivo'), async (req, res) => {
  try {
    if (!req.file) throw new Error('No se recibió ningún archivo.');
    const clientes = await listClientsFromSource(req.file.path);
    res.json({ clientes, archivoTemporal: req.file.filename });
  } catch (err) {
    console.error(err);
    res.status(400).json({ error: err.message });
  } finally {
    // El archivo se vuelve a subir en /generate/bol-from-source, así que no hace falta
    // conservar este; se limpia para no acumular uploads.
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});

/**
 * POST /generate/bol-from-source  (multipart, campo "archivo" + campos de texto: fila,
 * segundaDireccionPickup1/2 opcionales)
 * Lee la fila indicada de "PASTE HERE" y genera el BOL (xlsx o pdf según ?format=).
 */
app.post('/generate/bol-from-source', upload.single('archivo'), async (req, res) => {
  try {
    if (!req.file) throw new Error('No se recibió ningún archivo.');
    const fila = Number(req.body.fila);
    if (!fila) throw new Error('Falta el número de fila del cliente seleccionado.');

    const options = { fecha: new Date() };
    if (req.body.segundaDireccionPickup1 || req.body.segundaDireccionPickup2) {
      options.segundaDireccionPickup = [
        req.body.segundaDireccionPickup1 || '',
        req.body.segundaDireccionPickup2 || '',
      ];
    }

    const shipment = await readShipmentFromSource(req.file.path, fila, options);
    const xlsxPath = await generateBOL(shipment);

    if (req.query.format === 'pdf') {
      const pdfPath = await convertToPDF(xlsxPath);
      return res.download(pdfPath);
    }
    return res.download(xlsxPath);
  } catch (err) {
    console.error(err);
    return res.status(400).json({ error: err.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});

/**
 * POST /generate/bol
 * Body: ver forma de "shipment" documentada en generators/generateBOL.js
 * Query param opcional: ?format=pdf (default: xlsx)
 * (Generación manual, sin pasar por el archivo maestro — útil para pruebas o casos especiales.)
 */
app.post('/generate/bol', async (req, res) => {
  try {
    const shipment = normalizeShipmentPayload(req.body);
    const xlsxPath = await generateBOL(shipment);

    if (req.query.format === 'pdf') {
      const pdfPath = await convertToPDF(xlsxPath);
      return res.download(pdfPath);
    }

    return res.download(xlsxPath);
  } catch (err) {
    console.error(err);
    return res.status(400).json({ error: err.message });
  }
});

/**
 * GET /cron/check-upcoming-loads?secret=...
 * Pensado para ser llamado 1 vez al día por un disparador externo (GitHub Actions).
 * Revisa el archivo maestro (Dropbox), genera y envía por correo el BOL de cualquier
 * embarque que cargue en exactamente 7 días.
 *
 * Variables de entorno requeridas en Render:
 *   CRON_SECRET          - clave para que nadie más pueda llamar este endpoint
 *   DROPBOX_MASTER_URL   - link de descarga directa del archivo maestro (termina en dl=1)
 *   GMAIL_USER           - cuenta de Gmail que envía los correos
 *   GMAIL_APP_PASSWORD   - contraseña de aplicación de esa cuenta (no la contraseña normal)
 */
app.get('/cron/check-upcoming-loads', async (req, res) => {
  if (!process.env.CRON_SECRET || req.query.secret !== process.env.CRON_SECRET) {
    return res.status(401).json({ error: 'No autorizado.' });
  }
  try {
    const resumen = await checkUpcomingLoadsAndSend({
      dropboxUrl: process.env.DROPBOX_MASTER_URL,
      gmailUser: process.env.GMAIL_USER,
      gmailAppPassword: process.env.GMAIL_APP_PASSWORD,
    });
    console.log('[cron/check-upcoming-loads]', JSON.stringify(resumen));
    res.json(resumen);
  } catch (err) {
    console.error('[cron/check-upcoming-loads] ERROR:', err);
    res.status(500).json({ error: err.message });
  }
});

function normalizeShipmentPayload(body) {
  // El body llega como JSON desde el formulario; 'fecha' viaja como string ISO.
  if (!body.fecha) throw new Error('Falta el campo "fecha"');
  return { ...body, fecha: new Date(body.fecha) };
}

/**
 * Convierte un .xlsx a .pdf usando LibreOffice headless (instalado en la imagen Docker).
 */
function convertToPDF(xlsxPath) {
  return new Promise((resolve, reject) => {
    execFile(
      'soffice',
      ['--headless', '--convert-to', 'pdf', '--outdir', OUTPUT_DIR, xlsxPath],
      (error, stdout, stderr) => {
        if (error) return reject(new Error(`Fallo la conversión a PDF: ${stderr || error.message}`));
        const pdfPath = xlsxPath.replace(/\.xlsx$/, '.pdf').replace(path.dirname(xlsxPath), OUTPUT_DIR);
        resolve(pdfPath);
      }
    );
  });
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`BMM Document Generator escuchando en puerto ${PORT}`));
