/**
 * server.js
 * Backend único para la generación de documentos BMM.
 * Cada tipo de documento vive en su propia ruta /generate/{tipo}.
 * Hoy solo tiene BOL; el checklist y futuros documentos se agregan igual.
 */

// Render no soporta salida por IPv6. El flag --dns-result-order=ipv4first del Dockerfile
// arregla esto para fetch() (Google Drive), pero nodemailer/SMTP usa su propia conexión de
// bajo nivel que NO respeta ese flag — seguía intentando IPv6 para conectar a Gmail y
// tronaba con ENETUNREACH. Se intercepta dns.lookup directamente (a nivel de todo el
// proceso) para forzar IPv4 siempre, sin importar qué librería lo llame.
const dns = require('dns');
const dnsLookupOriginal = dns.lookup;
dns.lookup = (hostname, options, callback) => {
  if (typeof options === 'function') {
    callback = options;
    options = {};
  }
  return dnsLookupOriginal(hostname, { ...options, family: 4 }, callback);
};

const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const { execFile } = require('child_process');
const { generateBOL } = require('./generators/generateBOL');
const { readShipmentFromSource, listClientsFromSource } = require('./generators/readShipmentFromSource');
const { checkUpcomingLoadsAndSend, descargarArchivoMaestro } = require('./generators/checkUpcomingLoads');
const { generateWeekly, generateWeeklyImage, listarFiltrosDisponibles } = require('./generators/generateWeekly');
const { WEEKLY_TIPOS } = require('./config/weeklyConfig');
const { crearTransportadorGmail } = require('./generators/mailTransport');

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
/**
 * GET /parse-drive
 * Lee el archivo maestro directo desde Google Drive (la misma URL que usa el cron
 * diario) y devuelve la lista de clientes — sin que nadie tenga que subir nada a mano.
 */
app.get('/parse-drive', async (req, res) => {
  try {
    if (!process.env.DROPBOX_MASTER_URL) throw new Error('No hay un archivo maestro configurado en el servidor.');
    const filePath = await descargarArchivoMaestro(process.env.DROPBOX_MASTER_URL);
    const clientes = await listClientsFromSource(filePath);
    fs.unlink(filePath, () => {});
    res.json({ clientes });
  } catch (err) {
    console.error(err);
    res.status(400).json({ error: err.message });
  }
});

/**
 * POST /generate/bol-from-drive
 * Body JSON: { fila, segundaDireccionPickup1?, segundaDireccionPickup2? }
 * Igual que /generate/bol-from-source pero lee el archivo maestro directo de Drive.
 */
app.post('/generate/bol-from-drive', async (req, res) => {
  let filePath;
  try {
    if (!process.env.DROPBOX_MASTER_URL) throw new Error('No hay un archivo maestro configurado en el servidor.');
    const fila = Number(req.body.fila);
    if (!fila) throw new Error('Falta el número de fila del cliente seleccionado.');

    const options = { fecha: new Date() };
    if (req.body.segundaDireccionPickup1 || req.body.segundaDireccionPickup2) {
      options.segundaDireccionPickup = [
        req.body.segundaDireccionPickup1 || '',
        req.body.segundaDireccionPickup2 || '',
      ];
    }

    filePath = await descargarArchivoMaestro(process.env.DROPBOX_MASTER_URL);
    const shipment = await readShipmentFromSource(filePath, fila, options);
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
    if (filePath) fs.unlink(filePath, () => {});
  }
});

/**
 * POST /parse-source (subir archivo a mano — se conserva como opción alterna)
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

/**
 * GET /weekly/filtros?tipo=transportista
 * Devuelve la lista de valores disponibles (ej. transportistas con embarques activos)
 * para poblar el selector del formulario, leyendo el archivo maestro desde Drive.
 */
app.get('/weekly/filtros', async (req, res) => {
  let filePath;
  try {
    const tipo = req.query.tipo;
    if (!WEEKLY_TIPOS[tipo]) throw new Error(`Tipo de weekly desconocido: "${tipo}"`);
    if (!process.env.DROPBOX_MASTER_URL) throw new Error('No hay un archivo maestro configurado en el servidor.');

    filePath = await descargarArchivoMaestro(process.env.DROPBOX_MASTER_URL);
    const filtros = await listarFiltrosDisponibles(tipo, filePath);
    res.json({ filtros, formatoSalida: WEEKLY_TIPOS[tipo].formatoSalida });
  } catch (err) {
    console.error(err);
    res.status(400).json({ error: err.message });
  } finally {
    if (filePath) fs.unlink(filePath, () => {});
  }
});

/**
 * POST /generate/weekly
 * Body JSON: { tipo, filtroValor }
 * Genera la weekly correspondiente (Excel para transportista/agenteAduanal, imagen para
 * crew) leyendo el archivo maestro directo de Drive, y la devuelve para descargar.
 */
app.post('/generate/weekly', async (req, res) => {
  let filePath;
  try {
    const { tipo, filtroValor } = req.body;
    if (!WEEKLY_TIPOS[tipo]) throw new Error(`Tipo de weekly desconocido: "${tipo}"`);
    if (!filtroValor) throw new Error('Falta el valor a filtrar (transportista/broker/crew).');
    if (!process.env.DROPBOX_MASTER_URL) throw new Error('No hay un archivo maestro configurado en el servidor.');

    filePath = await descargarArchivoMaestro(process.env.DROPBOX_MASTER_URL);

    if (WEEKLY_TIPOS[tipo].formatoSalida === 'imagen') {
      const { pngPath } = await generateWeeklyImage(tipo, filtroValor, filePath);
      return res.download(pngPath);
    }
    const { outPath } = await generateWeekly(tipo, filtroValor, filePath);
    return res.download(outPath);
  } catch (err) {
    console.error(err);
    return res.status(400).json({ error: err.message });
  } finally {
    if (filePath) fs.unlink(filePath, () => {});
  }
});

/**
 * GET /cron/send-weeklies?secret=...
 * Pensado para correr 1 vez cada lunes. Genera TODAS las weeklies (todos los
 * transportistas, brokers y crews con embarques activos) y se las manda a Laura en un
 * solo correo, para que ella las reenvíe a cada quien.
 */
app.get('/cron/send-weeklies', async (req, res) => {
  if (!process.env.CRON_SECRET || req.query.secret !== process.env.CRON_SECRET) {
    return res.status(401).json({ error: 'No autorizado.' });
  }

  // Responde de inmediato: generar todas las weeklies (varios transportistas + brokers +
  // ~10 crews en imagen) tarda varios minutos, y Render corta la conexión HTTP antes de que
  // termine si se espera aquí. El trabajo real sigue en segundo plano después de responder.
  res.json({ iniciado: true, mensaje: 'Generación en curso, llegará un correo cuando termine.' });
  generarYEnviarTodasLasWeeklies().catch((err) => {
    console.error('[cron/send-weeklies] ERROR de fondo:', err);
  });
});

async function generarYEnviarTodasLasWeeklies() {
  console.log('[cron/send-weeklies] Arrancando...');
  let filePath;
  try {
    if (!process.env.DROPBOX_MASTER_URL) throw new Error('No hay un archivo maestro configurado en el servidor.');
    filePath = await descargarArchivoMaestro(process.env.DROPBOX_MASTER_URL);
    console.log('[cron/send-weeklies] Archivo maestro descargado:', filePath);

    const adjuntos = [];
    const resumen = { generadas: [], errores: [] };

    for (const tipo of Object.keys(WEEKLY_TIPOS)) {
      const filtros = await listarFiltrosDisponibles(tipo, filePath);
      console.log(`[cron/send-weeklies] ${tipo}: ${filtros.length} opciones ->`, filtros.join(', '));
      for (const filtroValor of filtros) {
        try {
          if (WEEKLY_TIPOS[tipo].formatoSalida === 'imagen') {
            const { pngPath } = await generateWeeklyImage(tipo, filtroValor, filePath);
            adjuntos.push({ filename: path.basename(pngPath), path: pngPath });
          } else {
            const { outPath } = await generateWeekly(tipo, filtroValor, filePath);
            adjuntos.push({ filename: path.basename(outPath), path: outPath });
          }
          resumen.generadas.push(`${tipo}: ${filtroValor}`);
          console.log(`[cron/send-weeklies] OK ${tipo}: ${filtroValor} (${resumen.generadas.length} generadas hasta ahora)`);
        } catch (err) {
          resumen.errores.push(`${tipo}: ${filtroValor} -> ${err.message}`);
          console.error(`[cron/send-weeklies] FALLÓ ${tipo}: ${filtroValor} ->`, err.message);
        }
      }
    }

    console.log('[cron/send-weeklies] Generación terminada, armando correo con', adjuntos.length, 'adjuntos...');
    const transporter = await crearTransportadorGmail(process.env.GMAIL_USER, process.env.GMAIL_APP_PASSWORD);
    await transporter.sendMail({
      from: process.env.GMAIL_USER,
      to: process.env.GMAIL_USER, // se manda a Laura misma, ella reenvía a cada quien
      subject: `Weeklies del ${new Date().toLocaleDateString('en-US')} — listas para reenviar`,
      text: `Se generaron ${resumen.generadas.length} weeklies:\n\n${resumen.generadas.join('\n')}` +
        (resumen.errores.length ? `\n\nCon errores:\n${resumen.errores.join('\n')}` : ''),
      attachments: adjuntos,
    });

    console.log('[cron/send-weeklies] Correo enviado. Resumen:', JSON.stringify(resumen));
  } catch (err) {
    console.error('[cron/send-weeklies] ERROR:', err);
    throw err;
  } finally {
    if (filePath) fs.unlink(filePath, () => {});
  }
}

// Por si algo se cae de forma silenciosa fuera de los try/catch de arriba (ej. dentro de
// una promesa que no se esperó bien) — sin esto, un error así no deja NINGÚN rastro en logs.
process.on('unhandledRejection', (err) => {
  console.error('[unhandledRejection]', err);
});
process.on('uncaughtException', (err) => {
  console.error('[uncaughtException]', err);
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
