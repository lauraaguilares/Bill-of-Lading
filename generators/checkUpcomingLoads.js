/**
 * checkUpcomingLoads.js
 * Revisa "PASTE HERE" del archivo maestro (descargado de Dropbox) y, por cada embarque
 * cuya fecha de carga sea exactamente 7 días a partir de hoy, genera el BOL y lo envía
 * por correo al PMA asignado.
 *
 * "Fecha de carga" (regla confirmada):
 *   Southbound / Bonded -> NOB Day for Truck to Arrive
 *   Northbound          -> CUSTOMS Day for Truck to Arrive
 * (Es el mismo dato que ya se usa como fechaDocumento en generateBOL.js)
 */

// Render (y algunos otros hosts) no soportan salida por IPv6, pero Node intenta IPv6
// primero por default — eso causaba "ENETUNREACH" al conectar a Google Drive. Se fuerza
// IPv4 primero para todas las conexiones salientes del proceso.
require('dns').setDefaultResultOrder('ipv4first');

const fs = require('fs');
const path = require('path');
const { enviarCorreo } = require('./mailTransport');
const ExcelJS = require('exceljs');
const { generateBOL } = require('./generateBOL');
const { readShipmentFromSource, listClientsFromSource } = require('./readShipmentFromSource');
const { CONTACT_POOL, resolveDirectionFromTags } = require('../config/bolConfig');

const TEMP_DIR = path.join(__dirname, '..', 'uploads');

/**
 * Descarga el archivo maestro desde la URL de Dropbox configurada (debe terminar en
 * "&dl=1" o "?dl=1" para forzar descarga directa en vez de abrir la vista previa).
 */
async function descargarArchivoMaestro(url, intento = 1) {
  let res;
  try {
    res = await fetch(url, {
      redirect: 'follow',
      // Timeout generoso: a veces Render tarda en conectar a Google y el default es corto.
      signal: AbortSignal.timeout(45000),
      headers: {
        // Sin esto, Dropbox a veces devuelve una página HTML en vez del archivo real.
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36',
      },
    });
  } catch (err) {
    // Reintento único: la falla suele ser una intermitencia de red de Render, no algo real.
    if (intento < 2) {
      console.log(`[descargarArchivoMaestro] Falló el intento ${intento} (${err.message}), reintentando...`);
      await new Promise((r) => setTimeout(r, 2000));
      return descargarArchivoMaestro(url, intento + 1);
    }
    throw new Error(`No se pudo conectar para descargar el archivo maestro después de 2 intentos: ${err.message}`);
  }

  if (res.status !== 200) {
    throw new Error(`No se pudo descargar el archivo maestro (HTTP ${res.status}). Revisa que el link de Dropbox siga siendo válido y termine en dl=1.`);
  }

  const buffer = Buffer.from(await res.arrayBuffer());

  // Si Dropbox (o algo en el camino) manda el archivo incompleto, el tamaño real no
  // coincide con el que el servidor anunció en el header — eso explica errores tipo
  // "can't find end of central directory" (el .xlsx se cortó antes de terminar).
  const contentLength = res.headers.get('content-length');
  if (contentLength && Number(contentLength) !== buffer.length) {
    throw new Error(
      `El archivo se descargó incompleto: se esperaban ${contentLength} bytes y llegaron ${buffer.length}. ` +
      `Intenta de nuevo; si persiste, puede ser un límite de tamaño o timeout en la descarga.`
    );
  }

  // Un .xlsx real es un .zip por dentro; siempre empieza con las letras "PK". Si no,
  // lo que se descargó fue una página HTML (el link no es de descarga directa) u otra cosa.
  const esZipValido = buffer.length > 2 && buffer[0] === 0x50 && buffer[1] === 0x4b;
  if (!esZipValido) {
    throw new Error(
      'El link de Dropbox no está devolviendo el archivo .xlsx real (parece una página web, no el ' +
      'archivo). Verifica: 1) que termine en "dl=1" y no "dl=0", 2) que el link no haya expirado, ' +
      '3) que el archivo siga compartido públicamente.'
    );
  }

  if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
  const filePath = path.join(TEMP_DIR, `maestro-cron-${Date.now()}.xlsx`);
  fs.writeFileSync(filePath, buffer);
  console.log(`[checkUpcomingLoads] Archivo maestro descargado: ${buffer.length} bytes`);
  return filePath;
}

/**
 * Fecha de carga de un embarque según su dirección (misma regla que fechaDocumento).
 * Devuelve un objeto Date, o null si el dato viene vacío/no parseable.
 */
function fechaDeCarga(shipment) {
  const direction = resolveDirectionFromTags(shipment.tags);
  const usaNOB = direction === 'southbound' || direction === 'bonded_canada';
  const raw = usaNOB ? shipment.nobDiaParaCamion : shipment.customsDiaParaCamion;
  if (!raw) return null;

  const match = String(raw).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!match) return null;
  const [, mm, dd, yyyy] = match;
  return new Date(Number(yyyy), Number(mm) - 1, Number(dd));
}

function esMismaFecha(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

/**
 * Revisa todos los embarques del archivo maestro y devuelve los que cargan en exactamente
 * `diasAntes` días a partir de hoy.
 */
async function embarquesProximosACargar(filePath, diasAntes = 7) {
  const clientes = await listClientsFromSource(filePath);
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const fechaObjetivo = new Date(hoy);
  fechaObjetivo.setDate(fechaObjetivo.getDate() + diasAntes);

  const resultado = [];
  for (const c of clientes) {
    try {
      const shipment = await readShipmentFromSource(filePath, c.row, { fecha: new Date() });
      const carga = fechaDeCarga(shipment);
      if (carga && esMismaFecha(carga, fechaObjetivo)) {
        resultado.push(shipment);
      }
    } catch (err) {
      console.error(`Fila ${c.row} (${c.clientName}) omitida por error al leer:`, err.message);
    }
  }
  return resultado;
}

/**
 * Envía el BOL (adjunto). NOTA: Resend sin dominio propio verificado solo permite mandar
 * a la MISMA dirección con la que te registraste — por eso esto va a Laura (gmailUser),
 * no directo al PMA, con el nombre del PMA en el asunto para que ella sepa a quién
 * reenviarlo. Si se verifica el dominio de la empresa en Resend, se puede volver a mandar
 * directo a cada PMA (pma.correo).
 */
async function enviarBOLPorCorreo(shipment, archivoPath, gmailUser) {
  const pma = CONTACT_POOL[shipment.pma];
  if (!pma) {
    throw new Error(`El PMA "${shipment.pma}" no se reconoce.`);
  }

  await enviarCorreo({
    apiKey: process.env.RESEND_API_KEY,
    from: 'onboarding@resend.dev',
    to: gmailUser,
    subject: `BOL listo para ${pma.nombre} — ${shipment.clientName} (carga en 7 días)`,
    text: `BOL de ${shipment.clientName}, cuya carga es en 7 días.\nPMA asignado: ${pma.nombre} — reenviar a: ${pma.correo || '(sin correo configurado)'}\n\nGenerado automáticamente por BMM Document Generator.`,
    attachments: [{ filename: path.basename(archivoPath), path: archivoPath }],
  });
}

/**
 * Punto de entrada: descarga el archivo maestro, revisa embarques a 7 días, genera y envía.
 * Devuelve un resumen para logging/respuesta del endpoint.
 */
async function checkUpcomingLoadsAndSend({ dropboxUrl, gmailUser }) {
  const filePath = await descargarArchivoMaestro(dropboxUrl);

  const embarques = await embarquesProximosACargar(filePath, 7);
  const resumen = { revisados: true, encontrados: embarques.length, enviados: [], errores: [] };

  for (const shipment of embarques) {
    try {
      // Se envía el .xlsx (no el PDF) para que el equipo pueda editarlo si hace falta.
      const xlsxPath = await generateBOL(shipment);
      await enviarBOLPorCorreo(shipment, xlsxPath, gmailUser);
      resumen.enviados.push({ cliente: shipment.clientName, pma: shipment.pma });
    } catch (err) {
      resumen.errores.push({ cliente: shipment.clientName, error: err.message });
    }
  }

  fs.unlink(filePath, () => {});
  return resumen;
}

module.exports = { checkUpcomingLoadsAndSend, embarquesProximosACargar, fechaDeCarga, descargarArchivoMaestro };
