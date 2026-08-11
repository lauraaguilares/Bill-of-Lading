/**
 * readShipmentFromSource.js
 * Lee una fila de la hoja "PASTE HERE" del archivo maestro (BMM Weeklies) y arma el objeto
 * "shipment" que generateBOL() necesita, aplicando las reglas de negocio confirmadas.
 */

const ExcelJS = require('exceljs');

// Columnas confirmadas de "PASTE HERE" (ver headers reales del archivo maestro)
const COLS = {
  CLIENT: 'A',
  TAGS: 'B',
  PMA: 'C',
  NOB_TRANSPORTATION_COMPANY: 'E',
  NOB_ADDRESS: 'I', // "Address for truck" del lado Canadá/origen — NUNCA usar direcciones de Customs
  ORDER_NUMBER: 'G', // Order Number del documento de origen (solo se usa con carrier ESTES/OLD DOMINION)
  SIZE_OF_SHIPMENT: 'Q',
  CUSTOMS_BROKER: 'U',
  SOB_ADDRESS: 'AC', // "Address for truck" del lado México/destino
  SOB_CONTACT_FOR_DRIVER: 'AD',
  NOB_DAY_FOR_TRUCK_TO_ARRIVE: 'J', // "X NOB Day for Truck to Arrive" - ya viene en MM/DD/YYYY
  NOB_TIME: 'K',
  CUSTOMS_DAY_FOR_TRUCK_TO_ARRIVE: 'V', // "X CUSTOMS Day for Truck to Arrive" - ya viene en MM/DD/YYYY
  CUSTOMS_TIME: 'W',
};

/**
 * Lee la hoja "PASTE HERE" y devuelve el shipment de la fila indicada (2 = primera fila de datos).
 * options: { fecha: Date, segundaDireccionPickup?: string[] } — datos que NO vienen del archivo
 * (la fecha de generación y, si aplica, el segundo stop de Bonded se deciden al generar).
 */
async function readShipmentFromSource(filePath, rowNumber, options = {}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const ws = workbook.getWorksheet('PASTE HERE');
  if (!ws) throw new Error('No se encontró la hoja "PASTE HERE" en el archivo.');

  const get = (col) => ws.getCell(`${col}${rowNumber}`).value;

  const clientRaw = get(COLS.CLIENT);
  if (!clientRaw) throw new Error(`La fila ${rowNumber} de "PASTE HERE" está vacía o no tiene CLIENT.`);

  const { nombreCompleto, apellido } = parseClientName(clientRaw);
  const pma = parsePMA(get(COLS.PMA));
  const customsBroker = parseCustomsBroker(get(COLS.CUSTOMS_BROKER));
  const cubicFt = parseCubicFt(get(COLS.SIZE_OF_SHIPMENT));

  return {
    clientName: nombreCompleto,
    clientLastName: apellido,
    pma,
    carrier: String(get(COLS.NOB_TRANSPORTATION_COMPANY) || '').trim(),
    // Solo se usa cuando el carrier es ESTES u Old Dominion (ver regla en generateBOL.js).
    numeroOrdenOrigen: get(COLS.ORDER_NUMBER) ? String(get(COLS.ORDER_NUMBER)).trim() : '',
    customsBroker,
    tags: String(get(COLS.TAGS) || ''),
    // "Address for truck" — nunca la de Customs. NOB Address y SOB Address son genéricas
    // (no significan literalmente Canadá/México), cada escenario las usa distinto según
    // confirmamos: Southbound Shipper = NOB, Consignee = broker. Northbound Shipper = broker,
    // Consignee = SOB. Bonded Shipper = NOB, Consignee = SOB.
    direccionNOB: parseDireccion(get(COLS.NOB_ADDRESS)),
    direccionSOB: parseDireccion(get(COLS.SOB_ADDRESS)),
    cubicFt,
    // Fechas crudas de las 2 columnas de referencia. generateBOL.js decide cuál va en "Date:"
    // y cuál en "Special Instructions" según el escenario (la regla cambia por dirección).
    nobDiaParaCamion: get(COLS.NOB_DAY_FOR_TRUCK_TO_ARRIVE) ? String(get(COLS.NOB_DAY_FOR_TRUCK_TO_ARRIVE)).trim() : '',
    nobHoraParaCamion: get(COLS.NOB_TIME) ? String(get(COLS.NOB_TIME)).trim() : '',
    customsDiaParaCamion: get(COLS.CUSTOMS_DAY_FOR_TRUCK_TO_ARRIVE) ? String(get(COLS.CUSTOMS_DAY_FOR_TRUCK_TO_ARRIVE)).trim() : '',
    customsHoraParaCamion: get(COLS.CUSTOMS_TIME) ? String(get(COLS.CUSTOMS_TIME)).trim() : '',
    fecha: options.fecha || new Date(),
    segundaDireccionPickup: options.segundaDireccionPickup || null,
    sobContactoParaConductor: get(COLS.SOB_CONTACT_FOR_DRIVER) ? String(get(COLS.SOB_CONTACT_FOR_DRIVER)).trim() : null,
  };
}

/**
 * CLIENT viene como "Apellido, Nombre" (a veces con dígitos pegados, ej. "Basmaji, Pierre1").
 * Devuelve nombre completo en mayúsculas para el documento y el apellido limpio para el Order #.
 */
function parseClientName(clientRaw) {
  const partes = String(clientRaw).split(',').map((s) => s.trim());
  const apellido = (partes[0] || '').replace(/\d+$/, '').trim();
  const nombre = (partes[1] || '').replace(/\d+$/, '').trim();
  const nombreCompleto = `${nombre} ${apellido}`.trim().toUpperCase();
  return { nombreCompleto, apellido };
}

/**
 * PMA viene como "Laura Aguilar", "Hannia Alcala", etc. Solo se necesita el primer nombre,
 * que debe coincidir con el pool de contactos (Hannia/Laura/Ana/Ariel).
 */
function parsePMA(pmaRaw) {
  const primerNombre = String(pmaRaw || '').trim().split(/\s+/)[0];
  const VALIDOS = ['Hannia', 'Laura', 'Ana', 'Ariel'];
  const match = VALIDOS.find((v) => v.toLowerCase() === primerNombre.toLowerCase());
  if (!match) {
    throw new Error(`PMA "${pmaRaw}" no coincide con ningún nombre del pool de contactos (Hannia/Laura/Ana/Ariel).`);
  }
  return match;
}

/**
 * Customs Broker viene ya como 'TIS' o 'LEGO' directamente en el archivo.
 */
function parseCustomsBroker(raw) {
  const v = String(raw || '').toUpperCase();
  if (v.includes('TIS')) return 'TIS';
  if (v.includes('LEGO')) return 'LEGO';
  throw new Error(`Customs Broker "${raw}" no reconocido (se esperaba TIS o LEGO).`);
}

/**
 * "Size of the Shipment" es texto libre ("Survey shows 1,400 ft3", "estimated 800 cf", etc.).
 * Se extrae el primer número.
 */
function parseCubicFt(raw) {
  const match = String(raw || '').match(/[\d,]+/);
  if (!match) throw new Error(`No se pudo extraer un número de "Size of the Shipment": "${raw}"`);
  return Number(match[0].replace(/,/g, ''));
}

/**
 * Las direcciones vienen como una sola línea de texto libre en el archivo maestro. Con el wrap
 * activado en la plantilla (ver setWrappedAddress en generateBOL.js) no hace falta partirlas
 * a mano: se manda tal cual y la celda las acomoda solita en las líneas que necesite.
 */
function parseDireccion(raw) {
  const texto = String(raw || '').trim();
  return texto ? [texto] : [''];
}

/**
 * Lee "PASTE HERE" y devuelve la lista de clientes disponibles (fila, nombre, tags) para
 * poblar un selector en el formulario, sin tener que abrir el Excel a mano.
 */
async function listClientsFromSource(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const ws = workbook.getWorksheet('PASTE HERE');
  if (!ws) throw new Error('No se encontró la hoja "PASTE HERE" en el archivo.');

  const clientes = [];
  let row = 2; // fila 1 = headers
  while (true) {
    const clientRaw = ws.getCell(`${COLS.CLIENT}${row}`).value;
    if (!clientRaw) break;
    clientes.push({
      row,
      clientName: String(clientRaw).trim(),
      tags: String(ws.getCell(`${COLS.TAGS}${row}`).value || '').trim(),
    });
    row += 1;
  }
  return clientes;
}

module.exports = { readShipmentFromSource, listClientsFromSource };
