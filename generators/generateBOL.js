/**
 * generateBOL.js
 * Genera el archivo .xlsx del Bill of Lading a partir de los datos de un embarque
 * (equivalente a una fila de la hoja "PASTE HERE").
 *
 * Uso:
 *   node generateBOL.js
 *   (ver ejemplo de datos de entrada al final del archivo)
 */

const ExcelJS = require('exceljs');
const path = require('path');
const {
  CUSTOMS_BROKERS,
  resolveTISBlock,
  buildContactBlock,
  resolveDirectionFromTags,
  calcularPeso,
  buildOrderNumbers,
} = require('../config/bolConfig');

const TEMPLATE_PATH = path.join(__dirname, '..', 'templates', 'BOL_plantilla_unica.xlsx');
const OUTPUT_DIR = path.join(__dirname, '..', 'output');

/**
 * shipment: {
 *   clientName: string,
 *   clientLastName: string,
 *   pma: string,               // 'Hannia' | 'Laura' | 'Ana' | 'Ariel'
 *   carrier: string,           // 'TIS' | 'LEGO' | 'Ground Freight Solutions' | otro
 *   numeroOrdenOrigen: string, // Order Number del archivo maestro; reemplaza el Order # generado cuando carrier = ESTES/Old Dominion
 *   customsBroker: 'LEGO' | 'TIS',
 *   tags: string,               // valor crudo de la columna TAGS ('LIVE', 'Northbound', 'Bonded', etc.)
 *   direccionNOB: string[],   // NOB Address ("Address for truck"): Southbound (Shipper), Northbound (Consignee), Bonded (Shipper)
 *   direccionSOB: string[],   // SOB Address ("Address for truck"): solo Bonded (Consignee)
 *   cubicFt: number,
 *   nobDiaParaCamion / nobHoraParaCamion: string,        // "X NOB Day/Time for Truck to Arrive"
 *   customsDiaParaCamion / customsHoraParaCamion: string, // "X CUSTOMS Day/Time for Truck to Arrive"
 *   instruccionesEspeciales: string,  // fallback manual si no se usan los campos nob/customs de arriba
 *   fecha: Date,                      // fecha de generación (para Order#/PO#); fallback del campo "Date:" si no hay nob/customs
 *   segundaDireccionPickup: string[] | null,  // solo bonded, si aplica 2 stops
 *   sobContactoParaConductor: string | null,  // solo bonded: "SOB Contact for Driver" del archivo base
 * }
 */
async function generateBOL(shipment) {
  const direction = resolveDirectionFromTags(shipment.tags);
  const { orderNumber, poNumber } = resolveOrderNumbers(shipment);
  const contacto = buildContactBlock(shipment.pma);
  const peso = calcularPeso(shipment.cubicFt, shipment.carrier);
  const { fechaDocumento, instruccionesEspeciales } = resolveFechas(direction, shipment);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(TEMPLATE_PATH);
  const ws = workbook.getWorksheet('Sheet1');

  stripHighlightFill(ws);
  fitToOnePage(ws);

  // --- Campos comunes ---
  ws.getCell('B9').value = shipment.carrier;
  ws.getCell('H9').value = fechaDocumento;
  ws.getCell('B10').value = orderNumber;
  ws.getCell('H10').value = poNumber;
  ws.getCell('C24').value = shipment.cubicFt;
  ws.getCell('J23').value = peso.esManual ? 'VER NOTA - MANUAL' : { formula: '=+C24*6.5' };

  let instrucciones = insertarDiaSemana(instruccionesEspeciales);

  // --- Shipper / Consignee según escenario ---
  if (direction === 'southbound') {
    const broker = shipment.customsBroker === 'TIS' ? resolveTISBlock('southbound') : CUSTOMS_BROKERS.LEGO;

    setShipper(ws, `${shipment.clientName} /BEST MEXICO MOVERS`, shipment.direccionNOB, contacto.texto);
    setConsignee(ws, broker.nombre, broker.direccion, `${broker.contacto} - ${broker.telefono}`);

  } else if (direction === 'northbound') {
    const broker = shipment.customsBroker === 'TIS' ? resolveTISBlock('northbound') : CUSTOMS_BROKERS.LEGO;

    setShipper(ws, broker.nombre, broker.direccion, `${broker.contacto} - ${broker.telefono}`);
    setConsignee(ws, `${shipment.clientName} /BEST MEXICO MOVERS`, shipment.direccionNOB, contacto.texto);

  } else if (direction === 'bonded_canada') {
    setInBondBanner(ws);

    setShipper(ws, `${shipment.clientName} /BEST MEXICO MOVERS`, shipment.direccionNOB, contacto.texto, shipment.segundaDireccionPickup);

    // En Bonded, el contacto del Consignee (destino) suma el "SOB Contact for Driver"
    // del archivo base (solo el dato, sin etiqueta), además del contacto dinámico.
    const contactoDestino = shipment.sobContactoParaConductor
      ? `${contacto.texto}\n${shipment.sobContactoParaConductor}`
      : contacto.texto;
    setConsignee(ws, `${shipment.clientName} /BEST MEXICO MOVERS`, shipment.direccionSOB, contactoDestino);

    // Bloque de referencia del broker (NO es el consignee, va debajo de la tabla de pieces,
    // aprovechando las filas en blanco que ya trae la tabla para piezas adicionales).
    const broker = shipment.customsBroker === 'TIS' ? resolveTISBlock('southbound') : CUSTOMS_BROKERS.LEGO;
    addBrokerReferenceBlock(ws, broker);
  }

  setSpecialInstructions(ws, instrucciones);

  // --- Guardar ---
  const outName = `BOL_${shipment.clientLastName}_${orderNumber}.xlsx`;
  const outPath = path.join(OUTPUT_DIR, outName);
  await workbook.xlsx.writeFile(outPath);
  return outPath;
}

/**
 * Quita el relleno amarillo que la plantilla trae como guía visual para llenado manual.
 * Ya no aplica: el documento se llena automáticamente, así que el resultado final
 * debe salir limpio, sin resaltados.
 */
function stripHighlightFill(ws) {
  ws.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      if (cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb === 'FFFFFF00') {
        // Clonar el estilo antes de tocarlo (ver nota en setWrappedContactLine sobre
        // estilos compartidos entre celdas en ExcelJS).
        cell.style = JSON.parse(JSON.stringify(cell.style || {}));
        cell.fill = { type: 'pattern', pattern: 'none' };
      }
    });
  });
}

/**
 * Configura la hoja para que siempre imprima/exporte en una sola página,
 * sin importar cuánto contenido varíe entre escenarios (southbound/northbound/bonded).
 */
function fitToOnePage(ws) {
  // El rango usado de la hoja llega hasta la columna N por una celda con un espacio suelto
  // (N15) heredada de la plantilla original — L, M, N no tienen contenido real. LibreOffice,
  // al convertir por línea de comandos, ignora el print area/centrado si esas columnas siguen
  // ahí, así que se eliminan de raíz en vez de solo excluirlas del print area.
  ws.getCell('N15').value = null;
  ['L', 'M', 'N'].forEach((col) => {
    ws.getColumn(col).width = 0;
  });

  ws.pageSetup = {
    ...ws.pageSetup,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    horizontalCentered: true,
    verticalCentered: true,
    printArea: 'A7:K43',
    margins: {
      left: 0.4, right: 0.4, top: 0.3, bottom: 0.3, header: 0.2, footer: 0.2,
    },
  };
  ws.pageSetup.scale = undefined; // fitToPage y scale son mutuamente excluyentes en Excel
}

/**
 * contactoLinea: string ya formateada, ej. "Moises - +1 (520) 486-1110, Ariel - +1 (915) 502-8546, Hannia - +1 (956) 301-0924"
 * o "RIGO DURAN - +1 (915) 892-4958" cuando es el contacto fijo de un broker.
 * segundaDireccionPickup: opcional, solo aplica en Bonded/Canadá con 2 stops de pickup.
 */
function setShipper(ws, nombre, direccionLineas, contactoLinea, segundaDireccionPickup) {
  ws.getCell('A13').value = nombre;

  if (segundaDireccionPickup && segundaDireccionPickup.length) {
    setTwoStopAddress(ws, direccionLineas, segundaDireccionPickup);
  } else {
    setWrappedAddress(ws, 'A14', 'D15', direccionLineas);
  }

  setWrappedContactLine(ws, 'A16', 'D17', `  CONTACT:  ${contactoLinea}`);
}

/**
 * Las direcciones del archivo maestro vienen como texto libre y a veces largo/desordenado.
 * Sin wrap, una dirección larga se desborda visualmente hacia la columna del Consignee.
 * Se fusiona y envuelve dentro de su propio bloque, igual que el patrón ya usado para
 * el contacto dinámico, para que nunca invada el lado contrario.
 */
function setWrappedAddress(ws, fromCell, toCell, direccionLineas) {
  const texto = (direccionLineas || []).filter(Boolean).join('\n');
  const cell = ws.getCell(fromCell);
  ws.mergeCells(`${fromCell}:${toCell}`);
  cell.style = JSON.parse(JSON.stringify(cell.style || {}));
  cell.value = texto;
  cell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
  cell.font = { ...cell.font, size: 9.5 };
  const filaTop = Number(fromCell.match(/\d+/)[0]);
  const filaBottom = Number(toCell.match(/\d+/)[0]);
  ws.getRow(filaTop).height = Math.max(ws.getRow(filaTop).height || 0, 24);
  ws.getRow(filaBottom).height = Math.max(ws.getRow(filaBottom).height || 0, 24);
}

/**
 * Bonded/Canadá puede tener 1 o 2 direcciones de pickup (se decide al generar el documento,
 * no viene de la fuente de datos). Cuando hay 2, se listan como "First stop: ..." / "Second stop: ..."
 * usando el mismo patrón de celda fusionada + wrap que el bloque de contacto, para que quepan
 * completas sin cortarse.
 */
function setTwoStopAddress(ws, primeraDireccion, segundaDireccion) {
  const texto = `First stop: ${primeraDireccion.join(', ')}\nSecond stop: ${segundaDireccion.join(', ')}`;
  const cell = ws.getCell('A14');
  ws.mergeCells('A14:D15');
  cell.style = JSON.parse(JSON.stringify(cell.style || {}));
  cell.value = texto;
  cell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
  cell.font = { ...cell.font, size: 9.5 };
  ws.getRow(14).height = 24;
  ws.getRow(15).height = 24;
}

function setConsignee(ws, nombre, direccionLineas, contactoLinea) {
  ws.getCell('G13').value = nombre;
  setWrappedAddress(ws, 'G14', 'K15', direccionLineas);
  setWrappedContactLine(ws, 'G16', 'K17', `  CONTACT:  ${contactoLinea}`);
}

/**
 * El bloque de contacto puede traer hasta 3 personas, una por línea ("Nombre - Tel\nNombre - Tel...").
 * Una sola celda angosta (columna A o G) partía cada línea a la mitad. Se fusiona un bloque de
 * 4 columnas x 2 filas (ancho Y alto) para tener espacio real en ambas dimensiones, así
 * "Nombre - Teléfono" siempre cabe en un renglón sin importar si son 2 o 3 personas.
 */
function setWrappedContactLine(ws, fromCell, toCell, texto) {
  const range = `${fromCell}:${toCell}`;
  ws.mergeCells(range);
  const cell = ws.getCell(fromCell);
  // Clonar el estilo antes de tocarlo: celdas con el mismo estilo original pueden compartir
  // el mismo objeto en memoria en ExcelJS, y modificarlo directo "contamina" otras celdas.
  cell.style = JSON.parse(JSON.stringify(cell.style || {}));
  cell.value = texto;
  cell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
  cell.font = { ...cell.font, size: 9.5 };

  const numPersonas = (texto.match(/\n/g) || []).length + 1;
  const filaTop = Number(fromCell.match(/\d+/)[0]);
  const filaBottom = Number(toCell.match(/\d+/)[0]);
  const alturaTotal = numPersonas * 15 + 10;
  // Shipper y Consignee comparten estas mismas filas (16 y 17). Si el consignee (ej. broker,
  // 1 línea) se procesa después del shipper (ej. 3 líneas), NO debe achicar la fila que el
  // shipper ya dejó más alta — se toma el máximo entre lo que ya había y lo que hace falta ahora.
  const rowTop = ws.getRow(filaTop);
  const rowBottom = ws.getRow(filaBottom);
  rowTop.height = Math.max(rowTop.height || 0, Math.ceil(alturaTotal / 2));
  rowBottom.height = Math.max(rowBottom.height || 0, Math.floor(alturaTotal / 2));
}

/**
 * Special Instructions puede ser 1 línea normalmente, o 2 cuando es Bonded ("* IN BOND*" +
 * la instrucción de entrega). Se controla explícitamente (fusionar ancho completo, wrap,
 * altura de fila) en vez de dejar la celda en su estado por defecto: se detectó contaminación
 * de estilo compartido de ExcelJS (ver nota en setWrappedContactLine) que activaba wrap sobre
 * el ancho angosto de una sola columna y partía el texto en fragmentos ilegibles.
 */
function setSpecialInstructions(ws, texto) {
  const cell = ws.getCell('C19');
  ws.mergeCells('C19:K19');
  cell.style = JSON.parse(JSON.stringify(cell.style || {}));
  cell.value = texto;

  const numLineas = (texto.match(/\n/g) || []).length + 1;
  if (numLineas > 1) {
    cell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
    ws.getRow(19).height = numLineas * 14 + 6;
  } else {
    cell.alignment = { wrapText: false, vertical: 'bottom', horizontal: 'left' };
    ws.getRow(19).height = 19.5;
  }
}

/**
 * Special Instructions trae la fecha de la cita en formato MM/DD/YYYY (ej. "DELIVERY
 * APPOINTMENT ON 09/20/2026 @ 9 AM"). Se le antepone el día de la semana automáticamente
 * ("DELIVERY APPOINTMENT ON Sunday, 09/20/2026 @ 9 AM") en vez de requerir que se escriba a mano.
 * Si el texto ya trae un día de la semana, o no encuentra una fecha, lo deja igual.
 */
function insertarDiaSemana(texto) {
  const DIAS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return texto.replace(/\b(\d{1,2})\/(\d{1,2})\/(\d{4})\b/g, (match, mm, dd, yyyy, offset, full) => {
    // Si ya hay un nombre de día justo antes (ej. "Monday, 09/20/2026"), no duplicar.
    const antes = full.slice(Math.max(0, offset - 12), offset);
    if (DIAS.some((d) => antes.includes(d))) return match;

    const fecha = new Date(Number(yyyy), Number(mm) - 1, Number(dd));
    const diaSemana = DIAS[fecha.getDay()];
    return `${diaSemana}, ${match}`;
  });
}

/**
 * "* IN BOND*" ya no va pegado al texto de Special Instructions (se veía chico y a la
 * izquierda, mezclado con la instrucción de entrega). Se pone como su propio banner,
 * centrado y en letra grande, en la fila 18 (libre en la plantilla, justo arriba de
 * "Special Instructions:").
 */
function setInBondBanner(ws) {
  const cell = ws.getCell('A18');
  ws.mergeCells('A18:K18');
  cell.style = JSON.parse(JSON.stringify(cell.style || {}));
  cell.value = '* IN BOND *';
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  cell.font = { ...cell.font, size: 16, bold: true };
  ws.getRow(18).height = 22;
}

/**
 * Regla de negocio confirmada para las fechas del BOL — el campo "Date:" y el de
 * "Special Instructions" NO siempre usan la misma columna del archivo maestro; depende
 * del escenario:
 *
 *   Southbound  -> Date: NOB Day for Truck to Arrive     | Special Instructions: CUSTOMS Day
 *   Northbound  -> Date: CUSTOMS Day for Truck to Arrive | Special Instructions: NOB Day
 *   Bonded      -> Date: NOB Day for Truck to Arrive     | Special Instructions: CUSTOMS Day
 *
 * Si el shipment no trae los campos nob/customsDiaParaCamion (ej. generado a mano, sin pasar
 * por readShipmentFromSource), se cae de vuelta al comportamiento anterior: "Date:" usa
 * shipment.fecha y "Special Instructions" usa shipment.instruccionesEspeciales tal cual.
 */
function resolveFechas(direction, shipment) {
  const tieneDatosDeArchivo = shipment.nobDiaParaCamion !== undefined || shipment.customsDiaParaCamion !== undefined;

  if (!tieneDatosDeArchivo) {
    return {
      fechaDocumento: shipment.fecha.toLocaleDateString('en-US'),
      instruccionesEspeciales: shipment.instruccionesEspeciales || '',
    };
  }

  const nob = { dia: shipment.nobDiaParaCamion || '', hora: shipment.nobHoraParaCamion || '9 AM' };
  const customs = { dia: shipment.customsDiaParaCamion || '', hora: shipment.customsHoraParaCamion || '9 AM' };

  const usaNobParaDate = direction === 'southbound' || direction === 'bonded_canada';
  const fechaDocumento = usaNobParaDate ? nob.dia : customs.dia;
  const paraInstrucciones = usaNobParaDate ? customs : nob;

  const instruccionesEspeciales = paraInstrucciones.dia
    ? `DELIVERY APPOINTMENT ON ${paraInstrucciones.dia} @ ${paraInstrucciones.hora}`
    : '';

  return { fechaDocumento, instruccionesEspeciales };
}

/**
 * Order # normalmente sigue el formato BMM-{INICIALES}{FECHA}-01. Excepción confirmada:
 * si el carrier es ESTES u Old Dominion, el Order # debe ser el mismo número que ya trae
 * el documento de origen (columna "Order Number" de PASTE HERE), no el generado. El PO #
 * mantiene siempre el formato normal.
 */
function resolveOrderNumbers(shipment) {
  const { orderNumber, poNumber } = buildOrderNumbers(shipment.clientLastName, shipment.fecha);

  const carrier = (shipment.carrier || '').toUpperCase();
  const esEstesOOldDominion = carrier.includes('ESTES') || carrier.includes('OLD DOMINION');

  if (esEstesOOldDominion && shipment.numeroOrdenOrigen) {
    return { orderNumber: shipment.numeroOrdenOrigen, poNumber };
  }
  return { orderNumber, poNumber };
}

function addBrokerReferenceBlock(ws, broker) {
  // Filas 25-30 son parte de la tabla de "Pieces" pero vienen en blanco en la plantilla
  // (reservadas por si hay más de una línea de carga). Se reutilizan para el bloque de
  // referencia del broker en Bonded/Canadá, ya que el diseño original no dejaba espacio
  // dedicado para esto en una plantilla de una sola página.
  ws.getCell('C25').value = 'Custom Broker:';
  ws.getCell('C26').value = broker.nombre;
  ws.getCell('C27').value = broker.direccion[0];
  ws.getCell('C28').value = `${broker.direccion[1]} - PH: ${broker.telefono}`;
  if (broker.email) ws.getCell('C29').value = `EMAIL: ${broker.email}`;
}

module.exports = { generateBOL };

// --- Ejemplo de uso ---
if (require.main === module) {
  const ejemplo = {
    clientName: 'JOHN SMITH',
    clientLastName: 'Smith',
    pma: 'Ariel',
    carrier: 'TIS Worldwide',
    customsBroker: 'TIS',
    origenPais: 'US',
    destinoPais: 'MX',
    direccionSOB: ['100 Oak Street', 'Elko, NV 89801'],
    cubicFt: 2400,
    instruccionesEspeciales: 'DELIVERY APPOINTMENT ON 08/20/2026 @ 9 AM',
    fecha: new Date('2026-08-20'),
  };

  generateBOL(ejemplo)
    .then((outPath) => console.log('BOL generado en:', outPath))
    .catch((err) => console.error('Error generando BOL:', err));
}
