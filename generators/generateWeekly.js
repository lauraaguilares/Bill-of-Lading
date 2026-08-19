/**
 * generateWeekly.js
 * Genera una Weekly (Excel o insumo para imagen) filtrando "PASTE HERE" por el criterio
 * del tipo indicado (transportista / agenteAduanal / crew), aplicando las reglas de color
 * confirmadas.
 */

const ExcelJS = require('exceljs');
const path = require('path');
const { COLORES, esTextoPendiente, WEEKLY_TIPOS } = require('../config/weeklyConfig');

const OUTPUT_DIR = path.join(__dirname, '..', 'output');

/**
 * Lee "PASTE HERE" completo como un array de objetos { columna: valor }, usando la fila 1
 * como encabezados.
 */
async function leerPasteHere(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const ws = workbook.getWorksheet('PASTE HERE');
  if (!ws) throw new Error('No se encontró la hoja "PASTE HERE" en el archivo.');

  const headers = {};
  ws.getRow(1).eachCell((cell, colNumber) => {
    if (cell.value) headers[String(cell.value).trim()] = colNumber;
  });

  const filas = [];
  let row = 2;
  while (true) {
    const clientDirecto = ws.getCell(row, headers['CLIENT']).value;
    if (!clientDirecto) break;

    const fila = {};
    Object.entries(headers).forEach(([nombre, col]) => {
      fila[nombre] = ws.getCell(row, col).value;
    });
    filas.push(fila);
    row += 1;
  }
  return filas;
}

/**
 * Normaliza texto para comparar filtros (trim, minúsculas). Así "Ross " y "Ross" matchean,
 * igual que "LEGO" y "lego".
 */
function normalizar(valor) {
  return String(valor || '').trim().toLowerCase();
}

/**
 * Un valor de fecha "MM/DD/YYYY" (columna cruda "X ..." del archivo maestro) a objeto Date,
 * o null si no es parseable (incluye "-", vacío, "TBD", etc.).
 */
function parsearFecha(valor) {
  const match = String(valor || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!match) return null;
  const [, mm, dd, yyyy] = match;
  return new Date(Number(yyyy), Number(mm) - 1, Number(dd));
}

function esFechaPasada(valor) {
  const fecha = parsearFecha(valor);
  if (!fecha) return false;
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  return fecha < hoy;
}

/**
 * Genera la Weekly de un tipo+filtro específico. Devuelve la ruta del .xlsx generado.
 */
async function generateWeekly(tipo, filtroValor, sourceFilePath, fechaWeeklies = new Date()) {
  const config = WEEKLY_TIPOS[tipo];
  if (!config) throw new Error(`Tipo de weekly desconocido: "${tipo}"`);

  const filas = await leerPasteHere(sourceFilePath);
  let filasFiltradas = config.sinFiltro
    ? filas.slice() // sin filtro: TODA la información, tal cual (ej. weekly maestra de SOB Crew)
    : filas.filter((f) => normalizar(f[config.filtroColumna]) === normalizar(filtroValor));

  if (!config.incluyeOnHold) {
    // "On Hold" se excluye de la mayoría de las weeklies (confirmado) — no aplica a esas
    // cargas hasta que se reactiven. Algunos tipos (ej. sobCrewCompleto) sí las incluyen.
    filasFiltradas = filasFiltradas.filter((f) => !normalizar(f['TAGS']).includes('on hold'));
  }

  // Orden por la fecha principal de esa weekly (NOB Day para transportistas, Arrival day
  // para brokers, etc.) — no alfabético por cliente. Filas sin fecha válida van al final.
  const colFechaOrden = config.columnas.find((c) => c.esFechaPrincipal);
  if (colFechaOrden) {
    filasFiltradas.sort((a, b) => {
      const fechaA = parsearFecha(a[colFechaOrden.col]);
      const fechaB = parsearFecha(b[colFechaOrden.col]);
      if (!fechaA && !fechaB) return 0;
      if (!fechaA) return 1; // sin fecha al final
      if (!fechaB) return -1;
      return fechaA - fechaB;
    });
  }

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet(filtroValor.slice(0, 31)); // límite de 31 caracteres en nombres de hoja

  // Fila 1: fecha de la weekly (resaltada en amarillo, igual que en la plantilla original)
  ws.getCell('A1').value = 'Date for Weeklies';
  ws.getCell('B1').value = fechaWeeklies.toLocaleDateString('en-US');
  ws.getCell('A1').font = { bold: true };
  ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  ws.getCell('B1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };

  // Fila 2: encabezados
  config.columnas.forEach((colDef, i) => {
    const cell = ws.getCell(2, i + 1);
    cell.value = colDef.label;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5496' } };
  });

  // Filas de datos
  filasFiltradas.forEach((fila, filaIdx) => {
    const rowNum = filaIdx + 3;
    const mismoProveedor = normalizar(fila['Customs Broker']) === normalizar(fila['NOB Transportation Company']) && normalizar(fila['Customs Broker']) !== '';
    // Bandeado alterno (blanco / azul claro), igual que tu plantilla original — hace que
    // cada renglón se distinga fácil del siguiente sin estorbar los colores de negocio.
    const colorBanda = filaIdx % 2 === 0 ? 'FFFFFFFF' : 'FFD9E2F3';

    config.columnas.forEach((colDef, colIdx) => {
      const valor = fila[colDef.col];
      const cell = ws.getCell(rowNum, colIdx + 1);
      cell.value = valor != null ? valor : '';

      aplicarColor(cell, colDef, valor, mismoProveedor, colorBanda);
    });
  });

  // Ancho de columna: máximo ~115px (≈15.5 unidades de Excel) con "Ajustar texto" activado,
  // para que ninguna columna sea más ancha que eso pero el contenido siga viéndose completo
  // (el texto se envuelve en 2+ líneas dentro de la celda en vez de cortarse).
  const ANCHO_MAX_COLUMNA = 15.5; // ≈115px
  config.columnas.forEach((colDef, i) => {
    const largos = [colDef.label.length, ...filasFiltradas.map((f) => String(f[colDef.col] || '').length)];
    const anchoIdeal = Math.max(...largos) + 2;
    ws.getColumn(i + 1).width = Math.min(Math.max(anchoIdeal, 12), ANCHO_MAX_COLUMNA);
  });

  // Ajustar texto (wrap) en encabezados y datos, para que el contenido de columnas angostas
  // se vea completo en varias líneas en vez de cortarse.
  for (let r = 2; r <= 2 + filasFiltradas.length; r += 1) {
    for (let c = 1; c <= config.columnas.length; c += 1) {
      const cell = ws.getCell(r, c);
      cell.alignment = { ...cell.alignment, wrapText: true, vertical: 'top' };
    }
  }

  // Alto de fila explícito: Excel calcula esto solo al abrir el archivo, pero al exportar
  // a PDF/imagen (LibreOffice, para la weekly de crew) NO lo recalcula — si no se fija a
  // mano, el texto envuelto se ve encimado con la fila de abajo. Se estima cuántas líneas
  // necesita la celda con más contenido de cada fila, según el ancho ya asignado a su columna.
  const ALTO_POR_LINEA = 14; // pt, aprox para fuente 10-11
  for (let r = 2; r <= 2 + filasFiltradas.length; r += 1) {
    let maxLineas = 1;
    for (let c = 1; c <= config.columnas.length; c += 1) {
      const texto = String(ws.getCell(r, c).value || '');
      const anchoColumna = ws.getColumn(c).width || 12;
      const charsPorLinea = Math.max(anchoColumna * 1.1, 5); // estimado conservador de chars visibles por línea
      const lineasPorSaltos = texto.split('\n').length;
      const lineasPorAncho = Math.ceil(texto.length / charsPorLinea);
      maxLineas = Math.max(maxLineas, lineasPorSaltos, lineasPorAncho);
    }
    ws.getRow(r).height = maxLineas * ALTO_POR_LINEA + 10;
  }

  // Formato de tabla: bordes en todas las celdas con datos, filtros en el encabezado y
  // la fila de encabezado congelada — para que se lea como tabla real y no como texto
  // suelto, sin perder los colores de las reglas de negocio ya aplicados arriba.
  const totalFilas = filasFiltradas.length;
  const totalColumnas = config.columnas.length;
  const BORDE_FINO = { style: 'thin', color: { argb: 'FFB7B7B7' } };
  for (let r = 2; r <= 2 + totalFilas; r += 1) {
    for (let c = 1; c <= totalColumnas; c += 1) {
      ws.getCell(r, c).border = { top: BORDE_FINO, bottom: BORDE_FINO, left: BORDE_FINO, right: BORDE_FINO };
    }
  }
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2 + totalFilas, column: totalColumnas } };

  // Congelar columnas desde CLIENT/CLIENTE hasta "NOB Day for Truck to Arrive" (o la fecha
  // principal de esa weekly si esa columna no existe en este tipo), igual que en la
  // plantilla original — así esas columnas quedan visibles siempre al hacer scroll lateral.
  const colCongelarHasta = config.columnas.findIndex(
    (c) => c.label === 'NOB Day for Truck to Arrive' || c.esFechaPrincipal
  );
  const xSplit = colCongelarHasta >= 0 ? colCongelarHasta + 1 : 4;
  ws.views = [{ state: 'frozen', xSplit, ySplit: 2 }];

  // Configuración de página: horizontal y ajustado al ancho, para que todas las columnas
  // quepan de lado a lado sin cortarse al exportar a PDF/imagen.
  ws.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0, // deja que el alto fluya a varias páginas si hace falta, sin achicar el ancho
    margins: { left: 0.3, right: 0.3, top: 0.3, bottom: 0.3, header: 0.2, footer: 0.2 },
  };

  if (!require('fs').existsSync(OUTPUT_DIR)) require('fs').mkdirSync(OUTPUT_DIR, { recursive: true });
  // Formato confirmado: "Weeklies - [para quien son] - [fecha de generación].xlsx"
  const fechaGeneracion = fechaWeeklies.toLocaleDateString('en-US').replace(/\//g, '-');
  const nombreArchivo = `Weeklies - ${filtroValor} - ${fechaGeneracion}.xlsx`;
  const outPath = path.join(OUTPUT_DIR, nombreArchivo);
  await workbook.xlsx.writeFile(outPath);
  return { outPath, totalFilas: filasFiltradas.length };
}

/**
 * Aplica la regla de color correspondiente a una celda, con la variante clara + tachado
 * si la fecha es pasada.
 */
function aplicarColor(cell, colDef, valor, mismoProveedor, colorBanda) {
  const pendiente = colDef.operativa && esTextoPendiente(valor);
  const pasada = esFechaPasada(valor);

  let colorBase = null;
  if (pendiente) {
    colorBase = COLORES.AMARILLO_PENDIENTE;
  } else if (colDef.esFechaPrincipal) {
    colorBase = COLORES.VERDE_FECHA_PRINCIPAL;
  } else if (colDef.esFechaLlegadaMX) {
    colorBase = COLORES.AZUL_LLEGADA_MEXICO;
  }

  // Naranja: mismo proveedor en Customs Broker / NOB Transportation Company (solo esas 2 celdas)
  if (!colorBase && mismoProveedor && (colDef.col === 'Customs Broker' || colDef.col === 'NOB Transportation Company')) {
    colorBase = COLORES.NARANJA_MISMO_PROVEEDOR;
  }

  // Amarillo: TAGS que contengan "NORTHBOUND"
  if (!colorBase && colDef.col === 'TAGS' && String(valor || '').toUpperCase().includes('NORTHBOUND')) {
    colorBase = COLORES.NARANJA_MISMO_PROVEEDOR;
  }

  if (colorBase && pasada) {
    const CLAROS = {
      [COLORES.AMARILLO_PENDIENTE]: COLORES.AMARILLO_PENDIENTE_CLARO,
      [COLORES.VERDE_FECHA_PRINCIPAL]: COLORES.VERDE_CLARO,
      [COLORES.AZUL_LLEGADA_MEXICO]: COLORES.AZUL_CLARO,
      [COLORES.NARANJA_MISMO_PROVEEDOR]: COLORES.NARANJA_CLARO,
    };
    colorBase = CLAROS[colorBase] || colorBase;
    cell.font = { ...cell.font, strike: true };
  }

  if (colorBase) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colorBase } };
  } else if (colorBanda) {
    // Sin ninguna regla especial en esta celda: usa el bandeado alterno de la fila.
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colorBanda } };
  }

  // Notas: texto rojo solo si hay contenido real (no pendiente)
  if (colDef.esNota && !pendiente && String(valor || '').trim() !== '') {
    cell.font = { ...cell.font, color: { argb: COLORES.TEXTO_ROJO_NOTAS } };
  }
}

/**
 * Para las weeklies que se mandan como imagen (crew), convierte el .xlsx generado a PNG
 * usando LibreOffice headless (mismo patrón que ya usamos para el BOL).
 */
async function generateWeeklyImage(tipo, filtroValor, sourceFilePath, fechaWeeklies = new Date()) {
  const { execFile } = require('child_process');
  const { outPath, totalFilas } = await generateWeekly(tipo, filtroValor, sourceFilePath, fechaWeeklies);

  const ejecutar = (cmd, args) => new Promise((resolve, reject) => {
    execFile(cmd, args, (error, stdout, stderr) => {
      if (error) return reject(new Error(stderr || error.message));
      resolve(stdout);
    });
  });

  await ejecutar('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', OUTPUT_DIR, outPath]);
  const pdfPath = outPath.replace(/\.xlsx$/, '.pdf');

  // pdftoppm en vez de la conversión PNG por defecto de LibreOffice: permite fijar la
  // resolución (200 DPI, el doble de lo normal) para que no se vea pixeleada al hacer zoom.
  const pngBase = pdfPath.replace(/\.pdf$/, '');
  await ejecutar('pdftoppm', ['-png', '-r', '200', '-singlefile', pdfPath, pngBase]);
  const pngPath = `${pngBase}.png`;

  return { pngPath, xlsxPath: outPath, totalFilas };
}

/**
 * Valores que existen en los datos pero NUNCA deben ofrecerse como weekly real.
 */
const EXCLUSIONES_POR_TIPO = {
  transportista: ['estes', 'old dominion'], // confirmado: no reciben weekly
  crew: ['client'], // "Client", "Client will have unloading crew", etc. no son crew real
};

/**
 * Lista los valores únicos de filtro disponibles para un tipo de weekly (ej. todos los
 * transportistas con al menos 1 embarque activo), para poblar un selector en el formulario.
 */
async function listarFiltrosDisponibles(tipo, sourceFilePath) {
  const config = WEEKLY_TIPOS[tipo];
  if (!config) throw new Error(`Tipo de weekly desconocido: "${tipo}"`);

  // Tipos sin filtro (ej. la weekly maestra de SOB Crew para Emilia): no hay nada que
  // elegir, siempre es "una sola opción" con el nombre fijo que trae la configuración.
  if (config.sinFiltro) return [config.nombreFijo || tipo];

  const exclusiones = EXCLUSIONES_POR_TIPO[tipo] || [];
  const filas = await leerPasteHere(sourceFilePath);
  const valores = new Map(); // normalizado -> texto original (primera aparición)
  filas.forEach((f) => {
    if (normalizar(f['TAGS']).includes('on hold')) return; // no cuenta para la lista de opciones
    const crudo = f[config.filtroColumna];
    if (!crudo) return;
    const limpio = String(crudo).trim();
    if (!limpio) return;
    const norm = normalizar(limpio);
    if (exclusiones.some((ex) => norm.includes(ex))) return;
    if (!valores.has(norm)) valores.set(norm, limpio);
  });
  return Array.from(valores.values()).sort();
}

module.exports = { generateWeekly, generateWeeklyImage, listarFiltrosDisponibles, leerPasteHere, normalizar };
