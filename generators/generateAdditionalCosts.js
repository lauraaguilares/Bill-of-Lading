/**
 * generateAdditionalCosts.js
 * Genera "Understanding Additional Costs for Larger Shipments" (Southbound, 28ft) a partir
 * de los datos extraídos del contrato firmado, replicando el diseño visual exacto de la
 * plantilla original (logo, colores, imágenes de camiones, 2 páginas).
 */

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { execFile } = require('child_process');
const { parsearContrato } = require('./contractParser');
const { BRACKETS_MEXICO, TRAILER_US, bracketParaCubicFt, precioParaBracket, tienePrecioExplicito } = require('../config/additionalCostsConfig');

const OUTPUT_DIR = path.join(__dirname, '..', 'output');
const ASSETS_DIR = path.join(__dirname, '..', 'assets');

// Colores de la referencia real que Laura compartió: azul fuerte para los encabezados de
// grupo, azul más claro para los sub-encabezados (ambos con texto blanco), y un azul muy
// claro para las rayas alternadas de las filas de datos.
const AZUL_OSCURO = 'FF6FA8DC'; // "28-Foot Trailer in the US", "Additional Costs...", "Total Additional Price"
const AZUL_MEDIO = 'FF9FC5E8'; // "Linear Feet You Use in the US", "Type of Truck Needed in Mexico", etc.
const AZUL_CLARO = 'FFD9E2F3'; // rayas alternadas en las filas de datos
const AZUL_MUY_CLARO = 'FFDEEBF7'; // fondo de "What Your Agreement Says"
const BORDE = { style: 'thin', color: { argb: 'FF000000' } };

// Qué imagen de camión corresponde a cada bracket (mapeo confirmado directo de la
// plantilla original: se revisaron los anclajes de imagen en el XML del xlsx), con su
// proporción real (ancho/alto) para no distorsionarla al insertarla — el tráiler de 53ft
// es mucho más ancho/plano (5:1) que los camiones tipo caja (~2:1); forzar el mismo tamaño
// a todos hacía que se vieran casi idénticos.
const ANCHO_ICONO = 95;
// Formato contabilidad: "$" pegado a la izquierda, número a la derecha, "-" en vez de "$0"
// (así se ve en la plantilla real) — se usa en las columnas "Additional Price for Trailer"
// y "Total", pero NO en "Additional Price for Truck in Mexico" (esa se queda simple).
const FORMATO_CONTABILIDAD = '_-$* #,##0_-;-$* #,##0_-;_-$* "-"_-;_-@_-';
const IMAGEN_POR_BRACKET_HASTA = {
  600: { path: path.join(ASSETS_DIR, 'trucks', 'small.png'), ratio: 223 / 125 },
  1000: { path: path.join(ASSETS_DIR, 'trucks', 'medium.png'), ratio: 381 / 159 },
  1450: { path: path.join(ASSETS_DIR, 'trucks', 'large.png'), ratio: 404 / 167 },
  2400: { path: path.join(ASSETS_DIR, 'trucks', 'trailer53.png'), ratio: 520 / 102 },
  3400: { path: path.join(ASSETS_DIR, 'trucks', 'trailer53.png'), ratio: 520 / 102 }, // mismo tráiler, bracket más alto
};

function generarTextoIntro(paisOrigenLabel) {
  return `At Best Mexico Movers, we want to make sure you understand your costs if your shipment exceeds the amounts in your agreement. That's why we made the chart on the next page.

Your precious household goods will be transported via two separate trucks; one in ${paisOrigenLabel} and one in Mexico. The chart on the next page will explain how your costs change if you bring more than in our agreement. After looking at this chart, you may decide to bring only what we originally contracted for (in which case your price will not change at all), or you may decide to bring more. The choice is yours. If you are considering bringing more, this chart will help you to see exactly how the size of your shipment will affect your price.

It is very important to us that we explain this well. If you have any questions whatsoever, please ask your Personal Moving Assistant. Once you do understand it, please sign on the second page at the bottom and return to your PMA.`;
}

/**
 * Calcula la tabla completa fila por fila, desde los pies lineales contratados por el
 * cliente hasta la capacidad máxima del trailer.
 */
function calcularTabla(datos) {
  const trailer = TRAILER_US[datos.trailerSizeUS];
  if (!trailer) {
    throw new Error(
      `Trailer de ${datos.trailerSizeUS} ft todavía no está configurado (solo 26/28/53ft por ahora). ` +
      `Avísame para agregar este tamaño.`
    );
  }

  if (trailer.modoResumen) return calcularTablaResumen(datos, trailer);

  // Para camiones con tope físico fijo (26ft): la tabla empieza en el pie lineal que
  // corresponde al Destination Amount del cliente (confirmado con 2 ejemplos reales:
  // Hartman con 600cf empieza en el pie 9, Coffeng con 1,000cf empieza en el pie 15 —
  // ambos coinciden exactamente con floor(destinationAmount / cf-por-pie)).
  // Para camiones sin tope (28ft): la tabla empieza en el punto específico del cliente
  // (Contracted Linear Feet), ya que ahí sí se negocia una cantidad parcial.
  const pieLinealInicio = trailer.maxCapacidadFisica != null
    ? Math.floor(datos.destinationAmount / trailer.cubicFtPorPieLineal)
    : datos.contractedLinearFeet;

  const filas = [];
  for (let pieLineal = pieLinealInicio; pieLineal <= trailer.maxPiesLineales; pieLineal += 1) {
    const cubicFt = pieLineal * trailer.cubicFtPorPieLineal;
    // Si el camión tiene un tope físico fijo (ej. 26ft topa en 21 pies), lo que exceda ese
    // tope no es que cueste más — literalmente no se puede usar ese camión.
    const excedeCapacidadFisica = trailer.maxCapacidadFisica != null && pieLineal > trailer.maxCapacidadFisica;
    const precioTrailerUS = excedeCapacidadFisica ? null : (pieLineal - datos.contractedLinearFeet) * datos.precioPorPieAdicional;
    const bracket = bracketParaCubicFt(cubicFt);
    const precioTruckMX = precioParaBracket(bracket, datos.preciosMexico);
    filas.push({
      pieLineal,
      cubicFt,
      precioTrailerUS,
      excedeCapacidadFisica,
      bracketLabel: bracket.label,
      bracketHasta: bracket.hasta,
      bracketConPrecio: tienePrecioExplicito(bracket, datos.preciosMexico),
      precioTruckMX,
      total: (precioTrailerUS || 0) + precioTruckMX,
    });
  }
  return filas;
}

/**
 * Cálculo para trailers en "modo resumen" (53ft): solo se muestra 1 fila por cada punto
 * donde cambia el bracket de México (no 1 fila por cada pie lineal individual), empezando
 * en el primer punto de quiebre que sea >= lo que el cliente ya tiene contratado.
 */
function calcularTablaResumen(datos, trailer) {
  const filas = [];
  // Confirmado: la tabla empieza en el primer punto de quiebre que sea >= al Destination
  // Amount del cliente (igual que 26ft) — no siempre desde el primer punto de todos.
  let pieLinealAnterior = 0;
  trailer.puntosDeQuiebre
    .filter((punto) => punto.cubicFt >= datos.destinationAmount)
    .forEach((punto) => {
      const precioTrailerUS = Math.max(0, punto.pieLineal - datos.contractedLinearFeet) * datos.precioPorPieAdicional;
      const bracket = bracketParaCubicFt(punto.cubicFt);
      const precioTruckMX = precioParaBracket(bracket, datos.preciosMexico);
      // Cuánto "espacio" relativo representa este punto dentro del tráiler — para que,
      // ej. "Up to 32" (que cubre del 0 al 32) se vea más grande que "Up to 47" (que solo
      // cubre del 32 al 47), proporcional al tramo real que cada uno representa.
      const alturaRelativa = punto.pieLineal - pieLinealAnterior;
      pieLinealAnterior = punto.pieLineal;
      filas.push({
        pieLineal: `Up to ${punto.pieLineal}`,
        cubicFt: `Up to ${punto.cubicFt.toLocaleString('en-US')}`,
        precioTrailerUS,
        excedeCapacidadFisica: false,
        alturaRelativa,
        bracketLabel: bracket.label,
        bracketHasta: bracket.hasta,
        bracketConPrecio: tienePrecioExplicito(bracket, datos.preciosMexico),
        precioTruckMX,
        total: precioTrailerUS + precioTruckMX,
      });
    });

  // Fila final "Over 48 / CAN'T USE" — confirmado que siempre se muestra, más allá de la
  // capacidad física máxima del tráiler de 53ft.
  const maxPunto = trailer.puntosDeQuiebre[trailer.puntosDeQuiebre.length - 1];
  filas.push({
    pieLineal: `Over ${maxPunto.pieLineal + 1}`,
    cubicFt: `Over ${maxPunto.cubicFt.toLocaleString('en-US')}`,
    precioTrailerUS: null,
    excedeCapacidadFisica: true,
    alturaRelativa: 1,
    bracketLabel: "CAN'T USE",
    bracketHasta: 'no-bracket', // no hay bracket real de México más allá del máximo
    bracketConPrecio: false,
    precioTruckMX: null,
    total: null,
  });

  return filas;
}

function barraEncabezado(ws, rango, texto) {
  ws.mergeCells(rango);
  const cell = ws.getCell(rango.split(':')[0]);
  cell.value = texto;
  cell.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_OSCURO } };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  return cell;
}

function ponerBordes(ws, filaInicio, filaFin, colInicio, colFin) {
  for (let r = filaInicio; r <= filaFin; r += 1) {
    for (let c = colInicio; c <= colFin; c += 1) {
      ws.getCell(r, c).border = { top: BORDE, bottom: BORDE, left: BORDE, right: BORDE };
    }
  }
}

/**
 * Dibuja una "caja" con borde exterior completo alrededor de un rango de columnas, más
 * líneas divisorias verticales internas en las columnas que se indiquen — así se ve como
 * una sola sección con separadores, en vez de una cuadrícula pareja en cada celda (que es
 * lo que se veía "poco estético": la imagen del camión quedaba fuera del recuadro porque
 * antes el borde no incluía su columna).
 */
/**
 * Convierte un desplazamiento en píxeles (desde la parte de arriba de filaInicio) a la
 * posición fraccional que ExcelJS necesita para anclar una imagen — contemplando que las
 * filas pueden tener alturas distintas entre sí (no todas 22px), si no el cálculo se
 * desfasa cuanto más alta sea la fila real.
 */
function offsetPixelesAFilaAnchor(ws, filaInicio, offsetPx) {
  let acumulado = 0;
  let r = filaInicio;
  for (let i = 0; i < 500; i += 1) {
    const alturaFila = ws.getRow(r).height || 15;
    if (acumulado + alturaFila >= offsetPx) {
      const fraccion = Math.max(0, (offsetPx - acumulado) / alturaFila);
      return (r - 1) + fraccion;
    }
    acumulado += alturaFila;
    r += 1;
  }
  return filaInicio - 1;
}

function cajaConDivisores(ws, filaInicio, filaFin, colInicio, colFin, colsDivisorDerecha = []) {
  for (let r = filaInicio; r <= filaFin; r += 1) {
    for (let c = colInicio; c <= colFin; c += 1) {
      const actual = ws.getCell(r, c).border || {};
      const border = { ...actual };
      if (r === filaInicio) border.top = BORDE;
      if (r === filaFin) border.bottom = BORDE;
      if (c === colInicio) border.left = BORDE;
      if (c === colFin) border.right = BORDE;
      if (colsDivisorDerecha.includes(c)) border.right = BORDE;
      if (Object.keys(border).length) ws.getCell(r, c).border = border;
    }
  }
}

/**
 * Línea horizontal divisoria entre bloques (ej. entre "Small Straight Truck" y "Medium
 * Straight Truck"), sin tocar los bordes exteriores de la caja.
 */
function divisorHorizontal(ws, fila, colInicio, colFin) {
  for (let c = colInicio; c <= colFin; c += 1) {
    const actual = ws.getCell(fila, c).border || {};
    ws.getCell(fila, c).border = { ...actual, bottom: BORDE };
  }
}

function ponerBordesTabla(ws, filaInicio, filaFin) {
  cajaConDivisores(ws, filaInicio, filaFin, 1, 3, [1, 2]); // A|B|C con divisores internos
  cajaConDivisores(ws, filaInicio, filaFin, 5, 7, [6]); // E-F-G: una sola caja, divisor entre F (texto) y G (precio)
  cajaConDivisores(ws, filaInicio, filaFin, 9, 9, []); // I
}

async function generateAdditionalCosts(contractPdfPath) {
  const datos = await parsearContrato(contractPdfPath);

  const esSouthbound = datos.paisOrigen === 'US' && datos.paisDestino === 'MX';
  const esBonded = datos.paisOrigen === 'CA' && datos.paisDestino === 'MX';
  if (!esSouthbound && !esBonded) {
    throw new Error(
      `Este generador por ahora solo soporta Southbound (EE.UU. → México) y Bonded (Canadá → ` +
      `México). Este contrato detectó origen=${datos.paisOrigen}, destino=${datos.paisDestino}.`
    );
  }
  // Confirmado con la plantilla real: Bonded dice "in Canada" en vez de "in the US" en
  // todos los encabezados y en el texto introductorio — es la única diferencia real.
  const paisOrigenLabel = esBonded ? 'Canada' : 'the US';

  const filas = calcularTabla(datos);
  const trailerConfig = TRAILER_US[datos.trailerSizeUS];

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Costos Adicionales');
  ws.pageSetup = {
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    verticalCentered: false,
    margins: { left: 0.6, right: 0.6, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 },
  };
  ws.headerFooter.oddFooter = '&8&KFF0000Confidential and Proprietary.  Do not use without prior written consent.';

  ws.getColumn('A').width = 11;
  ws.getColumn('B').width = 12;
  ws.getColumn('C').width = 13;
  ws.getColumn('D').width = 4;
  ws.getColumn('E').width = 15;
  ws.getColumn('F').width = 20;
  ws.getColumn('G').width = 14;
  ws.getColumn('H').width = 4;
  ws.getColumn('I').width = 13;

  const logoId = workbook.addImage({ filename: path.join(ASSETS_DIR, 'logo', 'bmm-logo.png'), extension: 'png' });
  ws.addImage(logoId, { tl: { col: 2.3, row: 0.2 }, ext: { width: 260, height: 80 } });
  ws.getRow(1).height = 70;

  let fila = 4;
  ws.mergeCells(`A${fila}:I${fila}`);
  ws.getCell(`A${fila}`).value = 'Understanding Additional Costs for Larger Shipments';
  ws.getCell(`A${fila}`).font = { bold: true, size: 15 };
  ws.getCell(`A${fila}`).alignment = { horizontal: 'center' };
  fila += 2;

  ws.mergeCells(`A${fila}:I${fila + 4}`);
  ws.getCell(`A${fila}`).value = generarTextoIntro(paisOrigenLabel);
  ws.getCell(`A${fila}`).alignment = { wrapText: true, vertical: 'top' };
  ws.getRow(fila).height = 140;
  fila += 6;

  ws.getRow(fila - 1).addPageBreak();
  fila += 1;

  ws.mergeCells(`A${fila}:I${fila}`);
  ws.getCell(`A${fila}`).value = 'Calculating the Price for Your Shipment if You Bring More';
  ws.getCell(`A${fila}`).font = { bold: true, size: 16 };
  ws.getCell(`A${fila}`).alignment = { horizontal: 'center' };
  fila += 2;

  // "What Your Agreement Says": el contenido varía según el tamaño de trailer — 28ft
  // muestra el desglose de 3 líneas; 26ft y 53ft muestran solo "Base Cost on Agreement"
  // (el precio total del contrato, Sección 9.A).
  {
    const filaAgreementBar = fila;
    barraEncabezado(ws, `A${fila}:I${fila}`, 'What Your Agreement Says');
    fila += 1;

    const resumenAgreement = (datos.trailerSizeUS === 26 || datos.trailerSizeUS === 53)
      ? [
          ['Base Cost on Agreement', datos.precioTotalContrato, '"$"#,##0'],
        ]
      : [
          ['How Many Linear Feet in the US with No Additional Cost', datos.contractedLinearFeet, '#,##0'],
          ['How Many Cubic Feet in the US with No Additional Cost', datos.origenAmount, '#,##0'],
          ['Price for Each Additional Linear Foot in the US if You Use More Than in the Agreement', datos.precioPorPieAdicional, '"$"#,##0'],
        ];
    resumenAgreement.forEach(([label, valor, formato]) => {
      ws.mergeCells(`A${fila}:G${fila}`);
      ws.getCell(`A${fila}`).value = label;
      ws.getCell(`A${fila}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_MUY_CLARO } };
      ws.mergeCells(`H${fila}:I${fila}`);
      ws.getCell(`H${fila}`).value = valor;
      ws.getCell(`H${fila}`).numFmt = formato;
      ws.getCell(`H${fila}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_MUY_CLARO } };
      ws.getCell(`H${fila}`).alignment = { horizontal: 'right' };
      fila += 1;
    });
    ponerBordes(ws, filaAgreementBar, fila - 1, 1, 9);
    fila += 2;
  }

  const filaGrupoHeader = fila;
  ws.mergeCells(`A${fila}:C${fila}`);
  ws.getCell(`A${fila}`).value = `${TRAILER_US[datos.trailerSizeUS].labelGrupo} in ${paisOrigenLabel}`;
  ws.mergeCells(`E${fila}:G${fila}`);
  ws.getCell(`E${fila}`).value = 'Additional Costs for Larger Shipments in Mexico';
  [`A${fila}`, `E${fila}`].forEach((c) => {
    ws.getCell(c).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    ws.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_OSCURO } };
    ws.getCell(c).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true, indent: 1 };
  });
  ws.getRow(fila).height = 46;
  fila += 1;

  const filaSubEncabezado = fila;

  // "Total Additional Price" abarca AMBAS filas de encabezado (la de grupo y la de
  // sub-encabezado) para que su borde inferior empate con el de los sub-encabezados de
  // las otras dos tablas — si no, su bloque de datos empieza una fila más arriba que las
  // demás columnas.
  ws.mergeCells(`I${filaGrupoHeader}:I${filaSubEncabezado}`);
  ws.getCell(`I${filaGrupoHeader}`).value = 'Total Additional Price';
  ws.getCell(`I${filaGrupoHeader}`).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
  ws.getCell(`I${filaGrupoHeader}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_OSCURO } };
  ws.getCell(`I${filaGrupoHeader}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true, indent: 1 };

  // "Type of Truck Needed in Mexico" abarca la columna de la imagen (E) Y la del texto (F)
  // como UN solo sub-encabezado — antes solo cubría F y se veía descuadrado con el bloque
  // de abajo (imagen + texto juntos).
  ws.mergeCells(`E${fila}:F${fila}`);
  ws.getCell(`E${fila}`).value = 'Type of Truck Needed in Mexico';

  const subEncabezados = { A: `Linear Feet You Use in ${paisOrigenLabel}`, B: `Cubic Feet You Use in ${paisOrigenLabel}`, C: `${TRAILER_US[datos.trailerSizeUS].labelPrecio} in ${paisOrigenLabel}`, G: 'Additional Price for Truck in Mexico' };
  Object.entries(subEncabezados).forEach(([col, texto]) => {
    ws.getCell(`${col}${fila}`).value = texto;
  });
  [`A${fila}`, `B${fila}`, `C${fila}`, `E${fila}`, `G${fila}`].forEach((c) => {
    const cell = ws.getCell(c);
    cell.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_MEDIO } };
    cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center', indent: 1 };
  });
  ws.getRow(fila).height = 55;
  fila += 1;

  // Agrupar filas consecutivas por bracket — cada tipo de camión ocupa un bloque de varias
  // filas (según cuántos pies lineales caen en ese rango). Dentro del bloque, la imagen, el
  // texto y el precio van UNA sola vez, centrados verticalmente — tal como en la plantilla
  // original (no repetidos fila por fila).
  const filaInicioTabla = fila;
  const grupos = [];
  const divisoresEntreBloques = [];
  filas.forEach((f) => {
    const grupoActual = grupos[grupos.length - 1];
    if (!grupoActual || grupoActual.bracketHasta !== f.bracketHasta) {
      grupos.push({ bracketHasta: f.bracketHasta, bracketLabel: f.bracketLabel, bracketConPrecio: f.bracketConPrecio, precioTruckMX: f.precioTruckMX, filas: [f] });
    } else {
      grupoActual.filas.push(f);
    }
  });

  // Alturas de fila: por default reparte el extra de altura equitativamente dentro de cada
  // grupo. Para 53ft (modoResumen), en vez de eso se calcula proporcional al tramo real que
  // representa cada punto ("Up to 32" cubre más rango que "Up to 47", así que se ve más
  // grande) — agrupando primero los brackets consecutivos que comparten la MISMA imagen
  // (ej. las 2 secciones de "53-Foot Trailer") para reservar el alto conjunto entre ellos.
  if (trailerConfig.modoResumen) {
    const familias = [];
    grupos.forEach((grupo) => {
      const imgPath = IMAGEN_POR_BRACKET_HASTA[grupo.bracketHasta]?.path || null;
      const anterior = familias[familias.length - 1];
      if (anterior && anterior.imgPath === imgPath && imgPath != null) {
        anterior.grupos.push(grupo);
      } else {
        familias.push({ imgPath, grupos: [grupo] });
      }
    });
    familias.forEach(({ imgPath, grupos: gruposFamilia }) => {
      const todasLasFilas = gruposFamilia.flatMap((g) => g.filas);
      const ratio = imgPath ? IMAGEN_POR_BRACKET_HASTA[gruposFamilia[0].bracketHasta].ratio : null;
      const alturaIcono = ratio ? Math.round(ANCHO_ICONO / ratio) : 0;
      const totalSpan = todasLasFilas.reduce((s, f) => s + (f.alturaRelativa || 1), 0);
      // El alto total de la familia se reparte PROPORCIONAL al tramo real que representa
      // cada fila (no solo el "extra" sobre un mínimo) — si no, con solo 2 filas casi
      // siempre el mínimo ya alcanza y la proporción nunca se nota visualmente.
      const alturaPorProporcion = totalSpan * 6; // px por cada pie lineal de tramo
      const alturaMinimaTotal = Math.max(alturaIcono + 30, alturaPorProporcion, todasLasFilas.length * 22);
      todasLasFilas.forEach((f) => {
        const peso = (f.alturaRelativa || 1) / totalSpan;
        f._alturaFila = Math.max(22, alturaMinimaTotal * peso);
      });
    });
  }

  grupos.forEach((grupo) => {
    const filaInicioGrupo = fila;
    const alturaIconoEstimeda = Math.round(ANCHO_ICONO / (IMAGEN_POR_BRACKET_HASTA[grupo.bracketHasta]?.ratio || 2));
    // Alto mínimo del bloque para que la imagen quepa cómoda, aunque el bloque sea de 1 sola
    // fila de datos (ej. "Large Straight Truck" cuando el cliente ya arranca justo en ese
    // bracket) — si no, la imagen se desborda del bloque.
    const alturaMinimaBloque = alturaIconoEstimeda + 30;
    const filasEnGrupo = grupo.filas.length;
    const alturaPorFilaNormal = 22;
    const alturaTotalNormal = filasEnGrupo * alturaPorFilaNormal;
    const alturaExtra = Math.max(0, alturaMinimaBloque - alturaTotalNormal);

    grupo.filas.forEach((f, idx) => {
      ws.getCell(`A${fila}`).value = f.pieLineal;
      ws.getCell(`B${fila}`).value = f.cubicFt;
      ws.getCell(`B${fila}`).numFmt = '#,##0';
      // La columna C (precio del trailer en EE.UU.) se llena después, en un solo bloque
      // fusionado por cada tramo de "$0" o "CAN'T USE" — ver más abajo.
      [`A${fila}`, `B${fila}`].forEach((c) => {
        ws.getCell(c).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.getCell(c).font = { size: 11 };
      });

      // Rayas alternadas, igual que la plantilla original.
      if (fila % 2 === 0) {
        [`A${fila}`, `B${fila}`, `C${fila}`, `I${fila}`].forEach((c) => {
          ws.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_CLARO } };
        });
      }
      // El extra de altura se reparte entre las filas del bloque para que se vea proporcional
      // (o, para 53ft, se usa la altura ya calculada arriba, proporcional al tramo real).
      ws.getRow(fila).height = f._alturaFila != null ? f._alturaFila : (alturaPorFilaNormal + alturaExtra / filasEnGrupo);
      // Línea delgada entre cada fila individual (además de la raya de color) — así se ve
      // igual que la referencia, no solo un bloque de color liso.
      if (idx < grupo.filas.length - 1 || grupos.indexOf(grupo) < grupos.length - 1 || f !== filas[filas.length - 1]) {
        [`A${fila}`, `B${fila}`].forEach((c) => {
          const actual = ws.getCell(c).border || {};
          ws.getCell(c).border = { ...actual, bottom: { style: 'thin', color: { argb: 'FFBDD7EE' } } };
        });
      }
      fila += 1;
    });
    const filaFinGrupo = fila - 1;
    const esFilaCantUse = grupo.bracketHasta === 'no-bracket';
    grupo._filaInicioGrupo = filaInicioGrupo;
    grupo._filaFinGrupo = filaFinGrupo;
    grupo._esFilaCantUse = esFilaCantUse;

    if (!esFilaCantUse) {
      ws.mergeCells(`G${filaInicioGrupo}:G${filaFinGrupo}`);
      ws.getCell(`G${filaInicioGrupo}`).value = grupo.precioTruckMX;
      ws.getCell(`G${filaInicioGrupo}`).numFmt = '$#,##0';
      ws.getCell(`G${filaInicioGrupo}`).alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getCell(`G${filaInicioGrupo}`).font = { size: 13 };
    }

    // Total: en camiones con tope físico fijo (26ft/53ft), el precio en EE.UU. siempre es
    // $0, así que el Total dentro de un mismo bracket de México es idéntico en todas sus
    // filas — se fusiona en un solo bloque igual que la columna del camión. En camiones
    // sin tope (28ft), el precio en EE.UU. sí escala fila por fila, así que el Total se
    // deja individual.
    if (trailerConfig.maxCapacidadFisica != null) {
      ws.mergeCells(`I${filaInicioGrupo}:I${filaFinGrupo}`);
      if (esFilaCantUse) {
        ws.getCell(`I${filaInicioGrupo}`).value = "CAN'T USE";
        ws.getCell(`I${filaInicioGrupo}`).font = { size: 12, bold: true, color: { argb: 'FFFF0000' } };
      } else {
        ws.getCell(`I${filaInicioGrupo}`).value = grupo.filas[0].total;
        ws.getCell(`I${filaInicioGrupo}`).numFmt = FORMATO_CONTABILIDAD;
        ws.getCell(`I${filaInicioGrupo}`).font = { size: 11 };
      }
      ws.getCell(`I${filaInicioGrupo}`).alignment = { horizontal: 'center', vertical: 'middle' };
      if (grupos.indexOf(grupo) % 2 === 1) {
        ws.getCell(`I${filaInicioGrupo}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_CLARO } };
      }
    } else {
      grupo.filas.forEach((f, idx) => {
        const filaActual = filaInicioGrupo + idx;
        const cell = ws.getCell(`I${filaActual}`);
        cell.value = f.total;
        cell.numFmt = FORMATO_CONTABILIDAD;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { size: 11 };
      });
    }

    // Línea divisoria entre bloques de tipo de camión (como en el original) — se dibuja
    // DESPUÉS de la caja completa (ver ponerBordesTabla), para no perderse al fusionar celdas.
    if (grupos.indexOf(grupo) < grupos.length - 1) {
      divisoresEntreBloques.push(filaFinGrupo);
    }
  });

  // Imagen + etiqueta (columnas E/F): se dibuja UNA sola vez por "familia" de brackets
  // consecutivos que comparten la misma imagen (ej. "53-Foot Trailer" cubriendo tanto "Up
  // to 32" como "Up to 47"), en vez de repetir el mismo camión dos veces — centrada en todo
  // el alto combinado de esa familia.
  const familiasImagen = [];
  grupos.forEach((grupo) => {
    const imgPath = IMAGEN_POR_BRACKET_HASTA[grupo.bracketHasta]?.path || null;
    const esFilaCantUse = grupo._esFilaCantUse;
    const anterior = familiasImagen[familiasImagen.length - 1];
    if (!esFilaCantUse && anterior && anterior.imgPath === imgPath && imgPath != null) {
      anterior.grupos.push(grupo);
    } else {
      familiasImagen.push({ imgPath, esFilaCantUse, grupos: [grupo] });
    }
  });
  familiasImagen.forEach(({ imgPath, esFilaCantUse, grupos: gruposFamilia }) => {
    const filaInicioFamilia = gruposFamilia[0]._filaInicioGrupo;
    const filaFinFamilia = gruposFamilia[gruposFamilia.length - 1]._filaFinGrupo;
    const esMultiple = gruposFamilia.length > 1;

    if (!esFilaCantUse && imgPath && fs.existsSync(imgPath)) {
      const ratio = IMAGEN_POR_BRACKET_HASTA[gruposFamilia[0].bracketHasta].ratio;
      const imgId = workbook.addImage({ filename: imgPath, extension: 'png' });
      const alto = Math.round(ANCHO_ICONO / ratio);
      let alturaBloquePx = 0;
      for (let r = filaInicioFamilia; r <= filaFinFamilia; r += 1) alturaBloquePx += ws.getRow(r).height;
      const offsetPx = Math.max(0, (alturaBloquePx - alto) / 2);
      const filaAnchor = offsetPixelesAFilaAnchor(ws, filaInicioFamilia, offsetPx);
      const offsetHorizontalCol = Math.max(0, (15 - ANCHO_ICONO / 7) / 2 / 15);
      ws.addImage(imgId, { tl: { col: 4 + offsetHorizontalCol, row: filaAnchor }, ext: { width: ANCHO_ICONO, height: alto } });
    }

    ws.mergeCells(`E${filaInicioFamilia}:E${filaFinFamilia}`);
    if (esFilaCantUse) {
      ws.mergeCells(`F${filaInicioFamilia}:G${filaFinFamilia}`);
    } else {
      ws.mergeCells(`F${filaInicioFamilia}:F${filaFinFamilia}`);
    }
    // Cuando la familia abarca más de un bracket (misma imagen, ej. 53ft), se usa la
    // etiqueta corta sin el rango de pies cúbicos, ya que ese rango varía entre los
    // brackets que comparte — el precio (columna G) sigue mostrando el valor correcto de
    // cada uno por separado.
    const etiqueta = esMultiple ? gruposFamilia[0].bracketLabel.split('\n')[0] : gruposFamilia[0].bracketLabel;
    ws.getCell(`F${filaInicioFamilia}`).value = etiqueta;
    ws.getCell(`F${filaInicioFamilia}`).alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
    ws.getCell(`F${filaInicioFamilia}`).font = esFilaCantUse
      ? { size: 14, bold: true, color: { argb: 'FFFF0000' } }
      : { size: 11 };
  });

  // "+" y "=" grandes, centrados verticalmente en toda la altura de la tabla (una sola vez,
  // no repetidos por fila).
  ws.mergeCells(`D${filaInicioTabla}:D${fila - 1}`);
  ws.getCell(`D${filaInicioTabla}`).value = '+';
  ws.getCell(`D${filaInicioTabla}`).font = { bold: true, size: 20 };
  ws.getCell(`D${filaInicioTabla}`).alignment = { horizontal: 'center', vertical: 'middle' };

  ws.mergeCells(`H${filaInicioTabla}:H${fila - 1}`);
  ws.getCell(`H${filaInicioTabla}`).value = '=';
  ws.getCell(`H${filaInicioTabla}`).font = { bold: true, size: 20 };
  ws.getCell(`H${filaInicioTabla}`).alignment = { horizontal: 'center', vertical: 'middle' };

  // Columna C: para camiones con tope físico fijo (26ft), se fusiona en UN bloque por tramo
  // (todo el tramo sin costo junto, y todo "CAN'T USE" junto) — igual que la referencia. Para
  // camiones sin tope (28ft), cada fila lleva su propio valor individual, porque sí cambia
  // fila por fila.
  if (trailerConfig.maxCapacidadFisica != null) {
    let filaTramoInicio = filaInicioTabla;
    for (let i = 0; i <= filas.length; i += 1) {
      const actual = filas[i];
      const anterior = filas[i - 1];
      const cambioDeTramo = !anterior || !actual || anterior.excedeCapacidadFisica !== actual.excedeCapacidadFisica;
      if (cambioDeTramo && anterior) {
        const filaTramoFin = filaInicioTabla + i - 1;
        ws.mergeCells(`C${filaTramoInicio}:C${filaTramoFin}`);
        const cell = ws.getCell(`C${filaTramoInicio}`);
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        if (anterior.excedeCapacidadFisica) {
          cell.value = "CAN'T USE";
          cell.font = { size: 12, bold: true, color: { argb: 'FFFF0000' } };
        } else {
          cell.value = 0;
          cell.numFmt = FORMATO_CONTABILIDAD;
          cell.font = { size: 11 };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
        // Línea divisoria entre el tramo de "$0" y el de "CAN'T USE" (si hay uno después).
        if (actual) {
          ws.getCell(`C${filaTramoFin}`).border = { ...(ws.getCell(`C${filaTramoFin}`).border || {}), bottom: BORDE };
        }
        filaTramoInicio = filaTramoFin + 1;
      }
    }
  } else {
    filas.forEach((f, i) => {
      const filaActual = filaInicioTabla + i;
      const cell = ws.getCell(`C${filaActual}`);
      cell.value = f.precioTrailerUS;
      cell.numFmt = FORMATO_CONTABILIDAD;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { size: 11 };
    });
  }

  ponerBordesTabla(ws, filaGrupoHeader, fila - 1);
  divisoresEntreBloques.forEach((filaDivisor) => divisorHorizontal(ws, filaDivisor, 5, 7));
  fila += 2;

  ws.mergeCells(`A${fila}:I${fila}`);
  ws.getCell(`A${fila}`).value = 'By signing this document, I am indicating that I have had the opportunity to ask questions and that I understand potential additional charges.';
  ws.getCell(`A${fila}`).alignment = { wrapText: true };
  ws.getRow(fila).height = 30;
  fila += 3;

  ws.getCell(`A${fila}`).value = '_______________________';
  fila += 1;
  ws.getCell(`A${fila}`).value = 'Signature';
  fila += 2;
  ws.getCell(`A${fila}`).value = '_______________________';
  fila += 1;
  ws.getCell(`A${fila}`).value = 'Date';

  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  const fechaGeneracion = new Date().toLocaleDateString('en-US').replace(/\//g, '-');
  const nombreBase = `${datos.clientName} - Additional Cost for Larger Shipments - ${fechaGeneracion}`;
  const xlsxPath = path.join(OUTPUT_DIR, `${nombreBase}.xlsx`);
  await workbook.xlsx.writeFile(xlsxPath);

  return { xlsxPath, datos, filas };
}

function ejecutar(cmd, args) {
  return new Promise((resolve, reject) => {
    execFile(cmd, args, (error, stdout, stderr) => {
      if (error) return reject(new Error(stderr || error.message));
      resolve(stdout);
    });
  });
}

async function generateAdditionalCostsPDF(contractPdfPath) {
  const { xlsxPath, datos, filas } = await generateAdditionalCosts(contractPdfPath);
  await ejecutar('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', OUTPUT_DIR, xlsxPath]);
  const pdfPath = xlsxPath.replace(/\.xlsx$/, '.pdf');
  return { pdfPath, xlsxPath, datos, filas };
}

module.exports = { generateAdditionalCosts, generateAdditionalCostsPDF, calcularTabla };
