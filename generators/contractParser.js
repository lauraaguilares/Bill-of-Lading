/**
 * contractParser.js
 * Extrae los datos clave del contrato (Moving Services Agreement / Acuerdo de Servicios de
 * Mudanza) que hacen falta para generar "Understanding Additional Costs for Larger
 * Shipments". Soporta inglés y español, y las 3 estructuras de contrato confirmadas:
 *
 *   28ft: menciona "a 28-foot trailer" / "un tráiler de 28 pies" explícito, tiene
 *         "(Contracted Linear Feet)" y un precio por pie lineal extra.
 *   26ft: el contrato NO menciona tamaño de camión (dice "a truck" / "un camión" genérico).
 *         Se reconoce porque el Origin Amount de la Sección 2 es exactamente 1,344 cubic ft.
 *         No tiene precio por pie lineal extra (siempre $0 en EE.UU.; el cargo real viene
 *         del camión en México).
 *   53ft: menciona "a 53-foot trailer" / "un remolque de 53 pies" explícito, pero el Origin
 *         Amount se da directo en pies cúbicos (sin pies lineales), y tampoco tiene precio
 *         por pie lineal extra.
 *
 * IMPORTANTE: esto depende de que el contrato mantenga el texto/formato boilerplate actual.
 * Si el contrato cambia de redacción, el parser puede fallar en extraer algún dato — por eso
 * generateAdditionalCosts.js siempre debe mostrar un resumen de lo que se extrajo ANTES de
 * generar el documento final, para que alguien lo confirme visualmente.
 */

const { PDFParse } = require('pdf-parse');
const fs = require('fs');
const { TRAILER_US } = require('../config/additionalCostsConfig');

async function extraerTextoPDF(filePath) {
  const buffer = fs.readFileSync(filePath);
  const parser = new PDFParse({ data: buffer });
  const result = await parser.getText();
  return result.text;
}

async function parsearContrato(filePath) {
  const texto = await extraerTextoPDF(filePath);
  const normalizado = texto.replace(/\r\n/g, '\n');
  const sinSaltos = normalizado.replace(/\s+/g, ' ');

  const idioma = /Acuerdo de Servicios de Mudanza/i.test(sinSaltos.slice(0, 600)) ? 'es' : 'en';

  const clientName = idioma === 'es'
    ? extraer(sinSaltos, /entre usted,\s*([A-Za-zÀ-ÿ.'\- ]+?)\s*\(["“]Cliente["”]/, 'nombre del cliente')
    : extraer(sinSaltos, /between you, ([A-Za-zÀ-ÿ.'\- ]+?)\s*\(["“]Client["”]/, 'nombre del cliente');

  const origenTexto = idioma === 'es'
    ? extraer(normalizado, /Origen \(Lugar a donde llegará el tráiler para cargar sus artículos del hogar\):\s*\n([\s\S]+?)\nDestino/, 'dirección de origen').trim()
    : extraer(normalizado, /Origin \(where we deliver the truck or trailer to pick up your household goods\):\s*\n([\s\S]+?)\nDestination/, 'dirección de origen').trim();

  const destinoTexto = idioma === 'es'
    ? extraer(normalizado, /Destino \(donde entregaremos sus artículos del hogar\):\s*\n([\s\S]+?)\n2\./, 'dirección de destino').trim()
    : extraer(normalizado, /Destination \(where we deliver your household goods\):\s*\n([\s\S]+?)\n2\./, 'dirección de destino').trim();

  const origenAmount = idioma === 'es'
    ? Number(extraer(sinSaltos, /hasta ([\d,]+) pies(?: cúbicos)?[\s\S]{0,60}?\(Monto Original\)/, 'Origin Amount').replace(/,/g, ''))
    : Number(extraer(sinSaltos, /up to ([\d,]+) (?:cubic )?feet[\s\S]{0,40}\(Origin Amount\)/, 'Origin Amount').replace(/,/g, ''));

  const destinationAmount = idioma === 'es'
    ? Number(extraer(sinSaltos, /hasta ([\d,]+) pies cúbicos[\s\S]{0,80}?\(Monto Destino\)/, 'Destination Amount').replace(/,/g, ''))
    : Number(extraer(sinSaltos, /up to ([\d,]+) cubic feet[\s\S]{0,60}\(Destination Amount\)/, 'Destination Amount').replace(/,/g, ''));

  const paisOrigen = detectarPais(origenTexto);
  const paisDestino = detectarPais(destinoTexto);
  const esNorthbound = paisOrigen === 'MX' && paisDestino !== 'MX';

  let trailerSizeUS;
  let contractedLinearFeet;
  let precioPorPieAdicional = 0; // 26ft y 53ft: siempre $0, confirmado (no hay ese cargo)
  let preciosMexico = [];
  let origenTrailerEs53 = false;

  if (esNorthbound) {
    // Northbound (México → EE.UU./Canadá): el trailer con desglose por pie lineal está en
    // Destino (Sección 4.E), no en Origen. El camión de Origen (México) no tiene tabla de
    // brackets — confirmado por Laura: se incluye el camión completo sin cargo adicional,
    // sin importar cuánto traigan (hasta su capacidad física).
    const trailerExplicitoDestino = idioma === 'es'
      ? sinSaltos.match(/(?:remolque|tráiler) de (\d+) pies/)
      : sinSaltos.match(/deliver to Destination using an? (\d+)-foot trailer/i);
    if (!trailerExplicitoDestino) {
      throw new Error(
        `No se pudo determinar el tamaño del tráiler de destino en un contrato Northbound. ` +
        `Por ahora solo está soportado el caso de 28ft en destino — avísame si es otro tamaño.`
      );
    }
    trailerSizeUS = Number(trailerExplicitoDestino[1]);
    if (trailerSizeUS !== 28) {
      throw new Error(
        `Northbound con tráiler de ${trailerSizeUS}ft en destino todavía no está configurado ` +
        `(solo 28ft por ahora, confirmado con Laura). Avísame para agregar este caso.`
      );
    }

    contractedLinearFeet = Number(
      idioma === 'es'
        ? extraer(sinSaltos, /hasta (\d+) pies lineales del tráiler sin cargo adicional/, 'pies lineales incluidos (Northbound 4.E)')
        : extraer(sinSaltos, /You may use up to (\d+) linear feet of the trailer at no additional charge/, 'pies lineales incluidos (Northbound 4.E)')
    );
    precioPorPieAdicional = Number(
      (idioma === 'es'
        ? extraer(sinSaltos, /[Cc]ada pie lineal adicional se cobra a \$([\d,.]+)/, 'precio por pie lineal adicional')
        : extraer(sinSaltos, /Each additional\s+linear foot is charged at \$([\d,.]+)/, 'precio por pie lineal adicional')
      ).replace(/,/g, '')
    );

    // Imagen del camión de Origen (México): large.png por default, trailer53.png solo si el
    // contrato menciona explícitamente un tráiler de 53 pies en la Sección 4.B (Origen) —
    // confirmado con Laura. En este contrato (Mattman) 4.B no menciona tamaño ("a truck"
    // genérico), así que usa large.png.
    const bloqueOrigenTexto = idioma === 'es'
      ? (normalizado.match(/4\.\s?B\.[\s\S]+?4\.\s?C\./) || [''])[0]
      : (normalizado.match(/4\.\s?B\.[\s\S]+?4\.\s?C\./) || [''])[0];
    origenTrailerEs53 = /53[\s-]foot|53 pies/i.test(bloqueOrigenTexto);
  } else {
    const trailerExplicito = idioma === 'es'
      ? sinSaltos.match(/(?:remolque|tráiler) de (\d+) pies/)
      : sinSaltos.match(/we will have delivered to Origin a (\d+)-foot trailer/i);

    if (trailerExplicito) {
      trailerSizeUS = Number(trailerExplicito[1]);
    } else if (origenAmount === 1344) {
      trailerSizeUS = 26; // confirmado: así se reconoce el 26ft, que no menciona tamaño
    } else {
      throw new Error(
        `No se pudo determinar el tamaño de trailer/camión — el contrato no menciona un tamaño ` +
        `explícito y el Origin Amount (${origenAmount} cf) no coincide con el patrón conocido de 26ft ` +
        `(1,344 cf). Puede ser un tamaño que todavía no está configurado — avísame.`
      );
    }

    if (trailerSizeUS === 28) {
      contractedLinearFeet = Number(
        idioma === 'es'
          ? extraer(sinSaltos, /hasta (\d+) pies lineales[\s\S]{0,30}?\(Pies Lineales Contratados\)/, 'Pies Lineales Contratados')
          : extraer(sinSaltos, /up to (\d+) linear feet\s*\(Contracted Linear Feet\)/, 'Contracted Linear Feet')
      );
      precioPorPieAdicional = Number(
        (idioma === 'es'
          ? extraer(sinSaltos, /[Cc]ada pie lineal adicional se cobra a \$([\d,.]+)/, 'precio por pie lineal adicional')
          : extraer(sinSaltos, /Each additional\s+linear foot is charged at \$([\d,.]+)/, 'precio por pie lineal adicional')
        ).replace(/,/g, '')
      );
    } else if (trailerSizeUS === 26) {
      contractedLinearFeet = Number(
        idioma === 'es'
          ? extraer(sinSaltos, /no pueden ocupar más de (\d+) pies lineales/, 'pies lineales (26ft)')
          : extraer(sinSaltos, /(?:no more than|occupy no more than) (\d+) linear feet/, 'linear feet (26ft)')
      );
    } else if (trailerSizeUS === 53) {
      const trailer53 = TRAILER_US[53];
      const punto = trailer53.puntosDeQuiebre.find((p) => p.cubicFt === origenAmount);
      if (!punto) {
        throw new Error(
          `El Origin Amount del contrato (${origenAmount} cf) no coincide exactamente con ningún ` +
          `punto de quiebre conocido de 53ft (600/1,000/1,450/2,400/3,400). Revisa el contrato a mano.`
        );
      }
      contractedLinearFeet = punto.pieLineal;
    } else {
      throw new Error(`Trailer de ${trailerSizeUS} ft todavía no está configurado. Avísame para agregarlo.`);
    }

    const bloqueMexico = idioma === 'es'
      ? extraer(
          normalizado,
          /excedentes\s+necesarios para el transporte al Destino:\s*\n([\s\S]+?)\n\s*4\.\s?F\./,
          'tabla de precios adicionales en México'
        )
      : extraer(
          normalizado,
          /needed for transportation to Destination:\s*\n([\s\S]+?)\n\s*4\.F/,
          'tabla de precios adicionales en México'
        );
    preciosMexico = parsearBloquePreciosMexico(bloqueMexico, idioma);
  }

  let precioTotalContrato = null;
  try {
    precioTotalContrato = Number(
      (idioma === 'es'
        ? extraer(sinSaltos, /es el siguiente:\s*\$([\d,.]+)/, 'precio total del contrato')
        : extraer(sinSaltos, /Total Price for All Services[\s\S]{0,120}?\$([\d,.]+)/, 'precio total del contrato')
      ).replace(/,/g, '')
    );
  } catch (err) {
    // se deja en null; solo hace falta para 26ft
  }

  return {
    clientName: clientName.trim(),
    idioma,
    origenTexto,
    destinoTexto,
    paisOrigen,
    paisDestino,
    esNorthbound,
    origenAmount,
    destinationAmount,
    contractedLinearFeet,
    precioPorPieAdicional,
    precioTotalContrato,
    trailerSizeUS,
    preciosMexico,
    origenTrailerEs53,
  };
}

function extraer(texto, regex, nombreCampo) {
  const match = texto.match(regex);
  if (!match) {
    throw new Error(
      `No se pudo encontrar "${nombreCampo}" en el contrato. El texto del contrato puede haber ` +
      `cambiado de redacción — revisa el PDF a mano para este campo.`
    );
  }
  return match[1];
}

function parsearBloquePreciosMexico(bloque, idioma) {
  const lineas = bloque.split('\n').map((l) => l.trim()).filter(Boolean);
  const resultado = [];
  for (const linea of lineas) {
    const soloHasta = idioma === 'es'
      ? linea.match(/^Hasta ([\d,]+) pies cúbicos:\s*(Sin cargo adicional|\$[\d,.]+)/i)
      : linea.match(/^Up to ([\d,]+) cubic feet:\s*(No charge|\$[\d,.]+)/i);
    const rango = idioma === 'es'
      ? linea.match(/^([\d,]+) a ([\d,]+) pies cúbicos:\s*(Sin cargo adicional|\$[\d,.]+)/i)
      : linea.match(/^([\d,]+) to ([\d,]+) cubic feet:\s*(No charge|\$[\d,.]+)/i);
    if (soloHasta) {
      resultado.push({ hasta: Number(soloHasta[1].replace(/,/g, '')), precio: parsearPrecio(soloHasta[2]) });
    } else if (rango) {
      resultado.push({
        desde: Number(rango[1].replace(/,/g, '')),
        hasta: Number(rango[2].replace(/,/g, '')),
        precio: parsearPrecio(rango[3]),
      });
    }
  }
  if (resultado.length === 0) {
    throw new Error('No se pudo interpretar ninguna línea de la tabla de precios adicionales en México.');
  }
  return resultado;
}

function parsearPrecio(texto) {
  if (/no charge|sin cargo adicional/i.test(texto)) return 0;
  return Number(texto.replace(/[$,]/g, ''));
}

// Abreviaciones de estado mexicano (2-4 letras) que aparecen antes del código postal en
// direcciones que NO mencionan "Mexico" explícitamente (ej. "San Miguel de Allende, GTO
// 37883") — confirmado con el contrato de Mattman (Northbound), donde el origen en México
// no incluye la palabra "Mexico" en ningún lado de la dirección.
const ESTADOS_MX = [
  'AGU', 'BCN', 'BCS', 'CAM', 'CHP', 'CHH', 'CDMX', 'COA', 'COL', 'DUR', 'GUA', 'GTO', 'GRO',
  'HID', 'JAL', 'MEX', 'MIC', 'MOR', 'NAY', 'NLE', 'OAX', 'PUE', 'QUE', 'QRO', 'ROO', 'SLP',
  'SIN', 'SON', 'TAB', 'TAM', 'TLA', 'VER', 'YUC', 'ZAC',
];
const ESTADOS_MX_REGEX = new RegExp(`\\b(${ESTADOS_MX.join('|')})\\b`);

function detectarPais(direccionTexto) {
  const t = direccionTexto.toLowerCase();
  if (t.includes('mexico') || t.includes('méxico')) return 'MX';
  if (t.includes('canada') || t.includes('canadá')) return 'CA';
  if (ESTADOS_MX_REGEX.test(direccionTexto)) return 'MX';
  return 'US';
}

module.exports = { parsearContrato, extraerTextoPDF };
