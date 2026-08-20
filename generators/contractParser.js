/**
 * contractParser.js
 * Extrae los datos clave del contrato (Moving Services Agreement) que hacen falta para
 * generar "Understanding Additional Costs for Larger Shipments".
 *
 * IMPORTANTE: esto depende de que el contrato mantenga el mismo texto/formato boilerplate
 * que la plantilla actual. Si el texto del contrato cambia de redacción, el parser puede
 * fallar en extraer algún dato — por eso generateAdditionalCosts.js siempre debe mostrar
 * un resumen de lo que se extrajo ANTES de generar el documento final, para que alguien
 * lo confirme visualmente en vez de confiar ciegamente en el parseo.
 */

const { PDFParse } = require('pdf-parse');
const fs = require('fs');

async function extraerTextoPDF(filePath) {
  const buffer = fs.readFileSync(filePath);
  const parser = new PDFParse({ data: buffer });
  const result = await parser.getText();
  return result.text;
}

/**
 * Parsea el contrato y devuelve todos los datos necesarios para el documento de costos
 * adicionales. Lanza error con un mensaje claro si no puede encontrar un campo obligatorio.
 */
async function parsearContrato(filePath) {
  const texto = await extraerTextoPDF(filePath);
  const normalizado = texto.replace(/\r\n/g, '\n');
  // Versión con saltos de línea colapsados a espacios simples, para los campos de una sola
  // frase — el PDF a veces corta la línea a media frase ("Each additional\nlinear foot..."),
  // y sin esto cualquier salto de línea inesperado rompe el regex.
  const sinSaltos = normalizado.replace(/\s+/g, ' ');

  const clientName = extraer(sinSaltos, /between you, ([A-Za-zÀ-ÿ.'\- ]+?)\s*\(“Client”/, 'nombre del cliente');

  // Direcciones: todo lo que está entre "Origin (...)" y "Destination (...)" es la
  // dirección de origen; entre "Destination (...)" y el siguiente numeral "2." es destino.
  const origenTexto = extraer(
    normalizado,
    /Origin \(where we deliver the truck or trailer to pick up your household goods\):\s*\n([\s\S]+?)\nDestination/,
    'dirección de origen'
  ).trim();
  const destinoTexto = extraer(
    normalizado,
    /Destination \(where we deliver your household goods\):\s*\n([\s\S]+?)\n2\./,
    'dirección de destino'
  ).trim();

  const origenAmount = Number(extraer(sinSaltos, /up to ([\d,]+) cubic feet[\s\S]{0,40}\(Origin Amount\)/, 'Origin Amount').replace(/,/g, ''));
  const destinationAmount = Number(extraer(sinSaltos, /up to ([\d,]+) cubic feet[\s\S]{0,60}\(Destination Amount\)/, 'Destination Amount').replace(/,/g, ''));

  const contractedLinearFeet = Number(extraer(sinSaltos, /up to (\d+) linear feet\s*\(Contracted Linear Feet\)/, 'Contracted Linear Feet'));
  const precioPorPieAdicional = Number(extraer(sinSaltos, /Each additional\s+linear foot is charged at \$([\d,.]+)/, 'precio por pie lineal adicional').replace(/,/g, ''));

  const trailerSizeTexto = extraer(sinSaltos, /We will have delivered to Origin a (\d+)-foot trailer/, 'tamaño del trailer en EE.UU.');
  const trailerSizeUS = Number(trailerSizeTexto);

  // Bloque de precios "Additional Costs for Larger Shipments in Mexico" — lista variable de
  // líneas "X to Y cubic feet: $Z" o "Up to X cubic feet: No charge".
  const bloqueMexico = extraer(
    normalizado,
    /following total additional charges apply for the additional cubic feet\s*\n\s*needed for transportation to Destination:\s*\n([\s\S]+?)\n\s*4\.F/,
    'tabla de precios adicionales en México'
  );
  const preciosMexico = parsearBloquePreciosMexico(bloqueMexico);

  const paisOrigen = detectarPais(origenTexto);
  const paisDestino = detectarPais(destinoTexto);

  return {
    clientName: clientName.trim(),
    origenTexto,
    destinoTexto,
    paisOrigen,
    paisDestino,
    origenAmount,
    destinationAmount,
    contractedLinearFeet,
    precioPorPieAdicional,
    trailerSizeUS,
    preciosMexico, // [{ hastaCubicFt, precio }] o [{ desde, hasta, precio }]
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

/**
 * "Up to 1,450 cubic feet: No charge\n1,451 to 2,400 cubic feet: $1,300"
 * -> [{ hasta: 1450, precio: 0 }, { desde: 1451, hasta: 2400, precio: 1300 }]
 */
function parsearBloquePreciosMexico(bloque) {
  const lineas = bloque.split('\n').map((l) => l.trim()).filter(Boolean);
  const resultado = [];
  for (const linea of lineas) {
    const soloHasta = linea.match(/^Up to ([\d,]+) cubic feet:\s*(No charge|\$[\d,.]+)/i);
    const rango = linea.match(/^([\d,]+) to ([\d,]+) cubic feet:\s*(No charge|\$[\d,.]+)/i);
    if (soloHasta) {
      resultado.push({
        hasta: Number(soloHasta[1].replace(/,/g, '')),
        precio: parsearPrecio(soloHasta[2]),
      });
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
  if (/no charge/i.test(texto)) return 0;
  return Number(texto.replace(/[$,]/g, ''));
}

/**
 * Heurística simple: si el texto de la dirección menciona México (o un estado mexicano
 * común) es 'MX'; si menciona Canadá es 'CA'; si no, se asume 'US'.
 */
function detectarPais(direccionTexto) {
  const t = direccionTexto.toLowerCase();
  if (t.includes('mexico') || t.includes('méxico')) return 'MX';
  if (t.includes('canada') || t.includes('canadá')) return 'CA';
  return 'US';
}

module.exports = { parsearContrato, extraerTextoPDF };
