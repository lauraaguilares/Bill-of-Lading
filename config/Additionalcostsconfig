/**
 * additionalCostsConfig.js
 * Reglas confirmadas para "Understanding Additional Costs for Larger Shipments".
 *
 * Los 4 "brackets" de tipo de camión en México son fijos y universales (confirmado en las
 * plantillas: la única diferencia entre archivos de un mismo trailer size era CUÁL precio
 * se había capturado, no la estructura). El precio de cada bracket sale del contrato
 * (preciosMexico, ver contractParser.js) — si el contrato no menciona un bracket, es
 * porque ya está incluido en lo contratado (se asume $0).
 */

const BRACKETS_MEXICO = [
  { hasta: 600, label: 'Small Straight Truck\nUp to 600 ft3' },
  { desde: 601, hasta: 1000, label: 'Medium Straight Truck\n601 to 1,000 cubic ft.' },
  { desde: 1001, hasta: 1450, label: 'Large Straight Truck\n1,001 to 1,450 cubic ft.' },
  { desde: 1451, hasta: 2400, label: '53-Foot Trailer, exclusive use\n1,451 to 2,400 cubic ft' },
  { desde: 2401, hasta: 3400, label: '53-Foot Trailer, exclusive use\n2,401 to 3,400 cubic ft' },
];

/**
 * Pies cúbicos por pie lineal, capacidad máxima (pies lineales), y las etiquetas exactas
 * que usa cada plantilla (varían — 28ft dice "Trailer", 26ft dice "Box Truck", por ejemplo)
 * — confirmado revisando cada plantilla original directamente, no asumido.
 * (53ft queda pendiente de construir en una siguiente fase: su plantilla tiene una
 * estructura distinta, de una sola fila resumen en vez de desglose por pie lineal.)
 */
const TRAILER_US = {
  28: {
    cubicFtPorPieLineal: 72,
    maxPiesLineales: 28,
    labelGrupo: '28-Foot Trailer in the US',
    labelPrecio: 'Additional Price for Trailer in the US',
  },
  26: {
    cubicFtPorPieLineal: 64,
    maxPiesLineales: 26,
    labelGrupo: '26-Foot Box Truck in the US',
    labelPrecio: 'Additional Price for Truck in the US',
  },
  // 53ft usa una estructura distinta: en vez de desglosar CADA pie lineal, la plantilla
  // muestra solo los puntos donde cambia de bracket en México ("Up to 8 / Up to 600", "Up
  // to 14 / Up to 1,000", etc.) — confirmado directo en la plantilla original, los pies
  // cúbicos de cada punto NO se derivan de un multiplicador constante por pie lineal.
  53: {
    modoResumen: true,
    maxPiesLineales: 47,
    labelGrupo: '53-Foot Trailer in the US',
    labelPrecio: 'Additional Price for Trailer in the US',
    puntosDeQuiebre: [
      { pieLineal: 8, cubicFt: 600 },
      { pieLineal: 14, cubicFt: 1000 },
      { pieLineal: 20, cubicFt: 1450 },
      { pieLineal: 32, cubicFt: 2400 },
      { pieLineal: 47, cubicFt: 3400 },
    ],
  },
};

/**
 * Encuentra a qué bracket de México pertenecen X cubic feet.
 */
function bracketParaCubicFt(cubicFt) {
  return BRACKETS_MEXICO.find((b) => cubicFt <= b.hasta) || BRACKETS_MEXICO[BRACKETS_MEXICO.length - 1];
}

/**
 * Encuentra el precio (del contrato) que corresponde a un bracket específico. Si el
 * contrato no lo menciona explícitamente, se asume $0 (ya incluido en lo contratado).
 */
function precioParaBracket(bracket, preciosMexico) {
  const encontrado = preciosMexico.find((p) => {
    if (p.desde != null) return p.desde === bracket.desde && p.hasta === bracket.hasta;
    return p.hasta === bracket.hasta && bracket.desde == null;
  });
  return encontrado ? encontrado.precio : 0;
}

module.exports = { BRACKETS_MEXICO, TRAILER_US, bracketParaCubicFt, precioParaBracket };
