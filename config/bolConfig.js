/**
 * bolConfig.js
 * Reglas de negocio para la generación del Bill of Lading (BOL).
 * Fuente: bol_variable_map.json (mapeo confirmado con el equipo BMM).
 */

const CUSTOMS_BROKERS = {
  LEGO: {
    nombre: 'LEGO GROUP LLC',
    direccion: ['1802 MARKLEY LN', 'LAREDO, TX 78041'],
    contacto: 'ELSA GONZALEZ',
    telefono: '(956) 726-9941',
    email: 'elsa@legogroup.net',
  },
  TIS_WORLDWIDE: {
    nombre: 'TIS Worldwide',
    direccion: ['9601 Carnegie Ave', 'El Paso, TX 79925'],
    contacto: 'RIGO DURAN',
    telefono: '+1 (915) 892-4958',
    // Southbound o Bonded
  },
  TIS_CDJ: {
    nombre: 'TIS CDJ',
    direccion: ['Ave. Antonio J. Bermudez 310-4, Partido Romero', 'Cd. Juarez, Chih. 32320'],
    contacto: 'RIGO DURAN',
    telefono: '+1 (915) 892-4958',
    // Northbound
  },
};

// A qué bloque de TIS corresponde según el escenario de dirección
function resolveTISBlock(direction) {
  return direction === 'northbound' ? CUSTOMS_BROKERS.TIS_CDJ : CUSTOMS_BROKERS.TIS_WORLDWIDE;
}

const SHIPPER_FIJO_BMM = {
  direccion: ['7740 W Roosevelt Road,', 'Oak Park, IL 60304'],
};

const CONTACT_POOL = {
  MOISES: { nombre: 'Moises', telefono: '+1 (520) 486-1110' },
  Hannia: { nombre: 'Hannia', telefono: '+1 (956) 301-0924' },
  Laura: { nombre: 'Laura', telefono: '+1 (929) 332-6992' },
  Ana: { nombre: 'Ana', telefono: '+1 (805) 656-7826' },
  Ariel: { nombre: 'Ariel', telefono: '+1 (915) 502-8546' },
};

const CONTACT_NAMES_VALID = ['Hannia', 'Laura', 'Ana', 'Ariel']; // pool de posibles PMA

/**
 * Arma el bloque de contacto dinámico:
 * Moises siempre primero, luego el PMA del embarque, luego Hannia
 * (si el PMA ya es Hannia, no se repite). Laura, Ana y Ariel ya no se listan
 * salvo que alguno de ellos sea el PMA asignado a ese embarque.
 * Formato de salida: "Nombre - Teléfono, Nombre - Teléfono, ..."
 */
function buildContactBlock(pmaFirstName) {
  const pma = CONTACT_POOL[pmaFirstName];
  if (!pma) {
    throw new Error(`PMA desconocido: "${pmaFirstName}". Debe ser uno de: ${CONTACT_NAMES_VALID.join(', ')}`);
  }

  const ordered = [CONTACT_POOL.MOISES, pma];
  if (pmaFirstName !== 'Hannia') {
    ordered.push(CONTACT_POOL.Hannia);
  }

  // Una persona por línea (separado por \n) en vez de todo en una sola línea larga:
  // más legible y más predecible para el wrap de la celda que un párrafo corrido.
  const texto = ordered.map((c) => `${c.nombre} - ${c.telefono}`).join('\n');

  return { texto };
}

/**
 * Determina el escenario de dirección a partir de la columna TAGS de "PASTE HERE".
 * TAGS contiene 'Northbound' -> northbound. TAGS contiene 'Bonded' -> bonded_canada.
 * 'LIVE' o vacío -> southbound (caso más común, es el default).
 */
function resolveDirectionFromTags(tags) {
  const t = (tags || '').toLowerCase();
  if (t.includes('northbound')) return 'northbound';
  if (t.includes('bonded')) return 'bonded_canada';
  return 'southbound';
}

/**
 * Determina el escenario de dirección según país de origen/destino.
 * origenPais / destinoPais: 'US' | 'MX' | 'CA'
 * (Se conserva por si se necesita en otro contexto; la fuente de datos real usa TAGS.)
 */
function resolveDirection(origenPais, destinoPais) {
  if (origenPais === 'CA' || destinoPais === 'CA') return 'bonded_canada';
  if (origenPais === 'US' && destinoPais === 'MX') return 'southbound';
  if (origenPais === 'MX' && destinoPais === 'US') return 'northbound';
  throw new Error(`No se pudo determinar el escenario de dirección para origen=${origenPais}, destino=${destinoPais}`);
}

/**
 * Calcula el peso según la regla de negocio confirmada.
 * cubicFt: number, carrier: string
 * Devuelve { valor, esManual }
 */
function calcularPeso(cubicFt, carrier) {
  if (carrier === 'Ground Freight Solutions') {
    return { valor: null, esManual: true }; // el usuario lo captura a mano
  }
  return { valor: Math.round(cubicFt * 6.5), esManual: false };
}

/**
 * Order # / PO # con el formato confirmado.
 * clientLastName: string, fecha: Date
 */
function buildOrderNumbers(clientLastName, fecha) {
  const iniciales = clientLastName.trim().slice(0, 3).toUpperCase();
  const mm = String(fecha.getMonth() + 1).padStart(2, '0');
  const dd = String(fecha.getDate()).padStart(2, '0');
  const yyyy = fecha.getFullYear();
  const fechaStr = `${mm}${dd}${yyyy}`;
  return {
    orderNumber: `BMM-${iniciales}${fechaStr}-01`,
    poNumber: `${iniciales}${fechaStr}-01`,
  };
}

module.exports = {
  CUSTOMS_BROKERS,
  resolveTISBlock,
  SHIPPER_FIJO_BMM,
  CONTACT_POOL,
  buildContactBlock,
  resolveDirection,
  resolveDirectionFromTags,
  calcularPeso,
  buildOrderNumbers,
};
