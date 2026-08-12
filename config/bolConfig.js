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
  Hannia: { nombre: 'Hannia', telefono: '+1 (956) 301-0924', correo: 'Hannia.Alcala@BestMexicoMovers.com' },
  Laura: { nombre: 'Laura', telefono: '+1 (929) 332-6992', correo: 'Laura.Aguilar@BestMexicoMovers.com' },
  Ana: { nombre: 'Ana', telefono: '+1 (805) 656-7826', correo: 'Ana.Bermudez@BestMexicoMovers.com' },
  Ariel: { nombre: 'Ariel', telefono: '+1 (915) 502-8546', correo: 'Ariel.Hererra@BestMexicoMovers.com' },
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
 * Tabla confirmada: cubic ft y peso dependen del TIPO DE CAMIÓN (NOB Truck Type), no de
 * "Size of the Shipment" ni de una fórmula genérica. Único caso con excepción: 26 ft, donde
 * si el archivo trae un Weight distinto especificado, ese manda sobre el default.
 */
const TRUCK_SPECS = {
  '53': { cubicFt: 2400, peso: 15600, pesoSobreescribible: false },
  '26': { cubicFt: 1344, peso: 8735, pesoSobreescribible: true },
  '28': { cubicFt: 1450, peso: 9425, pesoSobreescribible: false },
};

/**
 * Determina cubic ft y peso a partir del tipo de camión (texto libre como "53 ft",
 * "26 ft", "28 ft (PUP)") y, si aplica, el peso especificado en el archivo original.
 * Devuelve { cubicFt, peso, reconocido }. Si el tipo de camión no matchea ninguna
 * especificación conocida, reconocido=false y hay que resolverlo manualmente.
 */
function calcularCubicFtYPeso(truckType, pesoEspecificadoEnArchivo) {
  const texto = String(truckType || '');
  const match = texto.match(/^\s*(\d{2})/); // "53 ft" -> "53", "28 ft (PUP)" -> "28"
  const spec = match ? TRUCK_SPECS[match[1]] : null;

  if (!spec) {
    return { cubicFt: null, peso: null, reconocido: false };
  }

  const usaOverride = spec.pesoSobreescribible && pesoEspecificadoEnArchivo != null;
  return {
    cubicFt: spec.cubicFt,
    peso: usaOverride ? pesoEspecificadoEnArchivo : spec.peso,
    reconocido: true,
  };
}

/**
 * Calcula el peso según la regla de negocio ANTERIOR (cubicFt x 6.5), usada solo como
 * respaldo cuando el tipo de camión no se reconoce (ver calcularCubicFtYPeso).
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
  calcularCubicFtYPeso,
  buildOrderNumbers,
};
