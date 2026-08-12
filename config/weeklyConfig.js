/**
 * weeklyConfig.js
 * Reglas de negocio confirmadas para las Weeklies (reporte semanal a transportistas,
 * agentes aduanales, y crews de descarga).
 */

const COLORES = {
  AMARILLO_PENDIENTE: 'FFFFFF00',
  AMARILLO_PENDIENTE_CLARO: 'FFFFFFCC', // fecha pasada + pendiente (caso raro, por si acaso)
  NARANJA_MISMO_PROVEEDOR: 'FFFFC000',
  NARANJA_CLARO: 'FFFFE0B3',
  VERDE_FECHA_PRINCIPAL: 'FF92D050',
  VERDE_CLARO: 'FFD6EDB8',
  AZUL_LLEGADA_MEXICO: 'FF00B0F0',
  AZUL_CLARO: 'FFBEE7FB',
  TEXTO_ROJO_NOTAS: 'FFFF0000',
};

/**
 * Columnas que cuentan como "operativas" para el resaltado amarillo de pendiente.
 * CLIENT, Client's Phone, TAGS y PMA son columnas de identidad y NUNCA se resaltan como
 * pendientes, aunque a veces estén vacías.
 */
const COLUMNAS_IDENTIDAD = ["CLIENT", "Client's Phone", 'CLIENTE', 'Teléfono del cliente', 'TAGS', 'PMA'];

/**
 * Textos que indican información pendiente/incompleta (case-insensitive).
 */
const TEXTOS_PENDIENTE_EXACTOS = ['-']; // coincidencia exacta (para no marcar números de teléfono con guiones)
const TEXTOS_PENDIENTE_PARCIALES = ['tbd', '(en blanco)', 'not yet defined', 'no yet defined'];

function esTextoPendiente(valor) {
  if (valor === null || valor === undefined) return true;
  const texto = String(valor).trim().toLowerCase();
  if (texto === '') return true;
  if (TEXTOS_PENDIENTE_EXACTOS.includes(texto)) return true;
  return TEXTOS_PENDIENTE_PARCIALES.some((t) => texto.includes(t));
}

/**
 * WEEKLY_TIPOS: definición de columnas por tipo de weekly.
 * col = nombre exacto de la columna en "PASTE HERE".
 * label = encabezado que se muestra en la weekly (puede diferir del nombre original).
 * operativa: true = SOLO estas columnas se resaltan en amarillo si están pendientes/vacías.
 * Cuáles son "operativas" varía por tipo de weekly (ej. Time SÍ importa para transportistas,
 * pero NO para agentes aduanales) — por eso es una bandera por columna, no una lista global.
 */
const WEEKLY_TIPOS = {
  transportista: {
    filtroColumna: 'NOB Transportation Company',
    formatoSalida: 'excel',
    columnas: [
      { col: 'CLIENT', label: 'CLIENT', identidad: true },
      { col: "Client's Phone", label: "Client's Phone", identidad: true },
      { col: 'TAGS', label: 'TAGS', identidad: true },
      { col: 'PMA', label: 'PMA', identidad: true },
      { col: 'X NOB Day for Truck to Arrive', label: 'NOB Day for Truck to Arrive', esFechaPrincipal: true },
      { col: 'NOB Time', label: 'NOB Time', operativa: true },
      { col: 'NOB Truck Type', label: 'NOB Truck Type', operativa: true },
      { col: 'Weight', label: 'Weight' },
      { col: 'NOB Address', label: 'NOB Address', operativa: true },
      { col: 'NOB Notes for Transportation Company', label: 'NOB Notes for Transportation Company', esNota: true },
      { col: 'Size of the Shipment', label: 'Size of the Shipment', operativa: true },
      { col: 'Customs Broker', label: 'Customs Broker', operativa: true },
      { col: 'X CUSTOMS Day for Truck to Arrive', label: 'CUSTOMS Day for Truck to Arrive', esFechaLlegadaMX: true },
      { col: 'CUSTOMS Time', label: 'CUSTOMS Time', operativa: true },
    ],
  },

  agenteAduanal: {
    filtroColumna: 'Customs Broker',
    formatoSalida: 'excel',
    columnas: [
      { col: 'CLIENT', label: 'CLIENT', identidad: true },
      { col: 'TAGS', label: 'TAGS', identidad: true },
      { col: 'PMA', label: 'PMA', identidad: true },
      { col: 'X CUSTOMS Day for Truck to Arrive', label: 'Arrival day', esFechaPrincipal: true },
      { col: 'CUSTOMS Time', label: 'CUSTOMS Time' },
      { col: 'NOB Transportation Company', label: 'NOB Transportation Company' },
      { col: 'X NOB Day for Truck to Arrive', label: 'NOB Day for Truck to Arrive' },
      { col: 'NOB Truck Type', label: 'NOB Truck Type', operativa: true },
      { col: 'Trailer Number', label: 'Trailer Number' },
      { col: 'Size of the Shipment', label: 'Size of the Shipment', operativa: true },
      { col: 'SOB Truck Type', label: 'Truck Type in MX', operativa: true },
      { col: 'Crates', label: 'Crates', operativa: true },
      { col: 'Visa / Menaje / Immigration Status', label: 'Visa / Menaje / Immigration Status' },
      { col: 'CUSTOMS Notes', label: 'CUSTOMS Notes', esNota: true },
      { col: 'X SOB Day for Truck to Arrive', label: 'SOB Day for Truck to Arrive', esFechaLlegadaMX: true },
      { col: 'X SOB Earliest Day to Receive Shipment', label: 'SOB Earliest Day to Receive Shipment' },
      { col: 'X SOB Latest Day to Receive Shipment', label: 'SOB Latest Day to Receive Shipment' },
      { col: 'SOB Address', label: 'Address', operativa: true },
      { col: 'SOB Contact for Driver', label: 'Contact for Driver', operativa: true },
      { col: 'SOB Meeting Point', label: 'Meeting Point', operativa: true },
      { col: 'SOB (Un)Loading Plan', label: '(Un)Loading Plan', operativa: true },
      { col: 'SOB Shuttle', label: 'Shuttle', operativa: true },
    ],
  },

  crew: {
    filtroColumna: 'SOB Crew',
    formatoSalida: 'imagen',
    columnas: [
      { col: 'CLIENT', label: 'CLIENTE', identidad: true },
      { col: "Client's Phone", label: 'Teléfono del cliente', identidad: true },
      { col: 'TAGS', label: 'TAGS', identidad: true },
      { col: 'PMA', label: 'PMA', identidad: true },
      { col: 'X SOB Day for Truck to Arrive', label: 'Día que llega el camión', esFechaPrincipal: true, esFechaLlegadaMX: true },
      { col: 'SOB Truck Type', label: 'Tipo de camión', operativa: true },
      { col: 'Size of the Shipment', label: 'Tamaño de carga', operativa: true },
      { col: 'Crates', label: 'Crates', operativa: true },
      { col: 'SOB Address', label: 'Dirección', operativa: true },
      { col: 'SOB Meeting Point', label: 'Punto de encuentro', operativa: true },
      { col: 'SOB (Un)Loading Plan', label: 'Plan para cargar / descargar', operativa: true },
      { col: 'SOB Shuttle', label: 'Acarreo', operativa: true },
      { col: 'SOB Notes for Crew', label: 'Notas', esNota: true },
      { col: 'SOB Contact for Driver', label: 'Contacto para el chofer', operativa: true },
      { col: 'SOB Crew', label: 'Crew', identidad: true, operativa: true },
    ],
  },
};

module.exports = { COLORES, COLUMNAS_IDENTIDAD, esTextoPendiente, WEEKLY_TIPOS };
