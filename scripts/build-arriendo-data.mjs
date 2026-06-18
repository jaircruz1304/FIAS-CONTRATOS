import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import * as XLSX from 'xlsx';

const OUT_DIR = 'data';
const OUT_JSON = path.join(OUT_DIR, 'control-arriendo.json');
const OUT_META = path.join(OUT_DIR, 'control-arriendo-meta.json');

const SHARE_ID = 'IQC10PyN9fkiRZyU6zso_j3OAQMnmyHwJldxvOby5_U4vZc';
const SHARE_LINK = 'https://fiasec-my.sharepoint.com/:x:/g/personal/jcruzg_fias_org_ec/IQC10PyN9fkiRZyU6zso_j3OAQMnmyHwJldxvOby5_U4vZc?e=Aq3X3K';

const DOWNLOAD_CANDIDATES = [
  {
    name: 'Share link exacto con download=1',
    url: `${SHARE_LINK}&download=1`
  },
  {
    name: 'Share link exacto con download=1 y web=0',
    url: `${SHARE_LINK}&download=1&web=0`
  },
  {
    name: 'SharePoint download.aspx con share id',
    url: `https://fiasec-my.sharepoint.com/personal/jcruzg_fias_org_ec/_layouts/15/download.aspx?share=${SHARE_ID}`
  }
];

const MONTHS = {
  enero: 1,
  febrero: 2,
  marzo: 3,
  abril: 4,
  mayo: 5,
  junio: 6,
  julio: 7,
  agosto: 8,
  septiembre: 9,
  setiembre: 9,
  octubre: 10,
  noviembre: 11,
  diciembre: 12
};

const EXPECTED_HEADERS = [
  'Mes',
  'Año',
  'Canon programado',
  'Canon ejecutado',
  'Saldo canon',
  'Estado canon',
  'Factura canon',
  'Fecha pago canon',
  'Solicitud pago canon',
  'Link expediente canon',
  'Servicios básicos',
  'Estado servicios',
  'Factura servicios',
  'Fecha pago servicios',
  'Solicitud pago servicios',
  'Link expediente servicios',
  'Expensas programadas',
  'Expensas ejecutadas',
  'Saldo expensas',
  'Estado expensas',
  'Notificación expensas',
  'Fecha pago expensas',
  'Solicitud pago expensas',
  'Link expediente expensas',
  'Total recurrente programado',
  'Total ejecutado',
  'Saldo recurrente',
  'Estado mensual',
  'Observación de control'
];

function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function normalizeKey(value) {
  return normalizeText(value)
    .replace(/[%()$]/g, '')
    .replace(/[^a-z0-9]+/g, '');
}

function parseNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;

  let s = String(value).trim();
  s = s.replace(/[^\d,.-]/g, '');

  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }

  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function excelDateToISO(value) {
  if (!value) return null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;

    const yyyy = String(parsed.y).padStart(4, '0');
    const mm = String(parsed.m).padStart(2, '0');
    const dd = String(parsed.d).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  const text = String(value).trim();
  if (!text) return null;

  const iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (iso) {
    const [, y, m, d] = iso;
    return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }

  const latam = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (latam) {
    const [, d, m, y] = latam;
    return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }

  return text;
}

const FIELD_ALIASES = {
  mes: ['mes'],
  anio: ['año', 'ano', 'anio'],
  canon_programado: ['canon programado'],
  canon_ejecutado: ['canon ejecutado'],
  saldo_canon: ['saldo canon'],
  estado_canon: ['estado canon'],
  factura_canon: ['factura canon', 'factura/obs. canon', 'factura obs canon'],
  fecha_pago_canon: ['fecha pago canon', 'fecha canon'],
  solicitud_pago_canon: ['solicitud pago canon', 'solicitud canon'],
  link_expediente_canon: ['link expediente canon', 'enlace canon', 'expediente canon'],

  servicios_basicos: ['servicios básicos', 'servicios basicos', 'servicios usd'],
  estado_servicios: ['estado servicios'],
  factura_servicios: ['factura servicios', 'factura/obs. servicios', 'factura obs servicios'],
  fecha_pago_servicios: ['fecha pago servicios', 'fecha servicios'],
  solicitud_pago_servicios: ['solicitud pago servicios', 'solicitud servicios'],
  link_expediente_servicios: ['link expediente servicios', 'enlace servicios', 'expediente servicios'],

  expensas_programadas: ['expensas programadas', 'expensa programada'],
  expensas_ejecutadas: ['expensas ejecutadas', 'expensa ejecutada'],
  saldo_expensas: ['saldo expensas', 'saldo expensa'],
  estado_expensas: ['estado expensas', 'estado expensa'],
  notificacion_expensas: ['notificación expensas', 'notificacion expensas', 'notificación expensa', 'notificacion expensa'],
  fecha_pago_expensas: ['fecha pago expensas', 'fecha expensa'],
  solicitud_pago_expensas: ['solicitud pago expensas', 'solicitud expensa'],
  link_expediente_expensas: ['link expediente expensas', 'enlace expensa', 'expediente expensa'],

  total_recurrente_programado: ['total recurrente programado'],
  total_ejecutado: ['total ejecutado'],
  saldo_recurrente: ['saldo recurrente'],
  estado_mensual: ['estado mensual', 'estado mes'],
  observacion: ['observación de control', 'observacion de control', 'observación', 'observacion']
};

const NORMALIZED_ALIASES = Object.fromEntries(
  Object.entries(FIELD_ALIASES).map(([field, aliases]) => [
    field,
    aliases.map(normalizeKey)
  ])
);

function valueByAliases(row, field) {
  const aliases = NORMALIZED_ALIASES[field] || [];
  for (const [key, value] of Object.entries(row)) {
    if (aliases.includes(normalizeKey(key))) return value;
  }
  return '';
}

function isProbablyHtml(buffer) {
  const sample = buffer.subarray(0, 900).toString('utf8').trim().toLowerCase();
  return (
    sample.startsWith('<!doctype html') ||
    sample.startsWith('<html') ||
    sample.includes('<html') ||
    (sample.includes('microsoft') && sample.includes('sharepoint'))
  );
}

function isProbablyXlsx(buffer) {
  return buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4b;
}

async function downloadExcel() {
  const errors = [];

  for (const candidate of DOWNLOAD_CANDIDATES) {
    console.log(`Intentando descarga: ${candidate.name}`);
    console.log(candidate.url);

    try {
      const response = await fetch(candidate.url, {
        method: 'GET',
        redirect: 'follow',
        headers: {
          'User-Agent': 'Mozilla/5.0 GitHubActions DashboardArriendoFIAS',
          'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*'
        }
      });

      const contentType = response.headers.get('content-type') || 'sin content-type';
      console.log(`HTTP ${response.status} ${response.statusText}`);
      console.log(`Content-Type: ${contentType}`);

      const arrayBuffer = await response.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);

      console.log(`Bytes descargados: ${buffer.length}`);

      if (!response.ok) {
        const sample = buffer.subarray(0, 1000).toString('utf8');
        errors.push(`${candidate.name}: HTTP ${response.status}. Muestra: ${sample}`);
        continue;
      }

      if (buffer.length < 1000) {
        const sample = buffer.toString('utf8');
        errors.push(`${candidate.name}: respuesta demasiado pequeña. Muestra: ${sample}`);
        continue;
      }

      if (isProbablyHtml(buffer) || contentType.toLowerCase().includes('text/html')) {
        const sample = buffer.subarray(0, 1200).toString('utf8');
        console.log('La respuesta parece HTML, no XLSX. Primeros caracteres:');
        console.log(sample);
        errors.push(`${candidate.name}: SharePoint devolvió HTML, no XLSX.`);
        continue;
      }

      if (!isProbablyXlsx(buffer)) {
        const sample = buffer.subarray(0, 500).toString('utf8');
        console.log('La respuesta no parece XLSX ZIP. Primeros caracteres:');
        console.log(sample);
        errors.push(`${candidate.name}: la respuesta no parece un XLSX válido.`);
        continue;
      }

      console.log(`Descarga válida detectada desde: ${candidate.name}`);
      return {
        buffer,
        sourceName: candidate.name,
        sourceUrl: candidate.url,
        contentType,
        bytes: buffer.length
      };
    } catch (error) {
      errors.push(`${candidate.name}: ${error.message}`);
    }
  }

  throw new Error(
    'No se pudo descargar un XLSX válido desde OneDrive/SharePoint.\n\n' +
    'Diagnóstico:\n' +
    errors.map((e, i) => `${i + 1}. ${e}`).join('\n') +
    '\n\nVerifica que el enlace sea de descarga anónima real para cualquier persona.'
  );
}

function findHeaderRow(matrix) {
  return matrix.findIndex(row => {
    if (!Array.isArray(row)) return false;

    const normalizedCells = row.map(normalizeKey);

    const hasMonth = normalizedCells.includes('mes');
    const hasYear = normalizedCells.includes('ano') || normalizedCells.includes('anio');
    const hasCanon = normalizedCells.some(c => c.includes('canon'));
    const hasExpensas = normalizedCells.some(c => c.includes('expensas') || c.includes('expensa'));

    return hasMonth && hasYear && hasCanon && hasExpensas;
  });
}

function normalizeRow(row) {
  const mes = String(valueByAliases(row, 'mes') ?? '').trim();
  const anio = Math.trunc(parseNumber(valueByAliases(row, 'anio')));
  const monthNumber = MONTHS[normalizeText(mes)] || 0;

  const canonProgramado = parseNumber(valueByAliases(row, 'canon_programado'));
  const canonEjecutado = parseNumber(valueByAliases(row, 'canon_ejecutado'));
  const expensasProgramadas = parseNumber(valueByAliases(row, 'expensas_programadas'));
  const expensasEjecutadas = parseNumber(valueByAliases(row, 'expensas_ejecutadas'));
  const serviciosBasicos = parseNumber(valueByAliases(row, 'servicios_basicos'));

  const fijoProgramado = canonProgramado + expensasProgramadas;
  const fijoEjecutado = canonEjecutado + expensasEjecutadas;

  return {
    periodo: `${mes} ${anio}`.trim(),
    mes,
    anio,
    mes_num: monthNumber,
    period_key: `${anio}-${String(monthNumber).padStart(2, '0')}`,

    canon_programado: canonProgramado,
    canon_ejecutado: canonEjecutado,
    saldo_canon: parseNumber(valueByAliases(row, 'saldo_canon')),
    estado_canon: String(valueByAliases(row, 'estado_canon') ?? '').trim(),
    factura_canon: String(valueByAliases(row, 'factura_canon') ?? '').trim(),
    fecha_pago_canon: excelDateToISO(valueByAliases(row, 'fecha_pago_canon')),
    solicitud_pago_canon: String(valueByAliases(row, 'solicitud_pago_canon') ?? '').trim(),
    link_expediente_canon: String(valueByAliases(row, 'link_expediente_canon') ?? '').trim(),

    servicios_basicos: serviciosBasicos,
    estado_servicios: String(valueByAliases(row, 'estado_servicios') ?? '').trim(),
    factura_servicios: String(valueByAliases(row, 'factura_servicios') ?? '').trim(),
    fecha_pago_servicios: excelDateToISO(valueByAliases(row, 'fecha_pago_servicios')),
    solicitud_pago_servicios: String(valueByAliases(row, 'solicitud_pago_servicios') ?? '').trim(),
    link_expediente_servicios: String(valueByAliases(row, 'link_expediente_servicios') ?? '').trim(),

    expensas_programadas: expensasProgramadas,
    expensas_ejecutadas: expensasEjecutadas,
    saldo_expensas: parseNumber(valueByAliases(row, 'saldo_expensas')),
    estado_expensas: String(valueByAliases(row, 'estado_expensas') ?? '').trim(),
    notificacion_expensas: String(valueByAliases(row, 'notificacion_expensas') ?? '').trim(),
    fecha_pago_expensas: excelDateToISO(valueByAliases(row, 'fecha_pago_expensas')),
    solicitud_pago_expensas: String(valueByAliases(row, 'solicitud_pago_expensas') ?? '').trim(),
    link_expediente_expensas: String(valueByAliases(row, 'link_expediente_expensas') ?? '').trim(),

    total_recurrente_programado: parseNumber(valueByAliases(row, 'total_recurrente_programado')),
    total_ejecutado: parseNumber(valueByAliases(row, 'total_ejecutado')),
    saldo_recurrente: parseNumber(valueByAliases(row, 'saldo_recurrente')),
    estado_mensual: String(valueByAliases(row, 'estado_mensual') ?? '').trim(),
    observacion: String(valueByAliases(row, 'observacion') ?? '').trim(),

    fijo_programado: fijoProgramado,
    fijo_ejecutado: fijoEjecutado,
    saldo_fijo: fijoProgramado - fijoEjecutado,
    desembolso_total: fijoEjecutado + serviciosBasicos
  };
}

function rowsFromWorkbook(workbook) {
  let selected = null;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];

    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: null
    });

    const headerIndex = findHeaderRow(matrix);
    console.log(`Hoja "${sheetName}": filas ${matrix.length}; encabezado detectado en índice ${headerIndex}`);

    if (headerIndex >= 0) {
      selected = { sheetName, matrix, headerIndex };
      break;
    }
  }

  if (!selected) {
    throw new Error(
      'No se encontró la fila de encabezados. Se esperaba una fila con Mes, Año, Canon y Expensas.\n' +
      `Encabezados esperados: ${EXPECTED_HEADERS.join(' | ')}`
    );
  }

  const headers = selected.matrix[selected.headerIndex].map(h => String(h ?? '').trim());

  console.log(`Hoja seleccionada: ${selected.sheetName}`);
  console.log('Encabezados detectados:');
  console.log(headers);

  const rows = selected.matrix
    .slice(selected.headerIndex + 1)
    .filter(row => Array.isArray(row) && row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ''))
    .filter(row => !String(row[0] ?? '').trim().toUpperCase().startsWith('TOTAL'))
    .map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        if (header) obj[header] = row[index];
      });
      return normalizeRow(obj);
    })
    .filter(row => row.mes && row.anio);

  return {
    rows,
    sheetName: selected.sheetName,
    headerIndex: selected.headerIndex,
    headers
  };
}

function computeResumen(rows) {
  const canon = rows.find(r => Number(r.canon_programado) > 0)?.canon_programado || 0;
  const expensa = rows.find(r => Number(r.expensas_programadas) > 0)?.expensas_programadas || 0;

  return {
    canon: Number(canon) || 0,
    expensa: Number(expensa) || 0,
    fixed: (Number(canon) || 0) + (Number(expensa) || 0),
    fixed_executed: rows.reduce((s, r) => s + (Number(r.fijo_ejecutado) || 0), 0),
    fixed_balance: rows.reduce((s, r) => s + (Number(r.saldo_fijo) || 0), 0),
    services: rows.reduce((s, r) => s + (Number(r.servicios_basicos) || 0), 0),
    total: rows.reduce((s, r) => s + (Number(r.desembolso_total) || 0), 0),
    complete: rows.filter(r => r.estado_mensual === 'Completo').length,
    partial: rows.filter(r => r.estado_mensual === 'Parcial').length,
    pending: rows.filter(r => r.estado_mensual === 'Pendiente').length,
    months: rows.length
  };
}

async function main() {
  console.log('Descargando Excel público de contrato de arriendo desde OneDrive/SharePoint...');

  await fs.mkdir(OUT_DIR, { recursive: true });

  const download = await downloadExcel();

  console.log('Leyendo libro XLSX...');
  const workbook = XLSX.read(download.buffer, {
    type: 'buffer',
    cellDates: true
  });

  console.log(`Hojas encontradas: ${workbook.SheetNames.join(', ')}`);

  const parsed = rowsFromWorkbook(workbook);

  if (!parsed.rows.length) {
    throw new Error('No se generaron registros válidos desde el Excel.');
  }

  const payload = {
    registros: parsed.rows,
    resumen: computeResumen(parsed.rows)
  };

  const meta = {
    generatedAt: new Date().toISOString(),
    records: parsed.rows.length,
    sheetName: parsed.sheetName,
    headerRowIndex: parsed.headerIndex,
    headers: parsed.headers,
    sourceName: download.sourceName,
    sourceUrl: download.sourceUrl,
    contentType: download.contentType,
    bytes: download.bytes
  };

  await fs.writeFile(OUT_JSON, JSON.stringify(payload, null, 2), 'utf8');
  await fs.writeFile(OUT_META, JSON.stringify(meta, null, 2), 'utf8');

  console.log(`JSON generado correctamente: ${OUT_JSON}`);
  console.log(`Registros: ${parsed.rows.length}`);
  console.log(`Metadata generada: ${OUT_META}`);
}

main().catch(error => {
  console.error('Error generando data/control-arriendo.json');
  console.error(error);
  process.exit(1);
});
