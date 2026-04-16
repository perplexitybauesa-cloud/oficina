// ============================================================
//  SEGUIMIENTO DE USUARIOS — Web App Google Apps Script
//  Code.gs — Backend / Servidor
//  Oficina de Pobreza y Eficiencia Energetica - Sant Adria de Besos
// ============================================================

const SPREADSHEET_ID = '13-HCGphpkB8zi4m6299YdMXEADWtXLc0Hy1CJot_Cj0';

// ─────────────────────────────────────────────────────────────
//  PUNTO DE ENTRADA - Renderiza la interfaz HTML
// ─────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Oficina Pobresa i Eficiencia Energetica')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─────────────────────────────────────────────────────────────
//  UTILIDADES INTERNAS
// ─────────────────────────────────────────────────────────────

/**
 * Obtiene una hoja del spreadsheet por nombre
 * @param {string} name - Nombre de la hoja
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Hoja del spreadsheet
 */
function getSheet_(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

/**
 * Convierte los datos de una hoja en un array de objetos
 * @param {GoogleAppsScript.Spreadsheet.Sheet} ws - Hoja del spreadsheet
 * @returns {Object[]} Array de objetos con los datos
 */
function sheetToObjects_(ws) {
  const data = ws.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

/**
 * Obtiene el siguiente ID disponible para una fila
 * @param {GoogleAppsScript.Spreadsheet.Sheet} ws - Hoja del spreadsheet
 * @returns {number} Siguiente ID disponible
 */
function nextId_(ws) {
  return Math.max(ws.getLastRow(), 1);
}

// ─────────────────────────────────────────────────────────────
//  1. DASHBOARD — Estadisticas generales
//  Obtiene KPIs, graficos y ultimas atenciones para el panel principal
// ─────────────────────────────────────────────────────────────
function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const wsU = ss.getSheetByName('USUARIS');
    const wsA = ss.getSheetByName('ATENCIONS');

    // Conteo total de usuarios y atenciones
    const totalUsuaris   = Math.max(wsU.getLastRow() - 1, 0);
    const totalAtencions = Math.max(wsA.getLastRow() - 1, 0);

    // Obtener datos de atenciones para analisis
    const atData = wsA.getDataRange().getValues();
    const hA     = atData[0];
    const fechaIdx     = hA.indexOf('FECHA');
    const quiRepIdx    = hA.indexOf('QUI_REP');
    const derivacioIdx = hA.indexOf('DERIVACIÓ');
    const dniIdxA      = hA.indexOf('DNI');

    // Calcular atenciones del mes actual
    const ara = new Date();
    const mesActual = ara.getMonth();
    const anyActual = ara.getFullYear();

    const rows = atData.slice(1);

    const atencionesMes = rows.filter(r => {
      const d = new Date(r[fechaIdx]);
      return !isNaN(d) && d.getMonth() === mesActual && d.getFullYear() === anyActual;
    }).length;

    // Agrupar atenciones por tecnico y por derivacion
    const perTecnic = {};
    const perDerivacio = {};
    rows.forEach(r => {
      const t = r[quiRepIdx]    || 'Altres';
      const d = r[derivacioIdx] || 'Altres';
      perTecnic[t]    = (perTecnic[t]    || 0) + 1;
      perDerivacio[d] = (perDerivacio[d] || 0) + 1;
    });

    // Obtener datos de usuarios para enlazar nombres
    const uData  = wsU.getDataRange().getValues();
    const uH     = uData[0];
    const dniIdxU = uH.indexOf('DNI');
    const nomIdxU = uH.indexOf('NOM');

    // Obtener las ultimas 8 atenciones con datos del usuario
    const ultimes = rows.slice(-8).reverse().map(r => {
      const uRow = uData.find((u, i) => i > 0 && u[dniIdxU] === r[dniIdxA]);
      const data_ = r[fechaIdx];
      return {
        id:        r[0],
        dni:       r[dniIdxA] || '-',
        nom:       uRow ? uRow[nomIdxU] : '-',
        fecha:     data_ instanceof Date ? Utilities.formatDate(data_, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (data_ || '-'),
        derivacio: r[derivacioIdx] || '-',
        quiRep:    r[quiRepIdx]    || '-'
      };
    });

    return { ok: true, totalUsuaris, totalAtencions, atencionesMes, perTecnic, perDerivacio, ultimes };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  2. BUSCAR USUARIO POR DNI
//  Busca un usuario por su DNI/NIE y devuelve sus datos e historial
// ─────────────────────────────────────────────────────────────
function buscarUsuariPerDNI(dni) {
  try {
    dni = dni.toUpperCase().trim();
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const wsU  = ss.getSheetByName('USUARIS');
    const uData = wsU.getDataRange().getValues();
    const uH    = uData[0];
    const dniIdx = uH.indexOf('DNI');

    // Buscar la fila del usuario
    const filaIdx = uData.findIndex((r, i) => i > 0 && r[dniIdx] === dni);
    if (filaIdx === -1) return { trobat: false };

    // Construir objeto con datos del usuario
    const obj = {};
    uH.forEach((h, i) => {
      const v = uData[filaIdx][i];
      obj[h] = v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy') : v;
    });

    // Obtener historial de atenciones del usuario
    const wsA    = ss.getSheetByName('ATENCIONS');
    const aData  = wsA.getDataRange().getValues();
    const aH     = aData[0];
    const dniIdxA = aH.indexOf('DNI');
    const fechaIdx = aH.indexOf('FECHA');
    const derivIdx = aH.indexOf('DERIVACIÓ');
    const quiIdx   = aH.indexOf('QUI_REP');

    const historial = aData.slice(1)
      .filter(r => r[dniIdxA] === dni)
      .map(r => {
        const d = r[fechaIdx];
        return {
          id:        r[0],
          fecha:     d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (d || '-'),
          derivacio: r[derivIdx] || '-',
          quiRep:    r[quiIdx]   || '-',
          observ:    r[aH.indexOf('OBSERVACIONS')] || '-'
        };
      })
      .reverse()
      .slice(0, 10);

    // Obtener asesoramientos previos del usuario
    const wsAs  = ss.getSheetByName('ASSESSORAMENTS');
    const asData = wsAs.getDataRange().getValues();
    const asH    = asData[0];
    const dniIdxAs = asH.indexOf('DNI');
    const assessIdx = asH.indexOf('ASSESSORAMENT');
    const assessos  = asData.slice(1)
      .filter(r => r[dniIdxAs] === dni)
      .map(r => r[assessIdx])
      .slice(-10);

    return { trobat: true, usuari: obj, historial, totalAtencions: historial.length, assessos };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  3. GUARDAR NUEVA ATENCION (Usuario + Atencion + Gestion)
//  Guarda o actualiza datos del usuario, registra la atencion
//  y los asesoramientos y gestiones asociadas
// ─────────────────────────────────────────────────────────────
function guardarAtencio(fd) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dni = (fd.dni || '').toUpperCase().trim();
    if (!dni) return { ok: false, error: 'DNI obligatori' };

    // Obtener NOM y ENLACE existentes (si el usuario ya existe)
    const wsU   = ss.getSheetByName('USUARIS');
    const uData = wsU.getDataRange().getValues();
    const uH    = uData[0];
    const existingRow = uData.find((r, i) => i > 0 && r[uH.indexOf('DNI')] === dni);
    const nom    = fd.nom || (existingRow ? existingRow[uH.indexOf('NOM')]    : '') || '';
    const enlace = existingRow ? (existingRow[uH.indexOf('ENLACE')] || '') : '';

    // Guardar/actualizar datos del usuario
    _guardarUsuari(ss, fd, dni, nom, enlace, uData, uH);
    
    // Guardar la atencion
    const idAtencio = _guardarAtencio(ss, fd, dni, nom);

    // Guardar asesoramientos si existen
    if (fd.assessoraments && fd.assessoraments.length)
      _guardarAssessoraments(ss, fd.assessoraments, dni, nom, enlace, idAtencio);

    // Guardar gestiones en las hojas correspondientes
    _guardarGestions(ss, fd.assessoraments || [], dni, nom, enlace, idAtencio);

    return { ok: true, idAtencio, missatge: `Atenció #${idAtencio} registrada correctament` };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  FUNCIONES AUXILIARES DE ESCRITURA
// ─────────────────────────────────────────────────────────────

/**
 * Guarda o actualiza los datos de un usuario en la hoja USUARIS
 * Columnas: ID_USUARI | DNI | NOM | ENLACE | GENERE | LLOC DE NAIXEMENT |
 *           BARRI | TINENÇA | ADREÇA | TELF. 1 | E-MAIL | TOTAL PERSONES HABITATGE
 */
function _guardarUsuari(ss, fd, dni, nom, enlace, uData, uH) {
  const ws   = ss.getSheetByName('USUARIS');
  if (!uData) { uData = ws.getDataRange().getValues(); uH = uData[0]; }
  const dniIdx = uH.indexOf('DNI');
  const fila   = uData.findIndex((r, i) => i > 0 && r[dniIdx] === dni);

  // Mapear valores segun las cabeceras
  const vals = uH.map(h => {
    switch (h) {
      case 'ID_USUARI':               return fila === -1 ? ws.getLastRow() : uData[fila][0];
      case 'DNI':                     return dni;
      case 'NOM':                     return nom;
      case 'ENLACE':                  return enlace || (fila > -1 ? uData[fila][uH.indexOf('ENLACE')] : '') || '';
      case 'GÈNERE':                  return fd.genere   || '';
      case 'LLOC DE NAIXEMENT':       return fd.llocNaix || '';
      case 'BARRI':                   return fd.barri    || '';
      case 'TINENÇA':                 return fd.tinenca  || '';
      case 'ADREÇA':                  return fd.adreca   || '';
      case 'TELF. 1':                 return fd.telf     || '';
      case 'E-MAIL':                  return fd.email    || '';
      case 'TOTAL PERSONES HABITATGE':return fd.persones || '';
      default: return fila > -1 ? uData[fila][uH.indexOf(h)] : '';
    }
  });

  // Insertar nueva fila o actualizar existente
  if (fila === -1) ws.appendRow(vals);
  else ws.getRange(fila + 1, 1, 1, vals.length).setValues([vals]);
}

/**
 * Guarda una nueva atencion en la hoja ATENCIONS
 * Columnas: ID_ATENCIO | DNI | NOM | ENLACE | FECHA | DERIVACIO | QUI_REP |
 *           OBSERVACIONS | REFERENT | CIRCUMSTANCIES_ESPECIALS | CAL_ASSESSORAMENT
 */
function _guardarAtencio(ss, fd, dni, nom) {
  const ws      = ss.getSheetByName('ATENCIONS');
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const id      = nextId_(ws);
  
  // Convertir array de circunstancias a string separado por comas
  const circs   = Array.isArray(fd.circumstancies)
    ? fd.circumstancies.join(', ')
    : (fd.circumstancies || '');

  const vals = headers.map(h => {
    switch (h) {
      case 'ID_ATENCIO':              return id;
      case 'DNI':                     return dni;
      case 'NOM':                     return nom;
      case 'ENLACE':                  return '';
      case 'FECHA':                   return new Date();
      case 'DERIVACIÓ':               return fd.derivacio       || '';
      case 'QUI_REP':                 return fd.quiRep          || '';
      case 'OBSERVACIONS':            return fd.observacions    || '';
      case 'REFERENT':                return fd.referent        || '';
      case 'CIRCUMSTÀNCIES_ESPECIALS':return circs;
      case 'CAL_ASSESSORAMENT':       return fd.calAssessorament|| '';
      default: return '';
    }
  });

  ws.appendRow(vals);
  return id;
}

/**
 * Guarda los asesoramientos realizados en la hoja ASSESSORAMENTS
 * Columnas: ID_ASSESSORAMENT | ID_ATENCIO | DNI | NOM | ENLACE | FECHA | ORDRE | ASSESSORAMENT
 */
function _guardarAssessoraments(ss, assessos, dni, nom, enlace, idAtencio) {
  const ws      = ss.getSheetByName('ASSESSORAMENTS');
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const now     = new Date();

  assessos.forEach((a, i) => {
    if (!a) return;
    const id = nextId_(ws);
    const vals = headers.map(h => {
      switch (h) {
        case 'ID_ASSESSORAMENT': return id;
        case 'ID_ATENCIO':       return idAtencio;
        case 'DNI':              return dni;
        case 'NOM':              return nom;
        case 'ENLACE':           return enlace || '';
        case 'FECHA':            return now;
        case 'ORDRE':            return i + 1;
        case 'ASSESSORAMENT':    return a;
        default: return '';
      }
    });
    ws.appendRow(vals);
  });
}

/**
 * Guarda las gestiones en las hojas correspondientes segun el tipo
 * Escribe en: GESTIONS_LLUM, GESTIONS_GAS, GESTIONS_AIGUA, ALTRES_GESTIONS
 * basandose en las cabeceras reales de cada hoja
 */
function _guardarGestions(ss, assessos, dni, nom, enlace, idAtencio) {
  const now = new Date();

  // Mapeo de asesoramiento a nombre de columna en la hoja
  const LLUM_MAP = {
    'Canvi titularitat llum':         'CANVI_TITULARITAT',
    'Canvi comercialitzadora llum':   'CANVI_COMERCIALITZADORA',
    'Mercat lliure a mercat regulat': 'DATA_ML_MR',
    'Potència reduïda':               'POTÈNCIA_REDUÏDA',
    'Eliminació serveis extres llum': 'ELIMINACIÓ_SERVEIS_EXTRES',
    'Bo social':                      'BO_SOCIAL_DATA',
    'Nou bo social':                  'NOU_BO_SOCIAL_DATA',
    'Tall efectiu llum':              'TALL_EFECTIU_LLUM',
    'Bo tèrmic':                      'BO_TERMIC',
    'Punxada llum':                   'PUNXADA_LLUM',
    'Desconnexió llum':               'DESCONNEXIÓ',
    'Negociar deute llum':            'NEGOCIAR_DEUTE_LLUM',
    'Alta llum':                      'ALTA_LLUM',
    'Treva llum':                     'TREVA_LLUM'
  };

  const GAS_MAP = {
    'Canvi titularitat gas':          'DATA_CANVI_TIT_GAS',
    'Canvi comercialitzadora gas':    'CANVI_COMERCIALITZADORA',
    'Mercat lliure a TUR':            'ML_A_TUR',
    'Eliminació serveis extres gas':  'ELIMINACIÓ_SERVEIS_EXTRES',
    'Tall efectiu gas':               'TALL_EFECTIU_GAS',
    'Punxada gas':                    'PUNXADA_GAS',
    'Desconnexió gas':                'DESCONNEXIÓ',
    'Gestió de trams':                'GESTIÓ_TRAMS',
    'Canon social':                   'CANON_SOCIAL',
    'Treva gas':                      'TREVA_GAS'
  };

  const AIGUA_MAP = {
    'Canvi titularitat aigua':        'DATA_CANVI_TIT_AIGUA',
    'Ajuda MUSA (comptadors)':        'AJUDA_MUSA',
    'Tall efectiu aigua':             'TALL_EFECTIU_AIGUA',
    'Punxada aigua':                  'PUNXADA_AIGUA',
    'Desconnexió aigua':              'DESCONNEXIÓ',
    'Alta aigua':                     'ALTA_AIGUA',
    'Negociar deute aigua':           'NEGOCIAR_DEUTE',
    'IRER enviat':                    'IRER_ENVIAT',
    'Treva aigua':                    'TREVA_AIGUA'
  };

  const ALTRES_MAP = {
    'Projecte Porta a Porta':         'PROJECTE_PORTA_A_PORTA',
    'ONA':                            'ONA'
  };

  // Iterar sobre cada tipo de gestion y escribir en la hoja correspondiente
  [
    [LLUM_MAP,   'GESTIONS_LLUM'],
    [GAS_MAP,    'GESTIONS_GAS'],
    [AIGUA_MAP,  'GESTIONS_AIGUA'],
    [ALTRES_MAP, 'ALTRES_GESTIONS']
  ].forEach(([map, sheetName]) => {
    const matched = assessos.filter(a => map[a]);
    if (!matched.length) return;

    const ws      = ss.getSheetByName(sheetName);
    const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];

    const vals = headers.map(h => {
      switch (h) {
        case 'ID_ATENCIO': return idAtencio;
        case 'DNI':        return dni;
        case 'NOM':        return nom;
        case 'ENLACE':     return enlace || '';
        case 'FECHA':      return now;
        default: {
          // Si la cabecera corresponde a un asesoramiento seleccionado, poner fecha
          const assName = Object.keys(map).find(a => map[a] === h);
          return assName && matched.includes(assName) ? now : '';
        }
      }
    });

    ws.appendRow(vals);
  });
}

// ─────────────────────────────────────────────────────────────
//  4. ULTIMAS ATENCIONES (para listado general)
//  Obtiene las ultimas N atenciones para mostrar en el listado
// ─────────────────────────────────────────────────────────────
function getUltimesAtencions(limit) {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const wsA  = ss.getSheetByName('ATENCIONS');
    const wsU  = ss.getSheetByName('USUARIS');

    const aData = wsA.getDataRange().getValues();
    const aH    = aData[0];
    const uData = wsU.getDataRange().getValues();
    const uH    = uData[0];
    
    // Indices de columnas
    const dniIdxU = uH.indexOf('DNI');
    const nomIdxU = uH.indexOf('NOM');
    const fechaIdx = aH.indexOf('FECHA');
    const dniIdxA  = aH.indexOf('DNI');
    const derivIdx = aH.indexOf('DERIVACIÓ');
    const quiIdx   = aH.indexOf('QUI_REP');
    const calIdx   = aH.indexOf('CAL_ASSESSORAMENT');

    // Obtener las ultimas N atenciones con datos del usuario
    const rows = aData.slice(1).slice(-(limit || 30)).reverse().map(r => {
      const uRow = uData.find((u, i) => i > 0 && u[dniIdxU] === r[dniIdxA]);
      const d = r[fechaIdx];
      return {
        id:        r[0],
        dni:       r[dniIdxA] || '-',
        nom:       uRow ? uRow[nomIdxU] : (r[aH.indexOf('NOM')] || '-'),
        fecha:     d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (d || '-'),
        derivacio: r[derivIdx] || '-',
        quiRep:    r[quiIdx]   || '-',
        calAss:    r[calIdx]   || '-'
      };
    });

    return { ok: true, rows };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}
