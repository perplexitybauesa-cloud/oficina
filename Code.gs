// ============================================================
//  SEGUIMENT D'USUARIS — Web App Google Apps Script
//  Code.gs — Backend / Servidor
// ============================================================

const SPREADSHEET_ID = '13-HCGphpkB8zi4m6299YdMXEADWtXLc0Hy1CJot_Cj0';

// ─────────────────────────────────────────────────────────────
//  PUNT D'ENTRADA
// ─────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CRM · OPE Sant Adrià de Besòs')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─────────────────────────────────────────────────────────────
//  UTILITATS INTERNES
// ─────────────────────────────────────────────────────────────
function getSheet_(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

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

function nextId_(ws) {
  // Solución al bug de generación de IDs para inserciones múltiples
  const data = ws.getRange("A2:A" + Math.max(ws.getLastRow(), 2)).getValues();
  let max = 0;
  for (let i = 0; i < data.length; i++) {
    let v = Number(data[i][0]);
    if (v > max) max = v;
  }
  return max + 1;
}

// ─────────────────────────────────────────────────────────────
//  1. DASHBOARD — Estadístiques generals
// ─────────────────────────────────────────────────────────────
function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const wsU = ss.getSheetByName('USUARIS');
    const wsA = ss.getSheetByName('ATENCIONS');

    const totalUsuaris   = Math.max(wsU.getLastRow() - 1, 0);
    const totalAtencions = Math.max(wsA.getLastRow() - 1, 0);

    const atData = wsA.getDataRange().getValues();
    const hA     = atData[0];
    const fechaIdx     = hA.indexOf('FECHA');
    const quiRepIdx    = hA.indexOf('QUI_REP');
    const derivacioIdx = hA.indexOf('DERIVACIÓ');
    const dniIdxA      = hA.indexOf('DNI');

    const ara = new Date();
    const mesActual = ara.getMonth();
    const anyActual = ara.getFullYear();

    const rows = atData.slice(1);

    const atencionesMes = rows.filter(r => {
      const d = new Date(r[fechaIdx]);
      return !isNaN(d) && d.getMonth() === mesActual && d.getFullYear() === anyActual;
    }).length;
    const perTecnic = {};
    const perDerivacio = {};
    rows.forEach(r => {
      const t = r[quiRepIdx]    || 'Altres';
      const d = r[derivacioIdx] || 'Altres';
      perTecnic[t]    = (perTecnic[t]    || 0) + 1;
      perDerivacio[d] = (perDerivacio[d] || 0) + 1;
    });
    
    const uData  = wsU.getDataRange().getValues();
    const uH     = uData[0];
    const dniIdxU = uH.indexOf('DNI');
    const nomIdxU = uH.indexOf('NOM');

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
//  2. CERCAR USUARI PER DNI
// ─────────────────────────────────────────────────────────────
function buscarUsuariPerDNI(dni) {
  try {
    dni = dni.toUpperCase().trim();
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const wsU  = ss.getSheetByName('USUARIS');
    const uData = wsU.getDataRange().getValues();
    const uH    = uData[0];
    const dniIdx = uH.indexOf('DNI');
    const filaIdx = uData.findIndex((r, i) => i > 0 && r[dniIdx] === dni);
    
    if (filaIdx === -1) return { trobat: false };

    const obj = {};
    uH.forEach((h, i) => {
      const v = uData[filaIdx][i];
      obj[h] = v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy') : v;
    });
    
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
//  3. GUARDAR NOVA ATENCIÓ (Usuari + Atenció + Gestió)
// ─────────────────────────────────────────────────────────────
function guardarAtencio(fd) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dni = (fd.dni || '').toUpperCase().trim();
    if (!dni) return { ok: false, error: 'DNI obligatori' };
    
    const wsU   = ss.getSheetByName('USUARIS');
    const uData = wsU.getDataRange().getValues();
    const uH    = uData[0];
    const existingRow = uData.find((r, i) => i > 0 && r[uH.indexOf('DNI')] === dni);
    const nom    = fd.nom || (existingRow ? existingRow[uH.indexOf('NOM')] : '') || '';
    const enlace = existingRow ? (existingRow[uH.indexOf('ENLACE')] || '') : '';

    _guardarUsuari(ss, fd, dni, nom, enlace, uData, uH);
    const idAtencio = _guardarAtencio(ss, fd, dni, nom, enlace);

    if (fd.assessoraments && fd.assessoraments.length)
      _guardarAssessoraments(ss, fd.assessoraments, dni, nom, enlace, idAtencio);
      
    _guardarGestions(ss, fd.assessoraments || [], dni, nom, enlace, idAtencio);

    return { ok: true, idAtencio, missatge: `Atenció #${idAtencio} registrada correctament` };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  HELPERS D'ESCRIPTURA
// ─────────────────────────────────────────────────────────────

function _guardarUsuari(ss, fd, dni, nom, enlace, uData, uH) {
  const ws   = ss.getSheetByName('USUARIS');
  if (!uData) { uData = ws.getDataRange().getValues(); uH = uData[0]; }
  const dniIdx = uH.indexOf('DNI');
  const fila   = uData.findIndex((r, i) => i > 0 && r[dniIdx] === dni);
  const vals = uH.map(h => {
    switch (h) {
      case 'ID_USUARI':               return fila === -1 ? nextId_(ws) : uData[fila][0];
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
  if (fila === -1) ws.appendRow(vals);
  else ws.getRange(fila + 1, 1, 1, vals.length).setValues([vals]);
}

function _guardarAtencio(ss, fd, dni, nom, enlace) {
  const ws      = ss.getSheetByName('ATENCIONS');
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const id      = nextId_(ws);
  const circs   = Array.isArray(fd.circumstancies)
    ? fd.circumstancies.join(', ')
    : (fd.circumstancies || '');
    
  const vals = headers.map(h => {
    switch (h) {
      case 'ID_ATENCIO':              return id;
      case 'DNI':                     return dni;
      case 'NOM':                     return nom;
      case 'ENLACE':                  return enlace || '';
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

function _guardarAssessoraments(ss, assessos, dni, nom, enlace, idAtencio) {
  const ws      = ss.getSheetByName('ASSESSORAMENTS');
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const now     = new Date();
  
  let currentId = nextId_(ws); // Base ID para evitar duplicados en bulk
  
  assessos.forEach((a, i) => {
    if (!a) return;
    const vals = headers.map(h => {
      switch (h) {
        case 'ID_ASSESSORAMENT': return currentId;
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
    currentId++;
  });
}

function _guardarGestions(ss, assessos, dni, nom, enlace, idAtencio) {
  const now = new Date();
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
    'Gestió de trams':                'GESTIÓ_TRAMS',
    'Canon social':                   'CANON_SOCIAL',
    'Treva gas':                      'TREVA_GAS'
  };
  const AIGUA_MAP = {
    'Canvi titularitat aigua':        'DATA_CANVI_TIT_AIGUA',
    'Ajuda MUSA (comptadors)':        'AJUDA_MUSA',
    'Tall efectiu aigua':             'TALL_EFECTIU_AIGUA',
    'Punxada aigua':                  'PUNXADA_AIGUA',
    'Alta aigua':                     'ALTA_AIGUA',
    'Negociar deute aigua':           'NEGOCIAR_DEUTE',
    'IRER enviat':                    'IRER_ENVIAT',
    'Treva aigua':                    'TREVA_AIGUA'
  };
  const ALTRES_MAP = {
    'Projecte Porta a Porta':         'PROJECTE_PORTA_A_PORTA',
    'ONA':                            'ONA'
  };
  
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
          const assName = Object.keys(map).find(a => map[a] === h);
          return assName && matched.includes(assName) ? now : '';
        }
      }
    });

    ws.appendRow(vals);
  });
}

// ─────────────────────────────────────────────────────────────
//  4. ÚLTIMES ATENCIONS (per llistat general)
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
    
    const dniIdxU = uH.indexOf('DNI');
    const nomIdxU = uH.indexOf('NOM');
    const fechaIdx = aH.indexOf('FECHA');
    const dniIdxA  = aH.indexOf('DNI');
    const derivIdx = aH.indexOf('DERIVACIÓ');
    const quiIdx   = aH.indexOf('QUI_REP');
    const calIdx   = aH.indexOf('CAL_ASSESSORAMENT');
    
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
