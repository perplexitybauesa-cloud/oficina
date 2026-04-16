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

/**
 * Formatea una fecha para mostrar
 * @param {Date} date - Fecha a formatear
 * @returns {string} Fecha formateada dd/MM/yyyy
 */
function formatDate_(date) {
  if (!(date instanceof Date) || isNaN(date)) return '-';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
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
    const calAssIdx    = hA.indexOf('CAL_ASSESSORAMENT');

    // Calcular atenciones del mes actual y mes anterior
    const ara = new Date();
    const mesActual = ara.getMonth();
    const anyActual = ara.getFullYear();
    const mesAnterior = mesActual === 0 ? 11 : mesActual - 1;
    const anyAnterior = mesActual === 0 ? anyActual - 1 : anyActual;

    const rows = atData.slice(1);

    let atencionesMes = 0;
    let atencionesMesAnterior = 0;

    // Preparar tendencia mensual ultimos 6 meses
    const MONTH_NAMES = ['Gen','Feb','Mar','Abr','Mai','Jun','Jul','Ago','Set','Oct','Nov','Des'];
    const perMes = {};
    for (let i = 5; i >= 0; i--) {
      const d = new Date(anyActual, mesActual - i, 1);
      const key = d.getFullYear() + '-' + d.getMonth();
      perMes[key] = { label: MONTH_NAMES[d.getMonth()], count: 0, isCurrent: i === 0 };
    }

    rows.forEach(r => {
      const d = new Date(r[fechaIdx]);
      if (!isNaN(d)) {
        if (d.getMonth() === mesActual && d.getFullYear() === anyActual) {
          atencionesMes++;
        }
        if (d.getMonth() === mesAnterior && d.getFullYear() === anyAnterior) {
          atencionesMesAnterior++;
        }
        const key = d.getFullYear() + '-' + d.getMonth();
        if (perMes[key]) perMes[key].count++;
      }
    });

    const tendenciaMensual = Object.values(perMes);

    // Agrupar atenciones por tecnico y por derivacion
    const perTecnic = {};
    const perDerivacio = {};
    rows.forEach(r => {
      const t = r[quiRepIdx]    || 'Altres';
      const d = r[derivacioIdx] || 'Altres';
      perTecnic[t]    = (perTecnic[t]    || 0) + 1;
      perDerivacio[d] = (perDerivacio[d] || 0) + 1;
    });

    // Contar gestions pendents (atenciones con CAL_ASSESSORAMENT = 'Pendent')
    const pendingGestions = rows.filter(r => 
      r[calAssIdx] === 'Pendent' || r[calAssIdx] === 'Si'
    ).length;

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
        fecha:     formatDate_(data_),
        derivacio: r[derivacioIdx] || '-',
        quiRep:    r[quiRepIdx]    || '-'
      };
    });

    return { 
      ok: true, 
      totalUsuaris, 
      totalAtencions, 
      atencionesMes,
      atencionesMesAnterior,
      pendingGestions,
      perTecnic, 
      perDerivacio, 
      ultimes,
      tendenciaMensual
    };
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
      obj[h] = v instanceof Date ? formatDate_(v) : v;
    });

    // Obtener historial de atenciones del usuario
    const wsA    = ss.getSheetByName('ATENCIONS');
    const aData  = wsA.getDataRange().getValues();
    const aH     = aData[0];
    const dniIdxA = aH.indexOf('DNI');
    const fechaIdx = aH.indexOf('FECHA');
    const derivIdx = aH.indexOf('DERIVACIÓ');
    const quiIdx   = aH.indexOf('QUI_REP');
    const obsIdx   = aH.indexOf('OBSERVACIONS');

    // Obtener assessoraments para mapear a las atenciones
    const wsAs   = ss.getSheetByName('ASSESSORAMENTS');
    const asData = wsAs.getDataRange().getValues();
    const asH    = asData[0];
    const dniIdxAs = asH.indexOf('DNI');
    const idAtIdxAs = asH.indexOf('ID_ATENCIO');
    const assessIdx = asH.indexOf('ASSESSORAMENT');

    const historial = aData.slice(1)
      .filter(r => r[dniIdxA] === dni)
      .map(r => {
        const idAtencio = r[0];
        const d = r[fechaIdx];
        // Obtener assessoraments de esta atencion
        const assessoraments = asData.slice(1)
          .filter(a => a[idAtIdxAs] === idAtencio)
          .map(a => a[assessIdx]);
        
        return {
          id:        idAtencio,
          fecha:     formatDate_(d),
          derivacio: r[derivIdx] || '-',
          quiRep:    r[quiIdx]   || '-',
          observ:    r[obsIdx] || '-',
          assessoraments: assessoraments
        };
      })
      .reverse()
      .slice(0, 15);

    // Obtener notas del usuario (si existe la hoja NOTES)
    let notes = [];
    try {
      const wsN = ss.getSheetByName('NOTES');
      if (wsN) {
        const nData = wsN.getDataRange().getValues();
        const nH = nData[0];
        const dniIdxN = nH.indexOf('DNI');
        const textIdx = nH.indexOf('TEXT');
        const tipusIdx = nH.indexOf('TIPUS');
        const fechaIdxN = nH.indexOf('FECHA');
        const autorIdx = nH.indexOf('AUTOR');
        
        notes = nData.slice(1)
          .filter(r => r[dniIdxN] === dni)
          .map(r => ({
            text: r[textIdx],
            tipus: r[tipusIdx] || 'info',
            fecha: formatDate_(r[fechaIdxN]),
            autor: r[autorIdx] || '-'
          }))
          .reverse()
          .slice(0, 5);
      }
    } catch(e) {
      // Si no existe la hoja NOTES, continuar sin notas
    }

    return { 
      trobat: true, 
      usuari: obj, 
      historial, 
      totalAtencions: historial.length,
      notes
    };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  3. OBTENER TODOS LOS USUARIOS (para busqueda global)
//  Devuelve una lista simplificada de todos los usuarios
// ─────────────────────────────────────────────────────────────
function getAllUsers() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const wsU = ss.getSheetByName('USUARIS');
    const uData = wsU.getDataRange().getValues();
    const uH = uData[0];
    const dniIdx = uH.indexOf('DNI');
    const nomIdx = uH.indexOf('NOM');
    
    const users = uData.slice(1).map(r => ({
      DNI: r[dniIdx],
      NOM: r[nomIdx]
    }));
    
    return { ok: true, users };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  4. GUARDAR NUEVA ATENCION (Usuario + Atencion + Gestion)
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

    // Guardar nota interna si existe
    if (fd.notaInterna && fd.notaInterna.trim()) {
      _guardarNotaInterna(ss, dni, fd.notaInterna, fd.notaAlerta ? 'alert' : 'info', fd.quiRep);
    }

    return { ok: true, idAtencio, missatge: `Atencio #${idAtencio} registrada correctament` };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  FUNCIONES AUXILIARES DE ESCRITURA
// ─────────────────────────────────────────────────────────────

/**
 * Guarda o actualiza los datos de un usuario en la hoja USUARIS
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
 */
function _guardarGestions(ss, assessos, dni, nom, enlace, idAtencio) {
  const now = new Date();

  // Mapeo de asesoramiento a nombre de columna en la hoja
  const LLUM_MAP = {
    'Canvi titularitat llum':         'CANVI_TITULARITAT',
    'Canvi comercialitzadora llum':   'CANVI_COMERCIALITZADORA',
    'Mercat lliure a mercat regulat': 'DATA_ML_MR',
    'Potencia reduida':               'POTÈNCIA_REDUÏDA',
    'Eliminacio serveis extres llum': 'ELIMINACIÓ_SERVEIS_EXTRES',
    'Bo social':                      'BO_SOCIAL_DATA',
    'Nou bo social':                  'NOU_BO_SOCIAL_DATA',
    'Tall efectiu llum':              'TALL_EFECTIU_LLUM',
    'Bo termic':                      'BO_TERMIC',
    'Punxada llum':                   'PUNXADA_LLUM',
    'Desconnexio llum':               'DESCONNEXIÓ',
    'Negociar deute llum':            'NEGOCIAR_DEUTE_LLUM',
    'Alta llum':                      'ALTA_LLUM',
    'Treva llum':                     'TREVA_LLUM'
  };

  const GAS_MAP = {
    'Canvi titularitat gas':          'DATA_CANVI_TIT_GAS',
    'Canvi comercialitzadora gas':    'CANVI_COMERCIALITZADORA',
    'Mercat lliure a TUR':            'ML_A_TUR',
    'Eliminacio serveis extres gas':  'ELIMINACIÓ_SERVEIS_EXTRES',
    'Tall efectiu gas':               'TALL_EFECTIU_GAS',
    'Punxada gas':                    'PUNXADA_GAS',
    'Desconnexio gas':                'DESCONNEXIÓ',
    'Gestio de trams':                'GESTIÓ_TRAMS',
    'Canon social':                   'CANON_SOCIAL',
    'Treva gas':                      'TREVA_GAS'
  };

  const AIGUA_MAP = {
    'Canvi titularitat aigua':        'DATA_CANVI_TIT_AIGUA',
    'Ajuda MUSA (comptadors)':        'AJUDA_MUSA',
    'Tall efectiu aigua':             'TALL_EFECTIU_AIGUA',
    'Punxada aigua':                  'PUNXADA_AIGUA',
    'Desconnexio aigua':              'DESCONNEXIÓ',
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

/**
 * Guarda una nota interna en la hoja NOTES
 */
function _guardarNotaInterna(ss, dni, text, tipus, autor) {
  let ws = ss.getSheetByName('NOTES');
  
  // Crear la hoja si no existe
  if (!ws) {
    ws = ss.insertSheet('NOTES');
    ws.appendRow(['ID_NOTA', 'DNI', 'TEXT', 'TIPUS', 'FECHA', 'AUTOR']);
  }
  
  const id = nextId_(ws);
  ws.appendRow([id, dni, text, tipus, new Date(), autor || '']);
}

// ─────────────────────────────────────────────────────────────
//  5. GUARDAR NOTA (desde modal)
// ─────────────────────────────────────────────────────────────
function guardarNota(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dni = (data.dni || '').toUpperCase().trim();
    
    if (!dni || !data.text) {
      return { ok: false, error: 'Falten dades' };
    }
    
    _guardarNotaInterna(ss, dni, data.text, data.tipus || 'info', Session.getActiveUser().getEmail());
    
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  6. GESTIONS PENDENTS
//  Obtiene las gestiones que requieren seguimiento
// ─────────────────────────────────────────────────────────────
function getGestionsPendents() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const wsA = ss.getSheetByName('ATENCIONS');
    const wsU = ss.getSheetByName('USUARIS');
    const wsAs = ss.getSheetByName('ASSESSORAMENTS');
    
    const aData = wsA.getDataRange().getValues();
    const aH = aData[0];
    const uData = wsU.getDataRange().getValues();
    const uH = uData[0];
    const asData = wsAs.getDataRange().getValues();
    const asH = asData[0];
    
    const dniIdxU = uH.indexOf('DNI');
    const nomIdxU = uH.indexOf('NOM');
    const fechaIdx = aH.indexOf('FECHA');
    const dniIdxA = aH.indexOf('DNI');
    const quiIdx = aH.indexOf('QUI_REP');
    const calAssIdx = aH.indexOf('CAL_ASSESSORAMENT');
    const idAtIdxAs = asH.indexOf('ID_ATENCIO');
    const assessIdx = asH.indexOf('ASSESSORAMENT');
    
    // Filtrar atenciones pendientes
    const pendientes = aData.slice(1)
      .filter(r => r[calAssIdx] === 'Pendent' || r[calAssIdx] === 'Si')
      .slice(-50)
      .reverse();
    
    const rows = pendientes.map(r => {
      const idAtencio = r[0];
      const dni = r[dniIdxA];
      const uRow = uData.find((u, i) => i > 0 && u[dniIdxU] === dni);
      
      // Obtener assessoraments de esta atencion
      const assessoraments = asData.slice(1)
        .filter(a => a[idAtIdxAs] === idAtencio)
        .map(a => a[assessIdx]);
      
      // Determinar tipo de gestion basado en assessoraments
      let tipus = 'altres';
      if (assessoraments.some(a => a && a.toLowerCase().includes('llum'))) tipus = 'llum';
      else if (assessoraments.some(a => a && a.toLowerCase().includes('gas'))) tipus = 'gas';
      else if (assessoraments.some(a => a && a.toLowerCase().includes('aigua'))) tipus = 'aigua';
      
      return {
        id: idAtencio,
        dni: dni || '-',
        nom: uRow ? uRow[nomIdxU] : '-',
        fecha: formatDate_(r[fechaIdx]),
        tecnic: r[quiIdx] || '-',
        gestio: assessoraments.slice(0, 3).join(', ') || 'Sense especificar',
        estat: r[calAssIdx] === 'Pendent' ? 'pendent' : 'en-curs',
        tipus: tipus
      };
    });
    
    return { ok: true, rows };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  7. ACTUALIZAR ESTADO DE GESTION
// ─────────────────────────────────────────────────────────────
function updateGestioEstat(idAtencio, estat) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName('ATENCIONS');
    const data = ws.getDataRange().getValues();
    const headers = data[0];
    const calAssIdx = headers.indexOf('CAL_ASSESSORAMENT');
    
    // Buscar la fila
    const filaIdx = data.findIndex((r, i) => i > 0 && r[0] == idAtencio);
    if (filaIdx === -1) {
      return { ok: false, error: 'Atencio no trobada' };
    }
    
    // Actualizar el estado
    const newValue = estat === 'completat' ? 'No' : (estat === 'en-curs' ? 'Si' : 'Pendent');
    ws.getRange(filaIdx + 1, calAssIdx + 1).setValue(newValue);
    
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  8. ULTIMAS ATENCIONES (para listado general)
//  Obtiene las ultimas N atenciones para mostrar en el listado
// ─────────────────────────────────────────────────────────────
function getUltimesAtencions(limit, dateFrom, dateTo) {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const wsA  = ss.getSheetByName('ATENCIONS');
    const wsU  = ss.getSheetByName('USUARIS');
    const wsAs = ss.getSheetByName('ASSESSORAMENTS');

    const aData = wsA.getDataRange().getValues();
    const aH    = aData[0];
    const uData = wsU.getDataRange().getValues();
    const uH    = uData[0];
    const asData = wsAs.getDataRange().getValues();
    const asH    = asData[0];
    
    // Indices de columnas
    const dniIdxU = uH.indexOf('DNI');
    const nomIdxU = uH.indexOf('NOM');
    const fechaIdx = aH.indexOf('FECHA');
    const dniIdxA  = aH.indexOf('DNI');
    const derivIdx = aH.indexOf('DERIVACIÓ');
    const quiIdx   = aH.indexOf('QUI_REP');
    const calIdx   = aH.indexOf('CAL_ASSESSORAMENT');
    const idAtIdxAs = asH.indexOf('ID_ATENCIO');
    const assessIdx = asH.indexOf('ASSESSORAMENT');

    // Filtrar por fecha si se proporcionan
    let filteredData = aData.slice(1);
    if (dateFrom) {
      const fromDate = new Date(dateFrom);
      filteredData = filteredData.filter(r => {
        const d = new Date(r[fechaIdx]);
        return !isNaN(d) && d >= fromDate;
      });
    }
    if (dateTo) {
      const toDate = new Date(dateTo);
      toDate.setHours(23, 59, 59);
      filteredData = filteredData.filter(r => {
        const d = new Date(r[fechaIdx]);
        return !isNaN(d) && d <= toDate;
      });
    }

    // Obtener las ultimas N atenciones con datos del usuario
    const rows = filteredData.slice(-(limit || 30)).reverse().map(r => {
      const idAtencio = r[0];
      const uRow = uData.find((u, i) => i > 0 && u[dniIdxU] === r[dniIdxA]);
      const d = r[fechaIdx];
      
      // Obtener assessoraments de esta atencion
      const assessoraments = asData.slice(1)
        .filter(a => a[idAtIdxAs] === idAtencio)
        .map(a => a[assessIdx]);
      
      return {
        id:        idAtencio,
        dni:       r[dniIdxA] || '-',
        nom:       uRow ? uRow[nomIdxU] : (r[aH.indexOf('NOM')] || '-'),
        fecha:     formatDate_(d),
        derivacio: r[derivIdx] || '-',
        quiRep:    r[quiIdx]   || '-',
        calAss:    r[calIdx]   || '-',
        assessoraments: assessoraments
      };
    });

    return { ok: true, rows };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  9. EXPORTAR DATOS
//  Crea una copia de los datos para exportar
// ─────────────────────────────────────────────────────────────
function exportarDades() {
  try {
    // En una implementacion real, esto podria crear un CSV o PDF
    // Por ahora, simplemente confirmamos que se puede acceder a los datos
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const url = ss.getUrl();
    
    return { ok: true, url: url };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}
