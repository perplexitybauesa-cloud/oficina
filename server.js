const express = require('express');
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json());
app.use(express.static(__dirname));

const SPREADSHEET_ID = '13-HCGphpkB8zi4m6299YdMXEADWtXLc0Hy1CJot_Cj0';
const PORT = 5000;

function getAuth() {
  let credentials;
  const credentialsJson = process.env.GOOGLE_CREDENTIALS_JSON;
  if (credentialsJson) {
    try {
      credentials = JSON.parse(credentialsJson);
    } catch(e) {
      // Env var may be corrupted, fall back to file
    }
  }
  if (!credentials) {
    const credFile = path.join(__dirname, 'google-credentials.json');
    if (fs.existsSync(credFile)) {
      credentials = JSON.parse(fs.readFileSync(credFile, 'utf8'));
    } else {
      throw new Error('Google credentials not configured. Set GOOGLE_CREDENTIALS_JSON or provide google-credentials.json');
    }
  }
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth;
}

async function getSheets() {
  const auth = getAuth();
  return google.sheets({ version: 'v4', auth });
}

function formatDate(dateStr) {
  if (!dateStr) return '-';
  const d = new Date(dateStr);
  if (isNaN(d)) return String(dateStr);
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

function sheetToObjects(values) {
  if (!values || values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
    return obj;
  });
}

async function getSheetValues(sheets, sheetName) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: sheetName,
    });
    return res.data.values || [];
  } catch (e) {
    return [];
  }
}

async function appendRow(sheets, sheetName, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: sheetName,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [values] },
  });
}

async function updateRow(sheets, sheetName, rowIndex, values) {
  const range = `${sheetName}!A${rowIndex}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [values] },
  });
}

async function updateCell(sheets, sheetName, rowIndex, colIndex, value) {
  const col = String.fromCharCode(65 + colIndex);
  const range = `${sheetName}!${col}${rowIndex}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[value]] },
  });
}

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/api/dashboard', async (req, res) => {
  try {
    const sheets = await getSheets();
    const [uValues, aValues] = await Promise.all([
      getSheetValues(sheets, 'USUARIS'),
      getSheetValues(sheets, 'ATENCIONS'),
    ]);

    const totalUsuaris = Math.max((uValues.length || 1) - 1, 0);
    const totalAtencions = Math.max((aValues.length || 1) - 1, 0);

    const aH = aValues[0] || [];
    const fechaIdx = aH.indexOf('FECHA');
    const quiRepIdx = aH.indexOf('QUI_REP');
    const derivacioIdx = aH.indexOf('DERIVACIÓ');
    const dniIdxA = aH.indexOf('DNI');
    const calAssIdx = aH.indexOf('CAL_ASSESSORAMENT');

    const ara = new Date();
    const mesActual = ara.getMonth();
    const anyActual = ara.getFullYear();
    const mesAnterior = mesActual === 0 ? 11 : mesActual - 1;
    const anyAnterior = mesActual === 0 ? anyActual - 1 : anyActual;

    const rows = aValues.slice(1);
    let atencionesMes = 0;
    let atencionesMesAnterior = 0;
    const perTecnic = {};
    const perDerivacio = {};
    let pendingGestions = 0;

    const MONTH_NAMES = ['Gen','Feb','Mar','Abr','Mai','Jun','Jul','Ago','Set','Oct','Nov','Des'];
    const perMes = {};
    for (let i = 5; i >= 0; i--) {
      const d = new Date(anyActual, mesActual - i, 1);
      const key = `${d.getFullYear()}-${d.getMonth()}`;
      perMes[key] = { label: MONTH_NAMES[d.getMonth()], count: 0, isCurrent: i === 0 };
    }

    rows.forEach(r => {
      const d = new Date(r[fechaIdx]);
      if (!isNaN(d)) {
        if (d.getMonth() === mesActual && d.getFullYear() === anyActual) atencionesMes++;
        if (d.getMonth() === mesAnterior && d.getFullYear() === anyAnterior) atencionesMesAnterior++;
        const key = `${d.getFullYear()}-${d.getMonth()}`;
        if (perMes[key]) perMes[key].count++;
      }
      const t = r[quiRepIdx] || 'Altres';
      const dv = r[derivacioIdx] || 'Altres';
      perTecnic[t] = (perTecnic[t] || 0) + 1;
      perDerivacio[dv] = (perDerivacio[dv] || 0) + 1;
      if (r[calAssIdx] === 'Pendent' || r[calAssIdx] === 'Si') pendingGestions++;
    });

    const tendenciaMensual = Object.values(perMes);

    const uH = uValues[0] || [];
    const dniIdxU = uH.indexOf('DNI');
    const nomIdxU = uH.indexOf('NOM');

    const ultimes = rows.slice(-8).reverse().map(r => {
      const uRow = uValues.find((u, i) => i > 0 && u[dniIdxU] === r[dniIdxA]);
      return {
        id: r[0],
        dni: r[dniIdxA] || '-',
        nom: uRow ? uRow[nomIdxU] : '-',
        fecha: formatDate(r[fechaIdx]),
        derivacio: r[derivacioIdx] || '-',
        quiRep: r[quiRepIdx] || '-'
      };
    });

    res.json({ ok: true, totalUsuaris, totalAtencions, atencionesMes, atencionesMesAnterior, pendingGestions, perTecnic, perDerivacio, ultimes, tendenciaMensual });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

app.get('/api/user/:dni', async (req, res) => {
  try {
    const sheets = await getSheets();
    const dni = req.params.dni.toUpperCase().trim();

    const [uValues, aValues, asValues, nValues] = await Promise.all([
      getSheetValues(sheets, 'USUARIS'),
      getSheetValues(sheets, 'ATENCIONS'),
      getSheetValues(sheets, 'ASSESSORAMENTS'),
      getSheetValues(sheets, 'NOTES'),
    ]);

    const uH = uValues[0] || [];
    const dniIdx = uH.indexOf('DNI');
    const filaIdx = uValues.findIndex((r, i) => i > 0 && r[dniIdx] === dni);
    if (filaIdx === -1) return res.json({ trobat: false });

    const obj = {};
    uH.forEach((h, i) => {
      const v = uValues[filaIdx][i];
      obj[h] = v instanceof Date ? formatDate(v) : (v !== undefined ? v : '');
    });

    const aH = aValues[0] || [];
    const dniIdxA = aH.indexOf('DNI');
    const fechaIdx = aH.indexOf('FECHA');
    const derivIdx = aH.indexOf('DERIVACIÓ');
    const quiIdx = aH.indexOf('QUI_REP');
    const obsIdx = aH.indexOf('OBSERVACIONS');

    const asH = asValues[0] || [];
    const dniIdxAs = asH.indexOf('DNI');
    const idAtIdxAs = asH.indexOf('ID_ATENCIO');
    const assessIdx = asH.indexOf('ASSESSORAMENT');

    const historial = aValues.slice(1)
      .filter(r => r[dniIdxA] === dni)
      .map(r => {
        const idAtencio = r[0];
        const assessoraments = asValues.slice(1)
          .filter(a => a[idAtIdxAs] == idAtencio)
          .map(a => a[assessIdx]);
        return {
          id: idAtencio,
          fecha: formatDate(r[fechaIdx]),
          derivacio: r[derivIdx] || '-',
          quiRep: r[quiIdx] || '-',
          observ: r[obsIdx] || '-',
          assessoraments
        };
      })
      .reverse()
      .slice(0, 15);

    let notes = [];
    if (nValues && nValues.length > 1) {
      const nH = nValues[0];
      const dniIdxN = nH.indexOf('DNI');
      const textIdx = nH.indexOf('TEXT');
      const tipusIdx = nH.indexOf('TIPUS');
      const fechaIdxN = nH.indexOf('FECHA');
      const autorIdx = nH.indexOf('AUTOR');
      notes = nValues.slice(1)
        .filter(r => r[dniIdxN] === dni)
        .map(r => ({
          text: r[textIdx],
          tipus: r[tipusIdx] || 'info',
          fecha: formatDate(r[fechaIdxN]),
          autor: r[autorIdx] || '-'
        }))
        .reverse()
        .slice(0, 5);
    }

    res.json({ trobat: true, usuari: obj, historial, totalAtencions: historial.length, notes });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

app.get('/api/users', async (req, res) => {
  try {
    const sheets = await getSheets();
    const uValues = await getSheetValues(sheets, 'USUARIS');
    const uH = uValues[0] || [];
    const dniIdx = uH.indexOf('DNI');
    const nomIdx = uH.indexOf('NOM');
    const users = uValues.slice(1).map(r => ({ DNI: r[dniIdx], NOM: r[nomIdx] }));
    res.json({ ok: true, users });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

app.post('/api/atencio', async (req, res) => {
  try {
    const sheets = await getSheets();
    const fd = req.body;
    const dni = (fd.dni || '').toUpperCase().trim();
    if (!dni) return res.json({ ok: false, error: 'DNI obligatori' });

    const uValues = await getSheetValues(sheets, 'USUARIS');
    const uH = uValues[0] || [];
    const dniIdx = uH.indexOf('DNI');
    const existingRowIdx = uValues.findIndex((r, i) => i > 0 && r[dniIdx] === dni);
    const existingRow = existingRowIdx > -1 ? uValues[existingRowIdx] : null;
    const nom = fd.nom || (existingRow ? existingRow[uH.indexOf('NOM')] : '') || '';
    const enlace = existingRow ? (existingRow[uH.indexOf('ENLACE')] || '') : '';

    await guardarUsuari(sheets, fd, dni, nom, enlace, uValues, uH);
    const idAtencio = await guardarAtencio(sheets, fd, dni, nom);

    if (fd.assessoraments && fd.assessoraments.length)
      await guardarAssessoraments(sheets, fd.assessoraments, dni, nom, enlace, idAtencio);

    await guardarGestions(sheets, fd.assessoraments || [], dni, nom, enlace, idAtencio);

    if (fd.notaInterna && fd.notaInterna.trim()) {
      await guardarNotaInterna(sheets, dni, fd.notaInterna, fd.notaAlerta ? 'alert' : 'info', fd.quiRep);
    }

    res.json({ ok: true, idAtencio, missatge: `Atencio #${idAtencio} registrada correctament` });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

async function guardarUsuari(sheets, fd, dni, nom, enlace, uValues, uH) {
  const dniIdx = uH.indexOf('DNI');
  const filaIdx = uValues.findIndex((r, i) => i > 0 && r[dniIdx] === dni);
  const ws = await getSheetValues(sheets, 'USUARIS');

  const vals = uH.map(h => {
    switch (h) {
      case 'ID_USUARI': return filaIdx === -1 ? ws.length : uValues[filaIdx][0];
      case 'DNI': return dni;
      case 'NOM': return nom;
      case 'ENLACE': return enlace || (filaIdx > -1 ? uValues[filaIdx][uH.indexOf('ENLACE')] : '') || '';
      case 'GÈNERE': return fd.genere || '';
      case 'LLOC DE NAIXEMENT': return fd.llocNaix || '';
      case 'BARRI': return fd.barri || '';
      case 'TINENÇA': return fd.tinenca || '';
      case 'ADREÇA': return fd.adreca || '';
      case 'TELF. 1': return fd.telf || '';
      case 'E-MAIL': return fd.email || '';
      case 'TOTAL PERSONES HABITATGE': return fd.persones || '';
      default: return filaIdx > -1 ? (uValues[filaIdx][uH.indexOf(h)] || '') : '';
    }
  });

  if (filaIdx === -1) {
    await appendRow(sheets, 'USUARIS', vals);
  } else {
    await updateRow(sheets, 'USUARIS', filaIdx + 1, vals);
  }
}

async function guardarAtencio(sheets, fd, dni, nom) {
  const aValues = await getSheetValues(sheets, 'ATENCIONS');
  const headers = aValues[0] || [];
  const id = aValues.length;
  const circs = Array.isArray(fd.circumstancies) ? fd.circumstancies.join(', ') : (fd.circumstancies || '');
  const now = new Date().toISOString();

  const vals = headers.map(h => {
    switch (h) {
      case 'ID_ATENCIO': return id;
      case 'DNI': return dni;
      case 'NOM': return nom;
      case 'ENLACE': return '';
      case 'FECHA': return now;
      case 'DERIVACIÓ': return fd.derivacio || '';
      case 'QUI_REP': return fd.quiRep || '';
      case 'OBSERVACIONS': return fd.observacions || '';
      case 'REFERENT': return fd.referent || '';
      case 'CIRCUMSTÀNCIES_ESPECIALS': return circs;
      case 'CAL_ASSESSORAMENT': return fd.calAssessorament || '';
      default: return '';
    }
  });

  await appendRow(sheets, 'ATENCIONS', vals);
  return id;
}

async function guardarAssessoraments(sheets, assessos, dni, nom, enlace, idAtencio) {
  const asValues = await getSheetValues(sheets, 'ASSESSORAMENTS');
  const headers = asValues[0] || [];
  const now = new Date().toISOString();

  for (let i = 0; i < assessos.length; i++) {
    const a = assessos[i];
    if (!a) continue;
    const id = asValues.length + i;
    const vals = headers.map(h => {
      switch (h) {
        case 'ID_ASSESSORAMENT': return id;
        case 'ID_ATENCIO': return idAtencio;
        case 'DNI': return dni;
        case 'NOM': return nom;
        case 'ENLACE': return enlace || '';
        case 'FECHA': return now;
        case 'ORDRE': return i + 1;
        case 'ASSESSORAMENT': return a;
        default: return '';
      }
    });
    await appendRow(sheets, 'ASSESSORAMENTS', vals);
  }
}

async function guardarGestions(sheets, assessos, dni, nom, enlace, idAtencio) {
  const now = new Date().toISOString();

  const LLUM_MAP = {
    'Canvi titularitat llum': 'CANVI_TITULARITAT',
    'Canvi comercialitzadora llum': 'CANVI_COMERCIALITZADORA',
    'Mercat lliure a mercat regulat': 'DATA_ML_MR',
    'Potencia reduida': 'POTÈNCIA_REDUÏDA',
    'Eliminacio serveis extres llum': 'ELIMINACIÓ_SERVEIS_EXTRES',
    'Bo social': 'BO_SOCIAL_DATA',
    'Nou bo social': 'NOU_BO_SOCIAL_DATA',
    'Tall efectiu llum': 'TALL_EFECTIU_LLUM',
    'Bo termic': 'BO_TERMIC',
    'Punxada llum': 'PUNXADA_LLUM',
    'Desconnexio llum': 'DESCONNEXIÓ',
    'Negociar deute llum': 'NEGOCIAR_DEUTE_LLUM',
    'Alta llum': 'ALTA_LLUM',
    'Treva llum': 'TREVA_LLUM'
  };
  const GAS_MAP = {
    'Canvi titularitat gas': 'DATA_CANVI_TIT_GAS',
    'Canvi comercialitzadora gas': 'CANVI_COMERCIALITZADORA',
    'Mercat lliure a TUR': 'ML_A_TUR',
    'Eliminacio serveis extres gas': 'ELIMINACIÓ_SERVEIS_EXTRES',
    'Tall efectiu gas': 'TALL_EFECTIU_GAS',
    'Punxada gas': 'PUNXADA_GAS',
    'Desconnexio gas': 'DESCONNEXIÓ',
    'Gestio de trams': 'GESTIÓ_TRAMS',
    'Canon social': 'CANON_SOCIAL',
    'Treva gas': 'TREVA_GAS'
  };
  const AIGUA_MAP = {
    'Canvi titularitat aigua': 'DATA_CANVI_TIT_AIGUA',
    'Ajuda MUSA (comptadors)': 'AJUDA_MUSA',
    'Tall efectiu aigua': 'TALL_EFECTIU_AIGUA',
    'Punxada aigua': 'PUNXADA_AIGUA',
    'Desconnexio aigua': 'DESCONNEXIÓ',
    'Alta aigua': 'ALTA_AIGUA',
    'Negociar deute aigua': 'NEGOCIAR_DEUTE',
    'IRER enviat': 'IRER_ENVIAT',
    'Treva aigua': 'TREVA_AIGUA'
  };
  const ALTRES_MAP = {
    'Projecte Porta a Porta': 'PROJECTE_PORTA_A_PORTA',
    'ONA': 'ONA'
  };

  const maps = [
    [LLUM_MAP, 'GESTIONS_LLUM'],
    [GAS_MAP, 'GESTIONS_GAS'],
    [AIGUA_MAP, 'GESTIONS_AIGUA'],
    [ALTRES_MAP, 'ALTRES_GESTIONS'],
  ];

  for (const [map, sheetName] of maps) {
    const matched = assessos.filter(a => map[a]);
    if (!matched.length) continue;
    try {
      const gsValues = await getSheetValues(sheets, sheetName);
      const headers = gsValues[0] || [];
      const vals = headers.map(h => {
        switch (h) {
          case 'ID_ATENCIO': return idAtencio;
          case 'DNI': return dni;
          case 'NOM': return nom;
          case 'ENLACE': return enlace || '';
          case 'FECHA': return now;
          default: {
            const assName = Object.keys(map).find(a => map[a] === h);
            return assName && matched.includes(assName) ? now : '';
          }
        }
      });
      await appendRow(sheets, sheetName, vals);
    } catch (e) {
      console.warn(`Could not write to ${sheetName}:`, e.message);
    }
  }
}

async function guardarNotaInterna(sheets, dni, text, tipus, autor) {
  const nValues = await getSheetValues(sheets, 'NOTES');
  const id = nValues.length || 1;
  const now = new Date().toISOString();
  if (!nValues || nValues.length === 0) {
    await appendRow(sheets, 'NOTES', ['ID_NOTA', 'DNI', 'TEXT', 'TIPUS', 'FECHA', 'AUTOR']);
  }
  await appendRow(sheets, 'NOTES', [id, dni, text, tipus, now, autor || '']);
}

app.post('/api/nota', async (req, res) => {
  try {
    const sheets = await getSheets();
    const data = req.body;
    const dni = (data.dni || '').toUpperCase().trim();
    if (!dni || !data.text) return res.json({ ok: false, error: 'Falten dades' });
    await guardarNotaInterna(sheets, dni, data.text, data.tipus || 'info', data.autor || '');
    res.json({ ok: true });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

app.get('/api/gestions', async (req, res) => {
  try {
    const sheets = await getSheets();
    const [aValues, uValues, asValues] = await Promise.all([
      getSheetValues(sheets, 'ATENCIONS'),
      getSheetValues(sheets, 'USUARIS'),
      getSheetValues(sheets, 'ASSESSORAMENTS'),
    ]);

    const aH = aValues[0] || [];
    const uH = uValues[0] || [];
    const asH = asValues[0] || [];

    const dniIdxU = uH.indexOf('DNI');
    const nomIdxU = uH.indexOf('NOM');
    const fechaIdx = aH.indexOf('FECHA');
    const dniIdxA = aH.indexOf('DNI');
    const quiIdx = aH.indexOf('QUI_REP');
    const calAssIdx = aH.indexOf('CAL_ASSESSORAMENT');
    const idAtIdxAs = asH.indexOf('ID_ATENCIO');
    const assessIdx = asH.indexOf('ASSESSORAMENT');

    const pendientes = aValues.slice(1)
      .filter(r => r[calAssIdx] === 'Pendent' || r[calAssIdx] === 'Si')
      .slice(-50)
      .reverse();

    const rows = pendientes.map(r => {
      const idAtencio = r[0];
      const dni = r[dniIdxA];
      const uRow = uValues.find((u, i) => i > 0 && u[dniIdxU] === dni);
      const assessoraments = asValues.slice(1)
        .filter(a => a[idAtIdxAs] == idAtencio)
        .map(a => a[assessIdx]);

      let tipus = 'altres';
      if (assessoraments.some(a => a && a.toLowerCase().includes('llum'))) tipus = 'llum';
      else if (assessoraments.some(a => a && a.toLowerCase().includes('gas'))) tipus = 'gas';
      else if (assessoraments.some(a => a && a.toLowerCase().includes('aigua'))) tipus = 'aigua';

      return {
        id: idAtencio,
        dni: dni || '-',
        nom: uRow ? uRow[nomIdxU] : '-',
        fecha: formatDate(r[fechaIdx]),
        tecnic: r[quiIdx] || '-',
        gestio: assessoraments.slice(0, 3).join(', ') || 'Sense especificar',
        estat: r[calAssIdx] === 'Pendent' ? 'pendent' : 'en-curs',
        tipus
      };
    });

    res.json({ ok: true, rows });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

app.post('/api/gestions/:id/estat', async (req, res) => {
  try {
    const sheets = await getSheets();
    const idAtencio = req.params.id;
    const { estat } = req.body;

    const aValues = await getSheetValues(sheets, 'ATENCIONS');
    const headers = aValues[0] || [];
    const calAssIdx = headers.indexOf('CAL_ASSESSORAMENT');

    const filaIdx = aValues.findIndex((r, i) => i > 0 && r[0] == idAtencio);
    if (filaIdx === -1) return res.json({ ok: false, error: 'Atencio no trobada' });

    const newValue = estat === 'completat' ? 'No' : (estat === 'en-curs' ? 'Si' : 'Pendent');
    await updateCell(sheets, 'ATENCIONS', filaIdx + 1, calAssIdx, newValue);

    res.json({ ok: true });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

app.get('/api/registres', async (req, res) => {
  try {
    const sheets = await getSheets();
    const { limit, dateFrom, dateTo } = req.query;

    const [aValues, uValues, asValues] = await Promise.all([
      getSheetValues(sheets, 'ATENCIONS'),
      getSheetValues(sheets, 'USUARIS'),
      getSheetValues(sheets, 'ASSESSORAMENTS'),
    ]);

    const aH = aValues[0] || [];
    const uH = uValues[0] || [];
    const asH = asValues[0] || [];

    const dniIdxU = uH.indexOf('DNI');
    const nomIdxU = uH.indexOf('NOM');
    const fechaIdx = aH.indexOf('FECHA');
    const dniIdxA = aH.indexOf('DNI');
    const derivIdx = aH.indexOf('DERIVACIÓ');
    const quiIdx = aH.indexOf('QUI_REP');
    const calIdx = aH.indexOf('CAL_ASSESSORAMENT');
    const idAtIdxAs = asH.indexOf('ID_ATENCIO');
    const assessIdx = asH.indexOf('ASSESSORAMENT');

    let filteredData = aValues.slice(1);
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

    const rows = filteredData.slice(-(parseInt(limit) || 30)).reverse().map(r => {
      const idAtencio = r[0];
      const uRow = uValues.find((u, i) => i > 0 && u[dniIdxU] === r[dniIdxA]);
      const assessoraments = asValues.slice(1)
        .filter(a => a[idAtIdxAs] == idAtencio)
        .map(a => a[assessIdx]);
      return {
        id: idAtencio,
        dni: r[dniIdxA] || '-',
        nom: uRow ? uRow[nomIdxU] : (r[aH.indexOf('NOM')] || '-'),
        fecha: formatDate(r[fechaIdx]),
        derivacio: r[derivIdx] || '-',
        quiRep: r[quiIdx] || '-',
        calAss: r[calIdx] || '-',
        assessoraments
      };
    });

    res.json({ ok: true, rows });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

app.get('/api/export', async (req, res) => {
  res.json({ ok: true, url: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}` });
});

app.get('/api/health', (req, res) => {
  res.json({ ok: true, status: 'running' });
});

app.get('/download/Code.gs', (req, res) => {
  res.setHeader('Content-Type', 'text/plain; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename="Code.gs"');
  res.sendFile(path.join(__dirname, 'Code.gs'));
});

app.get('/download/Index.html', (req, res) => {
  res.setHeader('Content-Type', 'text/plain; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename="Index.html"');
  res.sendFile(path.join(__dirname, 'Index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on http://0.0.0.0:${PORT}`);
});
