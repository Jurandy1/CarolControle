import XLSX from 'xlsx';

const SHEET_NAMES = new Set([
  'PCA 2026', 'PCA 2025', '2025', 'MEMOS CAP', 'STARTGOV',
  'FISCAIS DE CONTRATO', 'SEI 2024', 'PROCESSOS CONECTA', 'MEMOS DIVERSOS'
]);

const HEADER_ROW = {
  'PCA 2026': 0,
  'PCA 2025': 0,
  '2025': 1,
  'MEMOS CAP': 1,
  'STARTGOV': 0,
  'FISCAIS DE CONTRATO': 1,
  'SEI 2024': 2,
  'PROCESSOS CONECTA': 2,
  'MEMOS DIVERSOS': 0
};

const SECAO_MARKERS = ['RELATORIO', 'RELATÓRIO', 'PROCESSOS DESPESA', 'PROCESSOS |'];

const FISCAIS_ORDEM = [
  'SERVFAZ', 'PORTARIA', 'DEDETIZAÇÃO', 'IGAS- LIMPEZA', 'AGUA E GAS',
  'EXPEDIENTE E LIMPEZA', 'MOBILIA', 'GENEROS', 'CESTA BASICA', 'FRALDA E LUVA',
  'ENXOVAL', 'QUENTINHAS', 'MEGA ON', 'MANUTENÇAO ELEVADOR', 'EXTINTOR', 'R&P'
];

export function clean(v) {
  if (v == null) return '';
  if (v instanceof Date) {
    const d = String(v.getDate()).padStart(2, '0');
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const a = v.getFullYear();
    return `${d}/${m}/${a}`;
  }
  if (typeof v === 'number') return String(v);
  return String(v).replace(/\u00A0/g, ' ').trim();
}

function isSectionBreak(v) {
  const u = v.toUpperCase();
  return SECAO_MARKERS.some(m => u.includes(m));
}

function newId() {
  return 'rec_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
}

function parseSheet(sheetName, rows) {
  const headerRow = HEADER_ROW[sheetName];
  if (headerRow == null || !rows || rows.length <= headerRow) return [];

  const result = [];
  const startRow = headerRow + 1;

  if (sheetName === 'PCA 2026') {
    for (let i = startRow; i < rows.length; i++) {
      const r = rows[i];
      const desc = clean(r[1]);
      if (!desc && !clean(r[0])) continue;
      result.push({ tab: sheetName, data: {
        contrato: clean(r[0]), descricao: desc,
        dataInicio: clean(r[2]), dataFim: clean(r[3]),
        situacao: clean(r[4]), modelo: clean(r[5]),
        proc26: clean(r[6]), proc2025: clean(r[7]),
        proc2024: clean(r[8]), repactuacao: clean(r[9]),
        aditivo: clean(r[10]), obsAditivos: clean(r[11]),
        requerimentoPag: clean(r[12]), saldoContrato: clean(r[13])
      }});
    }
  } else if (sheetName === 'PCA 2025') {
    for (let i = startRow; i < rows.length; i++) {
      const r = rows[i];
      const desc = clean(r[1]);
      if (!desc && !clean(r[0])) continue;
      result.push({ tab: sheetName, data: {
        contrato: clean(r[0]), descricao: desc,
        dataInicio: clean(r[2]), dataFim: clean(r[3]),
        situacao: clean(r[4]), modelo: clean(r[5]),
        proc26: clean(r[6]), proc2025: clean(r[7]),
        proc2024: clean(r[8]), repactuacao: clean(r[9]),
        aditivo: clean(r[10]), obsAditivos: clean(r[11]),
        vlrTotal: clean(r[12]), saldoContrato: clean(r[13])
      }});
    }
  } else if (sheetName === '2025') {
    for (let i = startRow; i < rows.length; i++) {
      const r = rows[i];
      if (!clean(r[0]) && !clean(r[1])) continue;
      result.push({ tab: sheetName, data: {
        contrato: clean(r[0]), descricao: clean(r[1]),
        dataFim: clean(r[2]), situacao: clean(r[3]),
        modelo: clean(r[4]),
        proc2025: clean(r[5]), proc2024: clean(r[6]),
        dez: clean(r[7]),
        jan: clean(r[8]), fev: clean(r[9]),
        mar: clean(r[10]), abr: clean(r[11]),
        mai: clean(r[12]), jun: clean(r[13]),
        jul: clean(r[14]), ago: clean(r[15]),
        set: clean(r[16]), out: clean(r[17]),
        nov: clean(r[18])
      }});
    }
  } else if (sheetName === 'MEMOS CAP') {
    for (let i = startRow; i < rows.length; i++) {
      const r = rows[i];
      const docCap = clean(r[0]);
      const docUnid = clean(r[4]);
      if (!docCap && !docUnid) continue;
      result.push({ tab: sheetName, data: {
        docCap, seiCap: clean(r[1]),
        situacaoCap: clean(r[2]),
        docUnidade: docUnid, seiUnidade: clean(r[5]),
        situacaoUnidade: clean(r[6])
      }});
    }
  } else if (sheetName === 'STARTGOV') {
    for (let i = startRow; i < rows.length; i++) {
      const r = rows[i];
      if (!clean(r[0]) && !clean(r[1])) continue;
      result.push({ tab: sheetName, data: {
        numero: clean(r[0]), descricao: clean(r[1]),
        valor: clean(r[2]), ano: clean(r[3])
      }});
    }
  } else if (sheetName === 'FISCAIS DE CONTRATO') {
    let lastFiscal = '';
    for (let i = startRow; i < rows.length; i++) {
      const r = rows[i];
      const serv = clean(r[0]);
      if (!serv || serv.toUpperCase() === 'FISCAL') continue;
      let fiscal = clean(r[1]);
      if (fiscal) lastFiscal = fiscal;
      else if (lastFiscal) fiscal = lastFiscal;
      result.push({ tab: sheetName, data: {
        servico: serv, fiscal, suplente: clean(r[2])
      }});
    }
  } else if (sheetName === 'SEI 2024') {
    for (let i = startRow; i < rows.length; i++) {
      const r = rows[i];
      const nome = clean(r[1]);
      if (!nome) continue;
      if (isSectionBreak(nome)) break;
      result.push({ tab: sheetName, data: {
        statusLinha: clean(r[0]), nome,
        numDoc: clean(r[2]), procSei: clean(r[3]), procConecta: clean(r[4]),
        fev: clean(r[5]), mar: clean(r[6]), abr: clean(r[7]), mai: clean(r[8]),
        jun: clean(r[9]), jul: clean(r[10]), ago: clean(r[11]), set: clean(r[12]),
        out: clean(r[13])
      }});
    }
  } else if (sheetName === 'PROCESSOS CONECTA') {
    for (let i = startRow; i < rows.length; i++) {
      const r = rows[i];
      const nome = clean(r[0]);
      if (!nome) continue;
      if (isSectionBreak(nome)) break;
      result.push({ tab: sheetName, data: {
        nome,
        numDoc: clean(r[1]), procSei: clean(r[2]), procConecta: clean(r[3]),
        fev: clean(r[4]), mar: clean(r[5]), abr: clean(r[6]), mai: clean(r[7]),
        jun: clean(r[8]), jul: clean(r[9]), ago: clean(r[10])
      }});
    }
  } else if (sheetName === 'MEMOS DIVERSOS') {
    for (let i = startRow; i < rows.length; i++) {
      const r = rows[i];
      const desc = clean(r[0]);
      if (!desc) continue;
      result.push({ tab: sheetName, data: {
        descricao: desc,
        proc1: clean(r[1]), proc2: clean(r[2]), proc3: clean(r[3]),
        proc4: clean(r[4]), proc5: clean(r[5])
      }});
    }
  }

  return result.map(item => ({
    id: newId(),
    tab: item.tab,
    data: item.data,
    linkSei: '',
    linkArquivo: '',
    linkObs: '',
    obsCoord: '',
    procAdicional: '',
    numAditivo: '',
    documentos: [],
    anotacoes: [],
    updatedAt: new Date().toISOString()
  }));
}

export function repairFiscaisContrato(records) {
  const fiscais = records.filter(r => r.tab === 'FISCAIS DE CONTRATO');
  if (!fiscais.length) return false;
  fiscais.sort((a, b) => {
    const ia = FISCAIS_ORDEM.findIndex(o => o.toUpperCase() === clean(a.data?.servico).toUpperCase());
    const ib = FISCAIS_ORDEM.findIndex(o => o.toUpperCase() === clean(b.data?.servico).toUpperCase());
    return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
  });
  let lastFiscal = '';
  let changed = false;
  for (const r of fiscais) {
    const fiscal = clean(r.data?.fiscal);
    if (fiscal) lastFiscal = fiscal;
    else if (lastFiscal && r.data) {
      r.data.fiscal = lastFiscal;
      changed = true;
    }
  }
  return changed;
}

export function parseWorkbookFromBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const all = [];
  for (const sheetName of wb.SheetNames) {
    if (!SHEET_NAMES.has(sheetName)) continue;
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
    all.push(...parseSheet(sheetName, rows));
  }
  repairFiscaisContrato(all);
  return { records: all, sheetNames: wb.SheetNames.filter(s => SHEET_NAMES.has(s)) };
}

export function parseWorkbookFromFile(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const all = [];
  for (const sheetName of wb.SheetNames) {
    if (!SHEET_NAMES.has(sheetName)) continue;
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
    all.push(...parseSheet(sheetName, rows));
  }
  repairFiscaisContrato(all);
  return { records: all, sheetNames: wb.SheetNames.filter(s => SHEET_NAMES.has(s)) };
}
