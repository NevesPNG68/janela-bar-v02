import fs from 'node:fs';
import path from 'node:path';
import * as XLSX from 'xlsx';

const TOKEN = process.env.DROPBOX_TOKEN;
const CONFIG_PATH = (process.env.DROPBOX_FILE_PATH || '').trim();
const SEARCH_NAME = (process.env.DROPBOX_SEARCH_NAME || 'Por venda_2026.xlsm').trim();
const OUT_FILE = 'dados/dashboard-data.json';

if (!TOKEN) {
  console.error('❌ Secret DROPBOX_TOKEN não encontrado.');
  process.exit(1);
}

function norm(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '');
}

function n(v) {
  const x = Number(String(v ?? '').replace(',', '.'));
  return Number.isFinite(x) ? x : 0;
}

function txt(v) {
  return String(v ?? '').trim();
}

async function dropboxApi(endpoint, body) {
  const resp = await fetch(`https://api.dropboxapi.com/2/${endpoint}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  });
  const raw = await resp.text();
  let data;
  try { data = raw ? JSON.parse(raw) : {}; } catch { data = { raw }; }
  if (!resp.ok) {
    const err = new Error(`Dropbox API ${endpoint} falhou: ${JSON.stringify(data)}`);
    err.data = data;
    throw err;
  }
  return data;
}

function extractPathFromSearchMatch(match) {
  const m = match?.metadata?.metadata || match?.metadata || match;
  return m?.path_display || m?.path_lower || '';
}

async function searchPathByName(fileName) {
  const data = await dropboxApi('files/search_v2', {
    query: fileName,
    options: { filename_only: true, max_results: 20 }
  });
  const paths = [...new Set((data.matches || []).map(extractPathFromSearchMatch).filter(Boolean))];
  if (paths.length) return paths[0];

  let page = await dropboxApi('files/list_folder', {
    path: '',
    recursive: true,
    include_deleted: false,
    include_non_downloadable_files: false
  });
  while (true) {
    for (const entry of page.entries || []) {
      if (entry['.tag'] === 'file' && entry.name === fileName) {
        return entry.path_display || entry.path_lower;
      }
    }
    if (!page.has_more) break;
    page = await dropboxApi('files/list_folder/continue', { cursor: page.cursor });
  }
  return '';
}

async function resolveDropboxPath() {
  if (CONFIG_PATH) {
    try {
      const meta = await dropboxApi('files/get_metadata', { path: CONFIG_PATH });
      if (meta?.['.tag'] === 'file') return meta.path_display || meta.path_lower || CONFIG_PATH;
    } catch (e) {
      console.warn('⚠️ DROPBOX_FILE_PATH não funcionou. Tentando localizar pelo nome do arquivo...');
    }
  }
  const found = await searchPathByName(SEARCH_NAME);
  if (!found) throw new Error(`Arquivo não encontrado no Dropbox: ${SEARCH_NAME}`);
  return found;
}

async function downloadDropboxFile(dropboxPath) {
  const resp = await fetch('https://content.dropboxapi.com/2/files/download', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      'Dropbox-API-Arg': JSON.stringify({ path: dropboxPath })
    }
  });
  if (!resp.ok) {
    const raw = await resp.text();
    throw new Error(`Falha ao baixar planilha do Dropbox: ${resp.status} ${raw}`);
  }
  const ab = await resp.arrayBuffer();
  return Buffer.from(ab);
}

function sheetRows(wb, sheetName) {
  if (!sheetName) return [];
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });
}

function parseAllData(wb) {
  const sn = wb.SheetNames.find(name => norm(name) === 'vendasoperador') || wb.SheetNames.find(name => norm(name).includes('venda'));
  if (!sn) throw new Error(`Aba VendasOperador não encontrada. Abas: ${wb.SheetNames.join(', ')}`);
  const raw = sheetRows(wb, sn);
  const rows = raw.slice(1).map(r => ({
    tipo: txt(r[6]),
    produto: txt(r[8]),
    qtd: n(r[9]),
    total: n(r[12]),
    pgto: txt(r[13]),
    grupo: txt(r[16]),
    dia: txt(r[17]),
    hora: txt(r[19]),
    mes: txt(r[20]),
    liquido: n(r[21]),
    qtven: n(r[22]),
    semana: n(r[18]),
    qt_double: n(r[25]),
    qt_final: n(r[26]),
    quinzena: txt(r[24])
  })).filter(r => r.liquido !== 0);
  return { sheet: sn, rows };
}

function parseDespesas(wb) {
  const sn = wb.SheetNames.find(name => norm(name) === 'despesas' || norm(name).includes('despesa'));
  const raw = sheetRows(wb, sn);
  return raw.slice(1).map(r => ({
    mes: txt(r[2]),
    ano: txt(r[3]),
    qtd: n(r[4]),
    item: txt(r[5]),
    fornecedor: txt(r[7]),
    total: n(r[13])
  })).filter(r => r.item);
}

function parseMatriz(wb) {
  const sn = wb.SheetNames.find(name => norm(name).includes('matrizanalitico') || norm(name).includes('matrizanaltico'));
  const raw = sheetRows(wb, sn);
  const rows = [];
  for (let i = 7; i < raw.length; i++) {
    const r = raw[i] || [];
    const mes = txt(r[1]);
    if (!mes) continue;
    const item = {
      mes,
      receita: n(r[7]),
      custoFixo: n(r[30]),
      custoFixoVar: n(r[31]),
      totalFixo: n(r[32]),
      custoGeral: n(r[35]),
      cmv: n(r[36]),
      mkp: n(r[37]),
      margem: n(r[38]),
      pctCmv: n(r[42]),
      lucPrej: n(r[43]),
      pctLuc: n(r[44])
    };
    if (item.receita > 0 || item.custoFixo > 0 || item.custoGeral > 0 || item.cmv > 0) rows.push(item);
  }
  return rows;
}

function parseDrinkAnalysis(wb) {
  const sn = wb.SheetNames.find(name => norm(name).includes('analisedrink') || norm(name).includes('anlisedrink'));
  const raw = sheetRows(wb, sn);
  const rows = [];
  for (let i = 5; i < raw.length; i++) {
    const r = raw[i] || [];
    const mes = txt(r[1] || r[27]);
    if (!mes) continue;
    const p = mes.split('/');
    const yy = parseInt(p[1] || '0', 10);
    const ano = yy ? String(yy + (yy < 100 ? 2000 : 0)) : '';
    rows.push({
      mes, ano,
      drinks: [n(r[2]), n(r[3]), n(r[4]), n(r[5]), n(r[6])], drinksTotal: n(r[7]),
      especiais: [n(r[8]), n(r[9]), n(r[10]), n(r[11]), n(r[12])], especiaisTotal: n(r[13]),
      semAlcool: [n(r[14]), n(r[15]), n(r[16]), n(r[17]), n(r[18])], semAlcoolTotal: n(r[19]),
      double: [n(r[20]), n(r[21]), n(r[22]), n(r[23]), n(r[24])], doubleTotal: n(r[25]),
      faturamento: n(r[28]), drinkTotal: n(r[29]), participacao: n(r[30])
    });
  }
  return rows;
}

function parseProjecao(wb) {
  const sn = wb.SheetNames.find(name => norm(name).includes('projeto') || norm(name).includes('projecao'));
  const raw = sheetRows(wb, sn);
  if (!raw.length) return null;
  const mesProj = raw[0] ? txt(raw[0][10]) : '';
  const historico = [];
  let q1Atual = 0, q2Atual = 0, totalAcum = 0;
  for (let i = 3; i < raw.length; i++) {
    const r = raw[i] || [];
    const mes = txt(r[1]);
    if (!mes) continue;
    const q1 = n(r[2]);
    const q2 = n(r[5]);
    const total = n(r[7]);
    if (mesProj && mes.toLowerCase() === mesProj.toLowerCase()) {
      q1Atual = q1;
      q2Atual = q2;
      totalAcum = total;
      continue;
    }
    if (total > 0 && q1 > 0) historico.push({ mes, q1, q2, total, fator: total / q1 });
  }
  if (!historico.length) return null;
  const fatores = historico.map(h => h.fator);
  const fatorMax = Math.max(...fatores);
  const fatorMin = Math.min(...fatores);
  const fatorMed = fatores.reduce((a, b) => a + b, 0) / fatores.length;
  return {
    mesProj, q1Atual, q2Atual, totalAcum, fatorMax, fatorMin, fatorMed,
    projMax: q1Atual > 0 ? q1Atual * fatorMax : 0,
    projMin: q1Atual > 0 ? q1Atual * fatorMin : 0,
    projMed: q1Atual > 0 ? q1Atual * fatorMed : 0,
    projMediaGeral: q1Atual > 0 ? ((q1Atual * fatorMax) + (q1Atual * fatorMed)) / 2 : 0,
    historico
  };
}

function sortMes(arr) {
  const meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  return [...arr].sort((a, b) => {
    const p = s => {
      const x = String(s || '').split('/');
      const m = x[0].slice(0, 3);
      const y0 = parseInt(x[1] || '0', 10);
      const y = y0 + (y0 < 100 ? 2000 : 0);
      return [m, y];
    };
    const [ma, ya] = p(a);
    const [mb, yb] = p(b);
    return ya !== yb ? ya - yb : meses.indexOf(ma) - meses.indexOf(mb);
  });
}

function buildSummary(allData) {
  const byMes = {};
  const byGrupo = {};
  const byPgto = {};
  const byProduto = {};
  for (const d of allData) {
    byMes[d.mes] = (byMes[d.mes] || 0) + d.liquido;
    byGrupo[d.grupo] = (byGrupo[d.grupo] || 0) + d.liquido;
    byPgto[d.pgto] = (byPgto[d.pgto] || 0) + d.liquido;
    byProduto[d.produto] = (byProduto[d.produto] || 0) + d.liquido;
  }
  const top = obj => Object.entries(obj).filter(([k]) => k).sort((a, b) => b[1] - a[1]).slice(0, 15).map(([label, value]) => ({ label, value }));
  const liquido = allData.reduce((s, d) => s + d.liquido, 0);
  const bruto = allData.reduce((s, d) => s + d.total, 0);
  const vendas = allData.reduce((s, d) => s + d.qtven, 0);
  const mesesOrdenados = sortMes(Object.keys(byMes).filter(Boolean));
  return {
    totalRegistros: allData.length,
    liquido,
    bruto,
    vendas,
    ticketMedio: vendas > 0 ? liquido / vendas : 0,
    meses: mesesOrdenados,
    periodo: { inicio: mesesOrdenados[0] || '', fim: mesesOrdenados[mesesOrdenados.length - 1] || '' },
    topGrupos: top(byGrupo),
    topPagamentos: top(byPgto),
    topProdutos: top(byProduto),
    faturamentoMensal: mesesOrdenados.map(mes => ({ mes, liquido: byMes[mes] || 0 }))
  };
}

const dropboxPath = await resolveDropboxPath();
console.log(`✅ Caminho Dropbox usado: ${dropboxPath}`);

const buffer = await downloadDropboxFile(dropboxPath);
console.log(`✅ Planilha baixada com sucesso: ${(buffer.length / 1024 / 1024).toFixed(2)} MB`);

const wb = XLSX.read(buffer, { type: 'buffer', dense: true, cellFormula: false, cellHTML: false, cellStyles: false, cellText: false });
const parsed = parseAllData(wb);
const allData = parsed.rows;

const output = {
  updatedAt: new Date().toISOString(),
  status: 'ok',
  source: {
    fileName: SEARCH_NAME,
    dropboxPath,
    vendasSheet: parsed.sheet,
    sheetsFound: wb.SheetNames
  },
  summary: buildSummary(allData),
  allData,
  matrizData: parseMatriz(wb),
  dashboardGoalsData: null,
  drinkAnalysisData: parseDrinkAnalysis(wb),
  projecaoData: parseProjecao(wb),
  dadosDespesas: parseDespesas(wb)
};

fs.mkdirSync(path.dirname(OUT_FILE), { recursive: true });
fs.writeFileSync(OUT_FILE, JSON.stringify(output, null, 2));
console.log(`✅ Dados tratados gerados em ${OUT_FILE}`);
console.log(`✅ Registros tratados: ${allData.length}`);
