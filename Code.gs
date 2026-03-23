// ═══════════════════════════════════════════════════════════════════════════════
//  GANTT TI — Google Apps Script
//  Lê: Planejamento + Marcos + Legenda
//  Gera: Gantt (dia a dia) + Conflitos + Dashboard
// ═══════════════════════════════════════════════════════════════════════════════

const CFG = {
  sheetPlan:      'Planejamento',
  sheetMarcos:    'Marcos',
  sheetLegenda:   'Legenda',
  sheetGantt:     'Gantt',
  sheetConflitos: 'Conflitos',
  sheetDashboard: 'Dashboard',
  baseDate:       new Date(2025, 11, 1),  // 01/12/2025 = col B da aba Planejamento
  marcosDataRow:  4,                       // primeira linha de dados em Marcos
  calendarDays:   365,                     // dias exibidos no Gantt (365 = ano inteiro)
};

// Paleta automática para projetos sem cor definida na Legenda
const AUTO_PALETTE = [
  ['#674EA7','#FFFFFF'], ['#0097A7','#FFFFFF'], ['#E65100','#FFFFFF'],
  ['#2E7D32','#FFFFFF'], ['#AD1457','#FFFFFF'], ['#1565C0','#FFFFFF'],
  ['#4E342E','#FFFFFF'], ['#00695C','#FFFFFF'], ['#6A1B9A','#FFFFFF'],
  ['#37474F','#FFFFFF'], ['#558B2F','#FFFFFF'], ['#F57F17','#000000'],
];

// Cores das barras por tipo de marco (independente de projeto)
const TIPO_COLOR = {
  'Go-live':         '#C00000',
  'Start UP':        '#C00000',
  'Testes':          '#70AD47',
  'UAT':             '#70AD47',
  'Infraestrutura':  '#2E75B6',
  'Demo':            '#7030A0',
  'Documentação':    '#BF8F00',
  'Beta':            '#375623',
  'Desenvolvimento': '#2E75B6',
  'Entrega':         '#2E75B6',
  'Apresentação':    '#7030A0',
  'Manutenção':      '#595959',
};

// Restrições que geram conflito
const RESTRICTION_TYPES = new Set([
  'Férias', 'Folga', 'Atestado Médico', 'ASO', 'Compromisso',
  'Viagem de retorno', 'Viagem para Lajeado', 'Fechamento de ano',
]);

// Abreviação de restrições para caber na célula estreita
const RESTR_ABBR = {
  'Férias':               'Fér',
  'Folga':                'Flg',
  'Atestado Médico':      'Ats',
  'ASO':                  'ASO',
  'Compromisso':          'Cmp',
  'Viagem de retorno':    'Vgm',
  'Viagem para Lajeado':  'Vgm',
  'Fechamento de ano':    'Fch',
};

const MONTHS_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const DAYS_PT   = ['D','S','T','Q','Q','S','S']; // Dom=0 … Sáb=6

// ─── MENU ─────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('🚀 TI Dashboard', [
    { name: '🔄  Atualizar tudo (Normalizar + Gantt + Métricas)', functionName: 'pipelineTudo'           },
    null,
    { name: '📋  Normalizar dados',                               functionName: 'normalizarDados'        },
    { name: '📅  Gerar Gantt',                                    functionName: 'gerarGantt'             },
    { name: '📊  Gerar Métricas',                                 functionName: 'gerarMetricas'          },
    null,
    { name: '⏰  Ativar atualização semanal',                      functionName: 'ativarTriggerSemanal'   },
    { name: '✖  Remover atualização automática',                   functionName: 'removerTriggers'        },
  ]);
}

function pipelineTudo() {
  normalizarDados();
  gerarGantt();
  gerarMetricas();
}

function ativarTriggerSemanal() {
  removerTriggers();
  ScriptApp.newTrigger('gerarGantt')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  SpreadsheetApp.getUi().alert('Atualização automática ativada: toda segunda-feira às 7h.');
}
function removerTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'gerarGantt') ScriptApp.deleteTrigger(t);
  });
}

// ─── FUNÇÃO PRINCIPAL ──────────────────────────────────────────────────────────
function gerarGantt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getActiveSpreadsheet().toast('Lendo dados...', 'Gerando Gantt', -1);

  const legendaCores = lerLegenda(ss);        // { 'Farmabase': ['#E06666','#FFFFFF'], ... }
  const plano        = lerPlanejamento(ss);
  const marcos       = lerMarcos(ss);
  const projetos     = getProjetosOrdem(marcos);
  const conflitos    = detectarConflitos(marcos, plano);

  gerarAbaGantt(ss, marcos, projetos, conflitos, plano, legendaCores);
  gerarAbaConflitos(ss, conflitos);
  gerarAbaDashboard(ss, marcos, projetos, conflitos);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${marcos.length} marcos | ${projetos.length} projetos | ${conflitos.length} conflito(s)`,
    'Gantt gerado!', 6
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// 1. LÊ CORES DA ABA LEGENDA
// ═══════════════════════════════════════════════════════════════════════════════
// Formato esperado na aba Legenda:
//   Coluna B: nome do projeto com cor de fundo = cor do projeto
//   A cor de fundo da célula B define a cor do projeto no Gantt
function lerLegenda(ss) {
  const ws = ss.getSheetByName(CFG.sheetLegenda);
  if (!ws) return {};

  const lastRow = ws.getLastRow();
  if (lastRow < 1) return {};

  const projColors = {};
  // Lê valores e backgrounds da coluna B de uma só vez (batch)
  const vals = ws.getRange(1, 2, lastRow, 1).getValues();
  const bgs  = ws.getRange(1, 2, lastRow, 1).getBackgrounds();

  for (let i = 0; i < lastRow; i++) {
    const projName = vals[i][0];
    if (!projName || typeof projName !== 'string' || !projName.trim()) continue;
    const bg = bgs[i][0];
    if (!bg || bg === '#ffffff' || bg === '#000000') continue;
    const fg = contrastColor(bg);
    projColors[projName.trim()] = [bg, fg];
  }
  return projColors;
}

// Calcula cor de texto (preto ou branco) com base no brilho do fundo
function contrastColor(hex) {
  try {
    const r = parseInt(hex.slice(1,3), 16);
    const g = parseInt(hex.slice(3,5), 16);
    const b = parseInt(hex.slice(5,7), 16);
    return (r*299 + g*587 + b*114) / 1000 > 140 ? '#000000' : '#FFFFFF';
  } catch(e) { return '#FFFFFF'; }
}

// Versão clara de uma cor hex (mistura com branco)
function lightenHex(hex, factor) {
  try {
    factor = factor || 0.55;
    const r = Math.round(parseInt(hex.slice(1,3),16) + (255 - parseInt(hex.slice(1,3),16)) * factor);
    const g = Math.round(parseInt(hex.slice(3,5),16) + (255 - parseInt(hex.slice(3,5),16)) * factor);
    const b = Math.round(parseInt(hex.slice(5,7),16) + (255 - parseInt(hex.slice(5,7),16)) * factor);
    return `#${r.toString(16).padStart(2,'0')}${g.toString(16).padStart(2,'0')}${b.toString(16).padStart(2,'0')}`;
  } catch(e) { return '#F5F5F5'; }
}

// Cor do projeto: Legenda > AUTO_PALETTE
const _autoColorMap = {};
let _autoIdx = 0;
function getProjColor(proj, legendaCores) {
  if (legendaCores && legendaCores[proj]) return legendaCores[proj];
  if (_autoColorMap[proj]) return _autoColorMap[proj];
  const c = AUTO_PALETTE[_autoIdx % AUTO_PALETTE.length];
  _autoColorMap[proj] = c; _autoIdx++;
  return c;
}

// ═══════════════════════════════════════════════════════════════════════════════
// 2. LÊ PLANEJAMENTO
// ═══════════════════════════════════════════════════════════════════════════════
function lerPlanejamento(ss) {
  const ws   = ss.getSheetByName(CFG.sheetPlan);
  const data = ws.getDataRange().getValues();

  const personRowMap = {};
  for (let r = 0; r < data.length; r++) {
    const name = data[r][0];
    if (name && typeof name === 'string' && name.trim())
      personRowMap[name.trim().toUpperCase()] = r;
  }

  // Mapa dateKey→colIdx a partir da linha 0 (row 1 da sheet)
  const dateColMap = {};
  const row0 = data[0];
  for (let c = 1; c < row0.length; c++) {
    if (row0[c] instanceof Date) dateColMap[dateKey(row0[c])] = c;
  }

  return { data, personRowMap, dateColMap };
}

// ═══════════════════════════════════════════════════════════════════════════════
// 3. LÊ MARCOS
// ═══════════════════════════════════════════════════════════════════════════════
function lerMarcos(ss) {
  const ws   = ss.getSheetByName(CFG.sheetMarcos);
  const data = ws.getDataRange().getValues();
  const marcos = [];

  for (let r = CFG.marcosDataRow - 1; r < data.length; r++) {
    const row  = data[r];
    const proj = row[0] ? String(row[0]).trim() : '';
    if (!proj) continue;

    const nome   = String(row[2]  || '').trim();
    const dataV  = row[3];
    if (!nome) continue;

    let mdate;
    if (dataV instanceof Date)
      mdate = new Date(dataV.getFullYear(), dataV.getMonth(), dataV.getDate());
    else continue;

    const dur        = row[4] ? Math.max(1, parseInt(row[4])) : 1;
    const presV      = String(row[5] || '').trim().toLowerCase();
    const presencial = ['sim','yes','s','true','1'].includes(presV);
    const dev1       = String(row[6] || '').trim();
    const dev2       = String(row[7] || '').trim();
    const recV       = String(row[8] || '').trim();
    const status     = String(row[9]  || 'Planejado').trim();
    const fase       = String(row[1]  || '').trim();
    const tipo       = String(row[11] || fase).trim();

    const mandatory = [dev1, dev2].filter(Boolean);
    const rec       = recV ? recV.split(',').map(s => s.trim()).filter(Boolean) : [];

    marcos.push({ proj, fase, nome, date: mdate, dur, presencial, mandatory, rec, status, tipo });
  }
  return marcos;
}

function getProjetosOrdem(marcos) {
  const seen = new Set(), order = [];
  marcos.forEach(ms => { if (!seen.has(ms.proj)) { seen.add(ms.proj); order.push(ms.proj); } });
  return order;
}

// ═══════════════════════════════════════════════════════════════════════════════
// 4. UTILITÁRIOS DE DATA
// ═══════════════════════════════════════════════════════════════════════════════
function dateKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function fmtDate(d) {
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}`;
}
function addDays(d, n) {
  const r = new Date(d); r.setDate(r.getDate() + n); return r;
}
function isWeekend(d) { return d.getDay() === 0 || d.getDay() === 6; }

// Retorna data exclusiva de fim: conta "dur" dias úteis (ou corridos se presencial=true)
function calcBarEndExclusive(startDate, dur, presencial) {
  if (presencial) return addDays(startDate, dur);
  let count = 0;
  let d = new Date(startDate);
  while (count < dur) {
    if (!isWeekend(d)) count++;
    if (count < dur) d = addDays(d, 1);
  }
  return addDays(d, 1);
}

function msEnd(ms) {
  if (ms.presencial) return addDays(ms.date, ms.dur - 1);
  return addDays(calcBarEndExclusive(ms.date, ms.dur, false), -1);
}
function isPast(ms, ganttStart) { return msEnd(ms) < ganttStart; }

function getGanttDays() {
  const today = new Date(); today.setHours(0,0,0,0);
  const wd = today.getDay();
  let ganttStart;
  if      (wd === 0) ganttStart = addDays(today,  1);
  else if (wd === 6) ganttStart = addDays(today,  2);
  else               ganttStart = addDays(today, -(wd - 1));

  const days = Array.from({ length: CFG.calendarDays }, (_, i) => addDays(ganttStart, i));
  return { today, ganttStart, days };
}

// Índices de dias (0-based) cobertos por um marco
// presencial=true → dias corridos; false → dur em dias úteis (fins de semana incluídos no span visual)
function getDayIndices(days, startDate, dur, presencial) {
  const end = calcBarEndExclusive(startDate, dur, presencial !== undefined ? presencial : true);
  return days.reduce((acc, d, i) => {
    if (d >= startDate && d < end) acc.push(i);
    return acc;
  }, []);
}

// ═══════════════════════════════════════════════════════════════════════════════
// 5. ALOCAÇÃO / CONFLITOS
// ═══════════════════════════════════════════════════════════════════════════════
const _personRowCache = {};
function findPersonRow(personRowMap, shortNameStr) {
  if (!shortNameStr) return null;
  if (_personRowCache[shortNameStr] !== undefined) return _personRowCache[shortNameStr];
  const shortParts = normName(shortNameStr).split(/\s+/).filter(Boolean);
  let bestRow = null, bestScore = 0;
  for (const [full, idx] of Object.entries(personRowMap)) {
    const fullSet = new Set(normName(full).split(/\s+/).filter(Boolean));
    const score   = shortParts.filter(p => fullSet.has(p)).length;
    if (score > bestScore) { bestScore = score; bestRow = idx; }
  }
  _personRowCache[shortNameStr] = bestRow;
  return bestRow;
}

function normName(s) {
  return s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();
}

function getAlloc(plano, person, d) {
  const rowIdx = findPersonRow(plano.personRowMap, person);
  if (rowIdx === null) return null;
  const colIdx = plano.dateColMap[dateKey(d)];
  if (colIdx === undefined) return null;
  const v = plano.data[rowIdx][colIdx];
  return v ? String(v).trim() : null;
}

function isRestricted(plano, person, d) {
  return RESTRICTION_TYPES.has(getAlloc(plano, person, d));
}

function detectarConflitos(marcos, plano) {
  const conflitos = [], seen = new Set();
  for (const ms of marcos) {
    for (const person of ms.mandatory) {
      const key = `${person}||${ms.nome}||${ms.proj}`;
      if (seen.has(key)) continue;
      for (let dd = 0; dd < ms.dur; dd++) {
        const checkDate = addDays(ms.date, dd);
        if (isWeekend(checkDate)) continue;
        if (isRestricted(plano, person, checkDate)) {
          const severity = ['Go-live','Start UP','UAT','Testes'].includes(ms.tipo) ? 'Alta' : 'Média';
          conflitos.push({ person, project: ms.proj, milestone: ms.nome,
            date: ms.date, restriction: getAlloc(plano, person, checkDate),
            presencial: ms.presencial, severity });
          seen.add(key);
          break;
        }
      }
    }
  }
  return conflitos;
}

// ═══════════════════════════════════════════════════════════════════════════════
// 6. NOMES ABREVIADOS
// ═══════════════════════════════════════════════════════════════════════════════
function shortName(full) {
  const p = full.trim().split(/\s+/);
  return p.length >= 2 ? `${p[0]} ${p[p.length-1][0]}.` : full;
}
function formatResp(mandatory, rec) {
  return [...mandatory.map(shortName), ...rec.map(n => `(${shortName(n)})`)].join(', ');
}

// ═══════════════════════════════════════════════════════════════════════════════
// 7. HELPER: cria/limpa aba
// ═══════════════════════════════════════════════════════════════════════════════
function getOrCreateSheet(ss, name, tabColor) {
  let ws = ss.getSheetByName(name);
  if (ws) {
    ws.clearContents();
    ws.clearFormats();
    try { ws.getRange(1, 1, ws.getMaxRows(), ws.getMaxColumns()).breakApart(); } catch(e) {}
  } else {
    ws = ss.insertSheet(name);
  }
  if (tabColor) ws.setTabColor(tabColor);
  ws.setHiddenGridlines(true);
  return ws;
}

function styleRange(range, opts) {
  if (opts.value !== undefined) range.setValue(opts.value);
  if (opts.bg)                  range.setBackground(opts.bg);
  if (opts.fg)                  range.setFontColor(opts.fg);
  if (opts.bold  !== undefined) range.setFontWeight(opts.bold ? 'bold' : 'normal');
  if (opts.sz)                  range.setFontSize(opts.sz);
  if (opts.align)               range.setHorizontalAlignment(opts.align);
  if (opts.italic)              range.setFontStyle('italic');
  range.setVerticalAlignment('middle').setFontFamily('Arial');
  return range;
}

// ═══════════════════════════════════════════════════════════════════════════════
// 8. ABA GANTT — DIA A DIA
// ═══════════════════════════════════════════════════════════════════════════════
function gerarAbaGantt(ss, marcos, projetos, conflitos, plano, legendaCores) {
  const ws = getOrCreateSheet(ss, CFG.sheetGantt, '#1F3864');
  const { today, ganttStart, days } = getGanttDays();
  const N     = days.length;     // total de dias
  const L     = 3;               // colunas fixas: A=Marco, B=Resp, C=Dur
  const DCOL  = L + 1;          // primeira coluna de dia (1-indexed)
  const TOTAL = L + N;
  const conflictSet = new Set(conflitos.map(c => `${c.milestone}||${c.project}`));

  // ── Garante colunas suficientes (Sheets cria aba com 26 por padrão) ────────
  const neededCols = TOTAL;
  const currentCols = ws.getMaxColumns();
  if (currentCols < neededCols) {
    ws.insertColumnsAfter(currentCols, neededCols - currentCols);
  }

  // ── Larguras de coluna ─────────────────────────────────────────────────────
  ws.setColumnWidth(1, 250);
  ws.setColumnWidth(2, 155);
  ws.setColumnWidth(3, 38);
  for (let i = 0; i < N; i++) {
    ws.setColumnWidth(DCOL + i, 28);
  }

  // ── Freeze antes de qualquer mescla ───────────────────────────────────────
  ws.setFrozenRows(3);
  ws.setFrozenColumns(L);

  let r = 1;

  // ── Linha 1: Título (partido no limite de freeze) ─────────────────────────
  ws.setRowHeight(r, 26);
  styleRange(ws.getRange(r, 1, 1, L).merge(), {
    value: `GANTT DE MARCOS — Engenharia TI | ${fmtDate(today)}/${today.getFullYear()}`,
    bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 12, align: 'left',
  });
  ws.getRange(r, DCOL, 1, N).merge().setBackground('#1F3864');
  r++;

  // ── Linha 2: Faixas de mês ────────────────────────────────────────────────
  ws.setRowHeight(r, 13);
  ws.getRange(r, 1, 1, L).merge().setBackground('#1F3864');

  let curMonth = -1, monthStartCol = DCOL;
  for (let i = 0; i < N; i++) {
    const d   = days[i];
    const col = DCOL + i;
    if (d.getMonth() !== curMonth) {
      if (curMonth >= 0) {
        styleRange(ws.getRange(r, monthStartCol, 1, col - monthStartCol).merge(), {
          value: `${MONTHS_PT[curMonth].toUpperCase()} ${days[i-1].getFullYear()}`,
          bg: '#2E75B6', fg: '#FFFFFF', bold: true, sz: 8, align: 'center',
        });
      }
      curMonth = d.getMonth(); monthStartCol = col;
    }
  }
  styleRange(ws.getRange(r, monthStartCol, 1, DCOL + N - monthStartCol).merge(), {
    value: `${MONTHS_PT[curMonth].toUpperCase()} ${days[N-1].getFullYear()}`,
    bg: '#2E75B6', fg: '#FFFFFF', bold: true, sz: 8, align: 'center',
  });
  r++;

  // ── Linha 3: Cabeçalhos de coluna + dias ──────────────────────────────────
  ws.setRowHeight(r, 30);

  // Colunas fixas
  [['Projeto / Marco',1],['Responsáveis',2],['Dur.',3]].forEach(([v,c]) => {
    styleRange(ws.getRange(r, c), {
      value: v, bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 9, align: 'center',
    });
  });

  // Colunas de dia — batch de valores, backgrounds, fontColors
  const hdrVals = [], hdrBgs = [], hdrFgs = [], hdrBolds = [];
  for (let i = 0; i < N; i++) {
    const d     = days[i];
    const wkEnd = isWeekend(d);
    const isNow = dateKey(d) === dateKey(today);
    hdrVals.push(wkEnd ? DAYS_PT[d.getDay()] : String(d.getDate()));
    hdrBgs.push(isNow ? '#4472C4' : (wkEnd ? '#555555' : '#1F3864'));
    hdrFgs.push('#FFFFFF');
    hdrBolds.push(isNow ? 'bold' : 'normal');
  }
  ws.getRange(r, DCOL, 1, N)
    .setValues([hdrVals]).setBackgrounds([hdrBgs])
    .setFontColors([hdrFgs]).setFontWeights([hdrBolds])
    .setFontSize(7).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontFamily('Arial');
  r++;

  // ── Linhas de dados ────────────────────────────────────────────────────────
  const rowsToHide = [];

  for (const proj of projetos) {
    const projMs = marcos.filter(ms => ms.proj === proj);
    if (projMs.every(ms => isPast(ms, ganttStart))) continue;

    const [pb, pf] = getProjColor(proj, legendaCores);
    const pbLight  = lightenHex(pb, 0.72);

    // ── Cabeçalho do projeto ──────────────────────────────────────────────
    ws.setRowHeight(r, 20);
    styleRange(ws.getRange(r, 1, 1, L).merge(), {
      value: `  ${proj.toUpperCase()}`, bg: pb, fg: pf, bold: true, sz: 10, align: 'left',
    });
    // Pinta todo o bloco de dias com a cor do projeto — cor sólida, sem separadores
    ws.getRange(r, DCOL, 1, N).setBackground(pb)
      .setBorder(true, false, true, true, false, false, pb, SpreadsheetApp.BorderStyle.SOLID);
    r++;

    // ── Marcos do projeto ─────────────────────────────────────────────────
    for (const ms of projMs) {
      const past    = isPast(ms, ganttStart);
      const tipo    = ms.tipo || ms.fase;
      let   barBg   = TIPO_COLOR[tipo] || pb;
      const hasConf = conflictSet.has(`${ms.nome}||${ms.proj}`);
      if (hasConf) barBg = '#C00000';

      const dayIdxs   = new Set(getDayIndices(days, ms.date, ms.dur, ms.presencial));
      const minDayIdx = dayIdxs.size ? Math.min(...dayIdxs) : -1;
      const resp      = formatResp(ms.mandatory, ms.rec);
      const durLbl    = ms.dur > 1 ? `${ms.dur}d` : '—';
      const prefix    = ms.presencial ? '(P) ' : '+ ';

      ws.setRowHeight(r, 20);

      // Col A: nome do marco
      styleRange(ws.getRange(r, 1), {
        value: `  ${prefix}${ms.nome}`,
        bg:   past ? '#F5F5F5' : (hasConf ? '#FDECEA' : '#FFFFFF'),
        fg:   past ? '#AAAAAA' : (hasConf ? '#C00000' : '#222222'),
        bold: !past, sz: 9, align: 'left', italic: ms.presencial,
      });
      // Col B: responsáveis
      styleRange(ws.getRange(r, 2), {
        value: resp, bg: '#F5F5F5', fg: past ? '#CCCCCC' : '#555555',
        bold: false, sz: 8, align: 'left',
      });
      // Col C: duração
      styleRange(ws.getRange(r, 3), {
        value: durLbl, bg: '#F5F5F5', fg: past ? '#CCCCCC' : '#555555',
        bold: false, sz: 8, align: 'center',
      });

      // Colunas de dia (batch)
      const msBgs = [], msVals = [], msFgs = [], msBolds = [];
      for (let i = 0; i < N; i++) {
        const d      = days[i];
        const inBar  = dayIdxs.has(i);
        const wkEnd  = isWeekend(d);

        if (inBar) {
          const bg = past ? '#CCCCCC' : barBg;
          msBgs.push(bg);
          msFgs.push('#FFFFFF');
          msBolds.push('bold');
          msVals.push(i === minDayIdx ? fmtDate(ms.date) : '');
        } else if (wkEnd) {
          msBgs.push('#D9D9D9');
          msFgs.push('#AAAAAA');
          msBolds.push('normal');
          msVals.push('');
        } else {
          msBgs.push(i % 2 === 0 ? '#F5F5F5' : '#FAFAFA');
          msFgs.push('#CCCCCC');
          msBolds.push('normal');
          msVals.push('');
        }
      }
      ws.getRange(r, DCOL, 1, N)
        .setValues([msVals]).setBackgrounds([msBgs])
        .setFontColors([msFgs]).setFontWeights([msBolds])
        .setFontSize(7).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setFontFamily('Arial')
        .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
      ws.getRange(r, 1, 1, L)
        .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

      // ── Mescla células da barra quando duração > 1 dia ───────────────
      if (dayIdxs.size > 1) {
        const minIdx = Math.min(...dayIdxs);
        const maxIdx = Math.max(...dayIdxs);
        ws.getRange(r, DCOL + minIdx, 1, maxIdx - minIdx + 1)
          .merge()
          .setBackground(past ? '#CCCCCC' : barBg)
          .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(7)
          .setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setFontFamily('Arial')
          .setValue(fmtDate(ms.date));
      }

      if (past) rowsToHide.push(r);
      r++;

      // ── Sub-linhas por dev obrigatório ────────────────────────────────
      // Mostra a alocação diária de cada dev: férias ficam visíveis no Gantt
      for (const dev of ms.mandatory) {
        ws.setRowHeight(r, 14);

        // Col A: nome do dev (indentado)
        styleRange(ws.getRange(r, 1), {
          value: `       ↳ ${shortName(dev)}`,
          bg: '#F8F8F8', fg: '#666666', bold: false, sz: 8, align: 'left',
        });
        ws.getRange(r, 2).setBackground('#F8F8F8');
        ws.getRange(r, 3).setBackground('#F8F8F8');

        // Colunas de dia para o dev (batch)
        const devBgs = [], devVals = [], devFgs = [];
        for (let i = 0; i < N; i++) {
          const d      = days[i];
          const wkEnd  = isWeekend(d);
          const inMs   = dayIdxs.has(i);   // dentro do período do marco

          if (wkEnd) {
            devBgs.push('#D9D9D9'); devVals.push(''); devFgs.push('#AAAAAA');
            continue;
          }

          const alloc = getAlloc(plano, dev, d);

          if (!alloc) {
            // Sem alocação registrada
            devBgs.push(inMs ? pbLight : '#F8F8F8');
            devVals.push('');
            devFgs.push('#AAAAAA');
          } else if (RESTRICTION_TYPES.has(alloc)) {
            // FÉRIAS / FOLGA / etc — destaque vermelho/laranja
            const isFer = alloc === 'Férias';
            devBgs.push(isFer ? '#E06666' : '#F6B26B');
            devVals.push(RESTR_ABBR[alloc] || alloc.slice(0,3));
            devFgs.push('#FFFFFF');
          } else if (inMs) {
            // Trabalhando neste período — cor clara do projeto
            devBgs.push(pbLight);
            devVals.push('');
            devFgs.push('#888888');
          } else {
            // Período fora do marco — apenas marca de alocação
            devBgs.push('#F0F0F0');
            devVals.push('');
            devFgs.push('#CCCCCC');
          }
        }
        ws.getRange(r, DCOL, 1, N)
          .setValues([devVals]).setBackgrounds([devBgs])
          .setFontColors([devFgs]).setFontWeights([Array(N).fill('bold')])
          .setFontSize(6).setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setFontFamily('Arial')
          .setBorder(false, true, true, true, true, true, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
        ws.getRange(r, 1, 1, L)
          .setBorder(false, true, true, true, true, true, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);

        if (past) rowsToHide.push(r);
        r++;
      }
    }

    // Espaçador entre projetos
    ws.setRowHeight(r, 5);
    ws.getRange(r, 1, 1, TOTAL).setBackground('#D8D8D8');
    r++;
  }

  // ── Oculta marcos e devs passados ─────────────────────────────────────────
  if (rowsToHide.length > 0) {
    let start = rowsToHide[0], prev = rowsToHide[0];
    for (let i = 1; i <= rowsToHide.length; i++) {
      if (i === rowsToHide.length || rowsToHide[i] !== prev + 1) {
        ws.hideRows(start, prev - start + 1);
        if (i < rowsToHide.length) start = rowsToHide[i];
      }
      if (i < rowsToHide.length) prev = rowsToHide[i];
    }
  }

  // ── Legenda — contida em A:C, uma linha por item ─────────────────────────
  r++;
  ws.setRowHeight(r, 15);
  styleRange(ws.getRange(r, 1, 1, L).merge(), {
    value: ' LEGENDA', bg: '#2D2D2D', fg: '#FFFFFF', bold: true, sz: 8, align: 'left',
  });
  r++;

  [
    ['■ Marco crítico (Go-live/Start-up)', '#C00000', '#FFFFFF'],
    ['■ Marco normal (Testes/Dev/Entrega)', '#70AD47', '#FFFFFF'],
    ['■ Marco Demo / Apresentação',         '#7030A0', '#FFFFFF'],
    ['■ Marco Documentação',                '#BF8F00', '#FFFFFF'],
    ['■ (P) Presencial',                   '#444444', '#FFFFFF'],
    ['■ Dev: Férias',                       '#E06666', '#FFFFFF'],
    ['■ Dev: Outra restrição (ASO, Folga…)','#F6B26B', '#FFFFFF'],
    ['■ Dev: Trabalhando no projeto',       '#D9EAD3', '#444444'],
    ['■ Conflito detectado (marco)',        '#FDECEA', '#C00000'],
  ].forEach(([lbl, bg, fg]) => {
    ws.setRowHeight(r, 13);
    styleRange(ws.getRange(r, 1, 1, L).merge(), {
      value: lbl, bg, fg, bold: true, sz: 7, align: 'left',
    });
    r++;
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// 9. ABA CONFLITOS
// ═══════════════════════════════════════════════════════════════════════════════
function gerarAbaConflitos(ss, conflitos) {
  const ws = getOrCreateSheet(ss, CFG.sheetConflitos, conflitos.length ? '#C00000' : '#38761D');
  const n  = conflitos.length;
  if (ws.getMaxColumns() < 9) ws.insertColumnsAfter(ws.getMaxColumns(), 9 - ws.getMaxColumns());

  ws.setRowHeight(1, 26);
  styleRange(ws.getRange(1, 1, 1, 9).merge(), {
    value: `CONFLITOS — ${n} detectado(s) | Cruzamento: Marcos × Alocação Diária`,
    bg: n ? '#C00000' : '#38761D', fg: '#FFFFFF', bold: true, sz: 12, align: 'left',
  });
  ws.setRowHeight(2, 6);

  const hdrs   = ['Dev','Projeto','Marco','Data','Restrição','Presencial?','Severidade','Status','Substituto / Ação'];
  const widths = [160, 120, 220, 90, 180, 80, 100, 130, 200];
  hdrs.forEach((h, i) => {
    ws.setColumnWidth(i+1, widths[i]);
    styleRange(ws.getRange(3, i+1), { value: h, bg: '#C00000', fg: '#FFFFFF', bold: true, sz: 9, align: 'center' });
  });
  ws.setRowHeight(3, 20);
  ws.setFrozenRows(3);

  if (!n) {
    styleRange(ws.getRange(4, 1, 1, 9).merge(), {
      value: 'Nenhum conflito detectado.',
      bg: '#D9EAD3', fg: '#38761D', bold: true, sz: 11, align: 'center',
    });
    ws.setRowHeight(4, 26);
    return;
  }

  conflitos.forEach((c, ri) => {
    const row   = 4 + ri;
    const alt   = ri % 2 === 1;
    const rowBg = alt ? '#FFF0F0' : '#FDECEA';
    const sevBg = c.severity === 'Alta' ? '#FDECEA' : '#FFF2CC';
    const sevFg = c.severity === 'Alta' ? '#C00000' : '#BF8F00';

    ws.setRowHeight(row, 20);
    [c.person, c.project, c.milestone, c.date, c.restriction,
     c.presencial ? 'Sim':'Não', c.severity, 'Em aberto', '']
      .forEach((v, ci) => {
        const cell = ws.getRange(row, ci+1);
        styleRange(cell, {
          value: v,
          bg:   ci === 6 ? sevBg : rowBg,
          fg:   ci === 6 ? sevFg : '#000000',
          bold: [0,2,6].includes(ci), sz: 10,
          align:[3,5,6,7].includes(ci) ? 'center' : 'left',
        });
        if (ci === 3 && v instanceof Date) cell.setNumberFormat('DD/MM/YYYY');
      });
  });

  ws.getRange(3, 1, n+1, 9)
    .setBorder(true,true,true,true,true,true,'#DDDDDD', SpreadsheetApp.BorderStyle.SOLID);
}

// ═══════════════════════════════════════════════════════════════════════════════
// 10. ABA DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════
function gerarAbaDashboard(ss, marcos, projetos, conflitos) {
  const ws    = getOrCreateSheet(ss, CFG.sheetDashboard, '#FF6D00');
  if (ws.getMaxColumns() < 8) ws.insertColumnsAfter(ws.getMaxColumns(), 8 - ws.getMaxColumns());
  const today = new Date(); today.setHours(0,0,0,0);
  const n     = conflitos.length;
  const ms30  = marcos
    .filter(ms => { const d = (ms.date - today)/86400000; return d >= 0 && d <= 30; })
    .sort((a,b) => a.date - b.date);
  const conflictSet = new Set(conflitos.map(c => `${c.milestone}||${c.project}`));

  ws.setRowHeight(1, 28);
  styleRange(ws.getRange(1, 1, 1, 8).merge(), {
    value: 'DASHBOARD — Visão executiva | Próximos 30 dias | Riscos por projeto',
    bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 13, align: 'left',
  });
  ws.setRowHeight(2, 8);

  const kpis = [
    ['Total marcos',    marcos.length,                         '#2E75B6'],
    ['Próx. 30 dias',   ms30.length,                          '#0070C0'],
    ['Conflitos',       n,                n ? '#C00000' : '#38761D'],
    ['Presenciais',     marcos.filter(m => m.presencial).length, '#7030A0'],
    ['Projetos',        projetos.length,                       '#1F3864'],
  ];
  kpis.forEach(([lbl, val, color], ki) => {
    const col = ki+1;
    ws.setColumnWidth(col, 120);
    styleRange(ws.getRange(3, col), { value: lbl, bg: color, fg: '#FFFFFF', bold: true, sz: 9, align: 'center' });
    styleRange(ws.getRange(4, col), { value: val, bg: '#FFFFFF', fg: color, bold: true, sz: 22, align: 'center' });
    ws.setRowHeight(3, 18); ws.setRowHeight(4, 40);
  });

  ws.setRowHeight(6, 12);
  ws.setRowHeight(7, 20);
  styleRange(ws.getRange(7, 1, 1, 8).merge(), {
    value: '  PRÓXIMOS 30 DIAS', bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });

  const dhHdrs   = ['Data','Projeto','Marco','Tipo','Presencial?','Devs obrigatórios','Conflito?','Dias restantes'];
  const dhWidths = [90,120,220,130,80,240,80,100];
  dhHdrs.forEach((h, i) => {
    ws.setColumnWidth(i+1, dhWidths[i]);
    styleRange(ws.getRange(8, i+1), { value: h, bg: '#2E75B6', fg: '#FFFFFF', bold: true, sz: 9, align: 'center' });
  });
  ws.setRowHeight(8, 20);

  if (!ms30.length) {
    styleRange(ws.getRange(9, 1, 1, 8).merge(), {
      value: 'Nenhum marco nos próximos 30 dias.', bg: '#D9EAD3', fg: '#38761D', sz: 11, align: 'center',
    });
  } else {
    ms30.forEach((ms, ri) => {
      const row     = 9 + ri;
      const diasR   = Math.round((ms.date - today)/86400000);
      const hasConf = conflictSet.has(`${ms.nome}||${ms.proj}`);
      const tipoBg  = TIPO_COLOR[ms.tipo] || '#2E75B6';
      const altBg   = ri%2===1 ? '#DEEAF1' : '#FFFFFF';
      ws.setRowHeight(row, 20);

      [
        { v: ms.date, bg: altBg, fg: '#000000', bold: false, fmt: 'DD/MM/YYYY', align: 'center' },
        { v: ms.proj, bg: altBg, fg: '#000000', bold: false,  align: 'left' },
        { v: ms.nome, bg: altBg, fg: '#000000', bold: true,   align: 'left' },
        { v: ms.tipo, bg: tipoBg,fg: '#FFFFFF', bold: true,   align: 'center' },
        { v: ms.presencial ? 'Sim':'Não', bg: ms.presencial ? '#EAD1F7':altBg, fg: ms.presencial ? '#7030A0':'#000000', bold: false, align: 'center' },
        { v: ms.mandatory.join(', '), bg: altBg, fg: '#000000', bold: false, align: 'left' },
        { v: hasConf ? '⚠ SIM':'OK', bg: hasConf ? '#FDECEA':'#D9EAD3', fg: hasConf ? '#C00000':'#38761D', bold: true, align: 'center' },
        { v: diasR, bg: diasR<7 ? '#FFF2CC':altBg, fg: diasR<7 ? '#BF8F00':'#000000', bold: diasR<7, align: 'center' },
      ].forEach((d, ci) => {
        const cell = ws.getRange(row, ci+1);
        styleRange(cell, { value: d.v, bg: d.bg, fg: d.fg, bold: d.bold||false, sz: 10, align: d.align });
        if (d.fmt) cell.setNumberFormat(d.fmt);
      });
    });
  }

  const gapRow = 10 + Math.max(1, ms30.length) + 1;
  ws.setRowHeight(gapRow, 20);
  styleRange(ws.getRange(gapRow, 1, 1, 8).merge(), {
    value: '  RISCO POR PROJETO', bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  ['Projeto','Marcos','Próx. marco','Conflitos','Presenciais','Status'].forEach((h,i) => {
    styleRange(ws.getRange(gapRow+1, i+1), { value: h, bg: '#2E75B6', fg: '#FFFFFF', bold: true, sz: 9, align: 'center' });
  });
  ws.setRowHeight(gapRow+1, 20);

  projetos.forEach((proj, pi) => {
    const row    = gapRow + 2 + pi;
    const pms    = marcos.filter(m => m.proj === proj);
    const future = pms.filter(m => m.date >= today).sort((a,b) => a.date - b.date);
    const nConf  = conflitos.filter(c => c.project === proj).length;
    const nPres  = pms.filter(m => m.presencial).length;
    const goLive = future.slice(0,2).some(m => ['Go-live','Start UP'].includes(m.tipo));
    const status = nConf > 0 ? '⚠ Em risco' : (goLive ? '★ Go-live iminente' : '✓ Normal');
    const sBg    = nConf > 0 ? '#FDECEA' : (goLive ? '#FFF2CC' : '#D9EAD3');
    const sFg    = nConf > 0 ? '#C00000' : (goLive ? '#BF8F00' : '#38761D');
    const altBg  = pi%2===1 ? '#DEEAF1' : '#FFFFFF';
    ws.setRowHeight(row, 22);
    [
      { v: proj,                     bg: '#1F3864',   fg: '#FFFFFF', bold: true },
      { v: pms.length,               bg: altBg,       fg: '#000000', bold: false },
      { v: future[0]?.date || '—',   bg: altBg,       fg: '#000000', bold: false, fmt: future[0] ? 'DD/MM/YYYY' : null },
      { v: nConf, bg: nConf>0?'#FDECEA':altBg, fg: nConf>0?'#C00000':'#000000', bold: nConf>0 },
      { v: nPres, bg: nPres>0?'#EAD1F7':altBg, fg: nPres>0?'#7030A0':'#000000', bold: false },
      { v: status,                   bg: sBg,         fg: sFg,       bold: true },
    ].forEach((d, ci) => {
      const cell = ws.getRange(row, ci+1);
      styleRange(cell, { value: d.v, bg: d.bg, fg: d.fg, bold: d.bold, sz: 10, align: ci===0?'left':'center' });
      if (d.fmt) cell.setNumberFormat(d.fmt);
    });
  });
}
