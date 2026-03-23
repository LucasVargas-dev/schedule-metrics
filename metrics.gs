// ═══════════════════════════════════════════════════════════════════════════════
//  MÉTRICAS TI — Google Apps Script
//  Lê: Quadro TI Normalizado
//  Gera: Resumo Executivo | Por Pessoa | Por Cliente | Alertas | Horas/Semana
// ═══════════════════════════════════════════════════════════════════════════════

// Statuses onde o trabalho NÃO começou → excluídos do comparativo de horas
const STATUS_NAO_FEITAS = new Set(['Backlog', 'A fazer', 'Em andamento']);

const CAP_SEMANA = 44;  // capacidade semanal em horas
const MIN_SEMANA = 36;  // mínimo aceitável por semana

// ─── ENTRY POINT ─────────────────────────────────────────────────────────────
function gerarMetricas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Lendo dados normalizados...', 'Gerando Métricas', -1);

  const dados = lerNormalizadoM(ss);
  if (!dados.length) {
    SpreadsheetApp.getUi().alert(
      'Aba "Quadro TI Normalizado" vazia ou não encontrada.\nExecute "Normalizar dados" primeiro.'
    );
    return;
  }

  gerarResumoM(ss, dados);
  gerarPorPessoaM(ss, dados);
  gerarPorClienteM(ss, dados);
  gerarConcluidasM(ss, dados);
  gerarAlertasM(ss, dados);
  gerarHorasSemanaM(ss, dados);

  ss.toast(`${dados.length} tarefas processadas`, 'Métricas geradas!', 6);
}

// ─── LÊ QUADRO TI NORMALIZADO ────────────────────────────────────────────────
// Detecta a linha de cabeçalho (contém 'Nome') e lê os dados a partir da próxima linha
function lerNormalizadoM(ss) {
  const ws = ss.getSheetByName('Quadro TI Normalizado');
  if (!ws) return [];
  const raw = ws.getDataRange().getValues();

  let dataStart = -1;
  for (let i = 0; i < raw.length; i++) {
    if (String(raw[i][0] || '').trim().toLowerCase() === 'nome') {
      dataStart = i + 1;
      break;
    }
  }
  if (dataStart < 0) return [];

  return raw.slice(dataStart)
    .filter(r => r[0] && r[0].toString().trim() !== '')
    .map(r => ({
      nome:        String(r[0]  || ''),
      cliente:     String(r[1]  || 'Sem cliente'),
      projeto:     String(r[2]  || ''),
      pessoa:      String(r[3]  || 'Sem responsável'),
      status:      String(r[4]  || 'Backlog'),
      dificuldade: String(r[5]  || ''),
      peso:        toN(r[6]),
      horasEst:    toN(r[7]),
      horasReal:   toN(r[8]),
      delta:       toN(r[9]),
      tipo:        String(r[10] || ''),
      prioridade:  String(r[11] || ''),
      dataCriacao: r[12] instanceof Date ? r[12] : null,
      dataFinal:   r[13] instanceof Date ? r[13] : null,
      atrasado:    r[14] === true || r[14] === '⚠ SIM',
    }));
}

function toN(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

// Helper: "João, Maria" → ['João', 'Maria']
function splitPessoasM(str) {
  return str.split(',').map(p => p.trim()).filter(p => p && p !== 'Sem responsável');
}

function pctM(part, total) {
  if (!total) return '0%';
  return `${Math.round(part / total * 100)}%`;
}

// ─── WEEK HELPERS ─────────────────────────────────────────────────────────────
function weekMonday(d) {
  const dt = new Date(d); dt.setHours(0, 0, 0, 0);
  const day = dt.getDay(); // 0=Dom
  dt.setDate(dt.getDate() + (day === 0 ? -6 : 1 - day));
  return dt;
}

function weekLabelM(mon) {
  const sun = new Date(mon); sun.setDate(sun.getDate() + 6);
  const f = d => `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}`;
  return `${f(mon)}–${f(sun)}`;
}

// ═══════════════════════════════════════════════════════════════════════════════
// 1. RESUMO EXECUTIVO
// ═══════════════════════════════════════════════════════════════════════════════
function gerarResumoM(ss, dados) {
  const ws    = getOrCreateSheet(ss, '📊 Resumo Executivo', '#1F3864');
  const hoje  = new Date(); hoje.setHours(0, 0, 0, 0);
  const dataStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  const total     = dados.length;
  const fechadas  = dados.filter(d => d.status === 'Closed').length;
  const andamento = dados.filter(d => d.status === 'Em andamento').length;
  const afazer    = dados.filter(d => d.status === 'A fazer' || d.status === 'Backlog').length;
  const aguard    = dados.filter(d => d.status.startsWith('Aguardando')).length;
  const atrasadas = dados.filter(d => d.atrasado).length;

  // Horas somente de tarefas que passaram do estágio inicial
  const horasDados  = dados.filter(d => !STATUS_NAO_FEITAS.has(d.status));
  const totEst      = horasDados.reduce((s, d) => s + d.horasEst,  0);
  const totReal     = horasDados.reduce((s, d) => s + d.horasReal, 0);
  const comDelta    = horasDados.filter(d => d.horasEst > 0 && d.horasReal > 0);
  const desvioMedio = comDelta.length
    ? comDelta.reduce((s, d) => s + (d.horasReal - d.horasEst) / d.horasEst, 0) / comDelta.length
    : 0;

  const NCOLS = 14;
  if (ws.getMaxColumns() < NCOLS) ws.insertColumnsAfter(ws.getMaxColumns(), NCOLS - ws.getMaxColumns());

  // Row 1: título
  ws.setRowHeight(1, 30);
  styleRange(ws.getRange(1, 1, 1, NCOLS).merge(), {
    value: `📊  MÉTRICAS DE EQUIPE TI  —  Gerado em ${dataStr}`,
    bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 13, align: 'left',
  });

  // Row 2: spacer
  ws.setRowHeight(2, 8);
  ws.getRange(2, 1, 1, NCOLS).setBackground('#1F3864');

  // Row 3: VISÃO GERAL
  ws.setRowHeight(3, 20);
  styleRange(ws.getRange(3, 1, 1, NCOLS).merge(), {
    value: '  VISÃO GERAL', bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });

  // Rows 4–5: KPI cards (6 cards em 6 colunas)
  const kpis = [
    ['TOTAL TAREFAS',  total,     '#1F3864'],
    ['CONCLUÍDAS',     fechadas,  '#38761D'],
    ['EM ANDAMENTO',   andamento, '#0070C0'],
    ['A FAZER',        afazer,    '#595959'],
    ['AGUARDANDO',     aguard,    '#BF8F00'],
    ['ATRASADAS',      atrasadas, atrasadas > 0 ? '#C00000' : '#38761D'],
  ];
  const kpiWidth = [130, 120, 120, 110, 110, 110];
  kpis.forEach(([lbl, val, color], ki) => {
    ws.setColumnWidth(ki + 1, kpiWidth[ki]);
    styleRange(ws.getRange(4, ki + 1), { value: lbl, bg: color,     fg: '#FFFFFF', bold: true, sz: 9,  align: 'center' });
    styleRange(ws.getRange(5, ki + 1), { value: val, bg: '#FFFFFF', fg: color,     bold: true, sz: 22, align: 'center' });
  });
  ws.setRowHeight(4, 18);
  ws.setRowHeight(5, 42);
  // Limpa colunas restantes dos KPIs
  ws.getRange(4, 7, 2, NCOLS - 6).setBackground('#FFFFFF');

  // Coluna 7 em diante: widths
  ws.setColumnWidth(7,  12);  // gap
  ws.setColumnWidth(8,  150); // label
  ws.setColumnWidth(9,  70);  // qtd
  ws.setColumnWidth(10, 70);  // %
  ws.setColumnWidth(11, 12);  // gap
  ws.setColumnWidth(12, 130); // label
  ws.setColumnWidth(13, 70);  // qtd
  ws.setColumnWidth(14, 70);  // %

  // Row 6: spacer
  ws.setRowHeight(6, 14);

  // ── Seção: DISTRIBUIÇÃO POR STATUS + TIPO + PRIORIDADE (3 tabelas lado a lado)
  // STATUS: cols 1-3  |  TIPO: cols 5-7  |  PRIORIDADE: cols 9-11
  ws.setColumnWidth(1, 160); ws.setColumnWidth(2, 70); ws.setColumnWidth(3, 60);
  ws.setColumnWidth(4, 14);
  ws.setColumnWidth(5, 150); ws.setColumnWidth(6, 70); ws.setColumnWidth(7, 60);
  ws.setColumnWidth(8, 14);
  ws.setColumnWidth(9, 140); ws.setColumnWidth(10, 70); ws.setColumnWidth(11, 60);
  ws.setColumnWidth(12, 14);

  let row = 7;
  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row, 1, 1, 3).merge(),  { value: '  DISTRIBUIÇÃO POR STATUS', bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left' });
  styleRange(ws.getRange(row, 5, 1, 3).merge(),  { value: '  DISTRIBUIÇÃO POR TIPO',   bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left' });
  styleRange(ws.getRange(row, 9, 1, 3).merge(),  { value: '  PRIORIDADE',               bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left' });
  ws.getRange(row, 4).setBackground('#FFFFFF');
  ws.getRange(row, 8).setBackground('#FFFFFF');
  ws.getRange(row, 12).setBackground('#FFFFFF');
  row++;

  ws.setRowHeight(row, 18);
  ['Status','Qtd','%'].forEach((h, i) => styleRange(ws.getRange(row, i+1), { value: h, bg: '#2E75B6', fg: '#FFFFFF', bold: true, sz: 9, align: 'center' }));
  ['Tipo','Qtd','%'].forEach((h, i)   => styleRange(ws.getRange(row, i+5), { value: h, bg: '#2E75B6', fg: '#FFFFFF', bold: true, sz: 9, align: 'center' }));
  ['Prioridade','Qtd','%'].forEach((h,i) => styleRange(ws.getRange(row, i+9), { value: h, bg: '#2E75B6', fg: '#FFFFFF', bold: true, sz: 9, align: 'center' }));
  ws.getRange(row, 4).setBackground('#2E75B6'); ws.getRange(row, 8).setBackground('#2E75B6'); ws.getRange(row, 12).setBackground('#FFFFFF');
  row++;

  // Monta listas
  const statusMap = {};
  STATUS_CANONICO.forEach(s => statusMap[s] = 0);
  dados.forEach(d => { if (d.status in statusMap) statusMap[d.status]++; });
  const statusList = STATUS_CANONICO.filter(s => statusMap[s] > 0).map(s => [s, statusMap[s]]);

  const tipoMap = {};
  dados.forEach(d => { const t = d.tipo || '(sem tipo)'; tipoMap[t] = (tipoMap[t]||0)+1; });
  const tipoList = Object.entries(tipoMap).sort((a,b) => b[1]-a[1]);

  const prioMap = {};
  dados.forEach(d => { const p = d.prioridade || '(sem prioridade)'; prioMap[p] = (prioMap[p]||0)+1; });
  const prioList = Object.entries(prioMap).sort((a,b) => b[1]-a[1]);

  const STATUS_BG = {
    'Closed':'#D9EAD3','Em andamento':'#DEEAF1','Em revisão':'#EAD1F7','Em testes':'#E6F3FF',
    'Ajustes de revisão':'#FDE9D9','A fazer':'#F4F4F4','Backlog':'#EEEEEE',
    'Aguardando':'#FFF3E0','Aguardando cliente':'#FFF3E0','Aguardando TA':'#FFF3E0',
  };

  const maxRows = Math.max(statusList.length, tipoList.length, prioList.length);
  for (let ri = 0; ri < maxRows; ri++) {
    const r = row + ri;
    ws.setRowHeight(r, 18);
    const alt = ri % 2 === 1 ? '#F0F4FA' : '#FFFFFF';

    if (ri < statusList.length) {
      const [st, cnt] = statusList[ri];
      styleRange(ws.getRange(r,1), { value: st,              bg: STATUS_BG[st]||alt, fg:'#000000', bold:false, sz:10, align:'left'   });
      styleRange(ws.getRange(r,2), { value: cnt,             bg: alt,                fg:'#000000', bold:false, sz:10, align:'center' });
      styleRange(ws.getRange(r,3), { value: pctM(cnt,total), bg: alt,                fg:'#595959', bold:false, sz:10, align:'center' });
    } else { ws.getRange(r,1,1,3).setBackground(alt); }

    ws.getRange(r,4).setBackground('#FFFFFF');

    if (ri < tipoList.length) {
      const [tp, cnt] = tipoList[ri];
      styleRange(ws.getRange(r,5), { value: tp,              bg: alt, fg:'#000000', bold:false, sz:10, align:'left'   });
      styleRange(ws.getRange(r,6), { value: cnt,             bg: alt, fg:'#000000', bold:false, sz:10, align:'center' });
      styleRange(ws.getRange(r,7), { value: pctM(cnt,total), bg: alt, fg:'#595959', bold:false, sz:10, align:'center' });
    } else { ws.getRange(r,5,1,3).setBackground(alt); }

    ws.getRange(r,8).setBackground('#FFFFFF');

    if (ri < prioList.length) {
      const [pr, cnt] = prioList[ri];
      styleRange(ws.getRange(r,9),  { value: pr,              bg: alt, fg:'#000000', bold:false, sz:10, align:'left'   });
      styleRange(ws.getRange(r,10), { value: cnt,             bg: alt, fg:'#000000', bold:false, sz:10, align:'center' });
      styleRange(ws.getRange(r,11), { value: pctM(cnt,total), bg: alt, fg:'#595959', bold:false, sz:10, align:'center' });
    } else { ws.getRange(r,9,1,3).setBackground(alt); }

    ws.getRange(r,12).setBackground('#FFFFFF');
  }

  row += maxRows + 1;
  ws.setRowHeight(row - 1, 12);

  // ── Seção: HORAS + DIFICULDADE
  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,3).merge(), { value:'  HORAS — ESTIMADO vs RASTREADO', bg:'#1F3864', fg:'#FFFFFF', bold:true, sz:10, align:'left' });
  styleRange(ws.getRange(row,5,1,3).merge(), { value:'  DIFICULDADE',                   bg:'#1F3864', fg:'#FFFFFF', bold:true, sz:10, align:'left' });
  ws.getRange(row,4).setBackground('#FFFFFF'); ws.getRange(row,8).setBackground('#FFFFFF');
  row++;

  ws.setRowHeight(row, 18);
  ['Métrica','Valor','Observação'].forEach((h,i) => styleRange(ws.getRange(row,i+1), { value:h, bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:9, align:'center' }));
  ['Dificuldade','Qtd','%'].forEach((h,i)     => styleRange(ws.getRange(row,i+5), { value:h, bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:9, align:'center' }));
  ws.getRange(row,4).setBackground('#2E75B6'); ws.getRange(row,8).setBackground('#2E75B6');
  row++;

  // Atualiza col widths para seção de horas
  ws.setColumnWidth(3, 220);

  const hrows = [
    ['Total estimado',    totEst,  `horas — tarefas fora de Backlog/A fazer/Em andamento`],
    ['Total rastreado',   totReal, totEst>0 ? `${Math.round(totReal/totEst*100)}% do estimado` : '—'],
    ['Diferença',         totReal - totEst, ''],
    ['Desvio médio %',    '',      `${(desvioMedio*100).toFixed(1)}%`],
    ['Tarefas OK (±10%)', comDelta.filter(d => Math.abs((d.horasReal-d.horasEst)/d.horasEst)<=0.10).length, ''],
    ['Tarefas acima',     comDelta.filter(d => (d.horasReal-d.horasEst)/d.horasEst>0.10).length,  '>10% acima do estimado'],
    ['Tarefas abaixo',    comDelta.filter(d => (d.horasEst-d.horasReal)/d.horasEst>0.10).length,  '>10% abaixo do estimado'],
    ['Sem rastreamento',  dados.filter(d => d.horasReal===0).length, 'tarefas sem horas reais'],
  ];

  hrows.forEach(([met, val, obs], ri) => {
    const r = row + ri;
    ws.setRowHeight(r, 18);
    const alt = ri%2===1 ? '#F0F4FA' : '#FFFFFF';
    const isNeg = typeof val === 'number' && val < 0;
    styleRange(ws.getRange(r,1), { value:met, bg:alt, fg:'#000000',              bold:true,  sz:10, align:'left'   });
    styleRange(ws.getRange(r,2), { value:val, bg:alt, fg:isNeg?'#C00000':'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(r,3), { value:obs, bg:alt, fg:'#595959',              bold:false, sz:9,  align:'left', italic:true });
    ws.getRange(r,4).setBackground('#FFFFFF');
  });

  const difMap = {};
  dados.forEach(d => { const df = d.dificuldade||'(não definida)'; difMap[df]=(difMap[df]||0)+1; });
  const difList = Object.entries(difMap).sort((a,b) => b[1]-a[1]);
  difList.forEach(([df, cnt], ri) => {
    const r = row + ri;
    ws.setRowHeight(r, 18);
    const alt = ri%2===1 ? '#F0F4FA' : '#FFFFFF';
    styleRange(ws.getRange(r,5),  { value:df,             bg:alt, fg:'#000000', bold:false, sz:10, align:'left'   });
    styleRange(ws.getRange(r,6),  { value:cnt,            bg:alt, fg:'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(r,7),  { value:pctM(cnt,total),bg:alt, fg:'#595959', bold:false, sz:10, align:'center' });
    ws.getRange(r,8).setBackground('#FFFFFF');
  });

  // ── Spacer after dificuldade section
  const rowAfterDif = row + Math.max(hrows.length, difList.length) + 1;
  ws.setRowHeight(rowAfterDif - 1, 12);

  // ── AGUARDANDO — ANÁLISE (cols 1-11)
  const rowAg = rowAfterDif;
  ws.setRowHeight(rowAg, 20);
  const aguardTodos = dados.filter(d => d.status.startsWith('Aguardando'));
  const aguardVencidos = aguardTodos.filter(d => d.dataFinal && d.dataFinal < hoje);
  const aguard14d = aguardTodos.filter(d => d.dataFinal && d.dataFinal >= hoje && Math.round((d.dataFinal-hoje)/86400000)<=14);
  const aguardSemPrazo = aguardTodos.filter(d => !d.dataFinal);

  styleRange(ws.getRange(rowAg,1,1,11).merge(), { value:'  TAREFAS AGUARDANDO — SITUAÇÃO', bg:'#BF8F00', fg:'#FFFFFF', bold:true, sz:10, align:'left' });
  ws.getRange(rowAg,12,1,3).setBackground('#FFFFFF');

  ws.setRowHeight(rowAg+1, 18);
  const agHdrs = ['Sub-status','Qtd','Vencidas','Vence ≤14d','Sem prazo'];
  agHdrs.forEach((h,i) => styleRange(ws.getRange(rowAg+1,i+1), { value:h, bg:'#BF8F00', fg:'#FFFFFF', bold:true, sz:9, align:'center' }));
  ws.getRange(rowAg+1,6,1,6).setBackground('#BF8F00');

  const agSubStatus = ['Aguardando','Aguardando cliente','Aguardando TA'];
  agSubStatus.forEach((sub, si) => {
    const r = rowAg + 2 + si;
    ws.setRowHeight(r, 18);
    const alt = si%2===1 ? '#FFF9E6' : '#FFFFFF';
    const subList = aguardTodos.filter(d => d.status === sub);
    const subVenc = subList.filter(d => d.dataFinal && d.dataFinal < hoje).length;
    const sub14d  = subList.filter(d => d.dataFinal && d.dataFinal >= hoje && Math.round((d.dataFinal-hoje)/86400000)<=14).length;
    const subSP   = subList.filter(d => !d.dataFinal).length;
    styleRange(ws.getRange(r,1), { value:sub,          bg:'#1F3864',                  fg:'#FFFFFF', bold:true,  sz:10, align:'left'   });
    styleRange(ws.getRange(r,2), { value:subList.length,bg:alt,                       fg:'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(r,3), { value:subVenc>0?subVenc:'—', bg:subVenc>0?'#FDECEA':alt, fg:subVenc>0?'#C00000':'#AAAAAA', bold:subVenc>0, sz:10, align:'center' });
    styleRange(ws.getRange(r,4), { value:sub14d>0?sub14d:'—',  bg:sub14d>0?'#FFF2CC':alt,  fg:sub14d>0?'#BF8F00':'#AAAAAA', bold:sub14d>0,  sz:10, align:'center' });
    styleRange(ws.getRange(r,5), { value:subSP>0?subSP:'—',    bg:subSP>0?'#FDECEA':alt,   fg:subSP>0?'#595959':'#AAAAAA', bold:false,      sz:10, align:'center' });
    ws.getRange(r,6,1,6).setBackground(alt);
  });

  // Totais aguardando
  const rTotAg = rowAg + 2 + agSubStatus.length;
  ws.setRowHeight(rTotAg, 18);
  styleRange(ws.getRange(rTotAg,1), { value:'TOTAL',                bg:'#7F6000', fg:'#FFFFFF', bold:true, sz:10, align:'left'   });
  styleRange(ws.getRange(rTotAg,2), { value:aguardTodos.length,     bg:'#7F6000', fg:'#FFFFFF', bold:true, sz:10, align:'center' });
  styleRange(ws.getRange(rTotAg,3), { value:aguardVencidos.length,  bg:aguardVencidos.length>0?'#C00000':'#7F6000', fg:'#FFFFFF', bold:true, sz:10, align:'center' });
  styleRange(ws.getRange(rTotAg,4), { value:aguard14d.length,       bg:aguard14d.length>0?'#BF8F00':'#7F6000', fg:'#FFFFFF', bold:true, sz:10, align:'center' });
  styleRange(ws.getRange(rTotAg,5), { value:aguardSemPrazo.length,  bg:'#7F6000', fg:'#FFFFFF', bold:true, sz:10, align:'center' });
  ws.getRange(rTotAg,6,1,6).setBackground('#7F6000');

  ws.setRowHeight(rTotAg+1, 12);

  // ── ENTREGUES (CLOSED) — QUALIDADE (cols 1-11)
  const rowCl = rTotAg + 2;
  const closed = dados.filter(d => d.status === 'Closed');
  const closedComH = closed.filter(d => d.horasEst > 0 && d.horasReal > 0);
  const closedComReal = closed.filter(d => d.horasReal > 0);
  const closedOK = closedComH.filter(d => Math.abs((d.horasReal-d.horasEst)/d.horasEst) <= 0.10);
  const closedDesvio = closedComH.length
    ? closedComH.reduce((s,d) => s+(d.horasReal-d.horasEst)/d.horasEst,0)/closedComH.length : 0;
  const pesoTotalEntregue = closed.reduce((s,d) => s + (d.peso||0), 0);

  ws.setRowHeight(rowCl, 20);
  styleRange(ws.getRange(rowCl,1,1,11).merge(), { value:'  ENTREGUES (CLOSED) — QUALIDADE', bg:'#38761D', fg:'#FFFFFF', bold:true, sz:10, align:'left' });
  ws.getRange(rowCl,12,1,3).setBackground('#FFFFFF');

  ws.setRowHeight(rowCl+1, 18);
  ['Métrica','Valor','Referência'].forEach((h,i) => styleRange(ws.getRange(rowCl+1,i*4+1), { value:h, bg:'#38761D', fg:'#FFFFFF', bold:true, sz:9, align:'center' }));

  const clRows = [
    ['Total entregues',         closed.length,             `${pctM(closed.length,total)} do total`],
    ['H. rastreadas (Closed)',  closedComReal.reduce((s,d)=>s+d.horasReal,0), 'horas totais'],
    ['Taxa de rastreamento',    '',                         `${pctM(closedComReal.length,closed.length)} das Closed têm horas`],
    ['Precisão ±10%',           closedOK.length,           `de ${closedComH.length} com dados`],
    ['Desvio médio',            '',                         `${(closedDesvio*100).toFixed(1)}%`],
    ['Peso total entregue',     pesoTotalEntregue,          'soma de dificuldade das Closed'],
  ];
  clRows.forEach(([m,v,o], ri) => {
    const r = rowCl+2+ri;
    ws.setRowHeight(r, 18);
    const alt = ri%2===1?'#EBF5EB':'#FFFFFF';
    styleRange(ws.getRange(r,1), { value:m, bg:alt, fg:'#000000', bold:true,  sz:10, align:'left'   });
    styleRange(ws.getRange(r,2), { value:v, bg:alt, fg:'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(r,3), { value:o, bg:alt, fg:'#595959', bold:false, sz:9,  align:'left', italic:true });
    ws.getRange(r,4,1,8).setBackground(alt);
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// 2. POR PESSOA
// ═══════════════════════════════════════════════════════════════════════════════
function gerarPorPessoaM(ss, dados) {
  const ws    = getOrCreateSheet(ss, '👥 Por Pessoa', '#1F3864');
  const hoje  = new Date(); hoje.setHours(0,0,0,0);
  const dataStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const total = dados.length;

  ws.setRowHeight(1, 30);
  // Set NCOLS later based on content; start with minimum
  const NCOLS_BASE = 9;

  // Agrega por pessoa com dificuldade
  const pessoaMap = {};
  dados.forEach(d => {
    splitPessoasM(d.pessoa).forEach(p => {
      if (!pessoaMap[p]) pessoaMap[p] = {
        total:0, fechadas:0, abertas:0, andamento:0,
        horasEst:0, horasReal:0, tipos:{}, difs:{}, indicePonderado:0,
      };
      const pm = pessoaMap[p];
      pm.total++;
      if (d.status === 'Closed') {
        pm.fechadas++;
        const dif = d.dificuldade || '(não definida)';
        if (!pm.difs[dif]) pm.difs[dif] = { cnt:0, horasEst:0, horasReal:0, pesoSum:0 };
        pm.difs[dif].cnt++;
        pm.difs[dif].horasEst  += d.horasEst;
        pm.difs[dif].horasReal += d.horasReal;
        pm.difs[dif].pesoSum   += d.peso || 0;
        // Índice ponderado: peso × horasEst — crédito por tarefas difíceis E longas
        pm.indicePonderado += (d.peso > 0 ? d.peso : 1) * Math.max(d.horasEst, 1);
      } else {
        pm.abertas++;
      }
      if (d.status === 'Em andamento') pm.andamento++;
      pm.horasEst  += d.horasEst;
      pm.horasReal += d.horasReal;
      const t = d.tipo || '(sem tipo)';
      pm.tipos[t] = (pm.tipos[t]||0)+1;
    });
  });

  const pessoas = Object.keys(pessoaMap).sort((a,b) => pessoaMap[b].total - pessoaMap[a].total);

  // Expand columns
  const maxTipos = Math.max(...pessoas.map(p => Object.keys(pessoaMap[p].tipos).length), 0);
  const NCOLS = Math.max(NCOLS_BASE, 1 + maxTipos + 2);
  if (ws.getMaxColumns() < NCOLS) ws.insertColumnsAfter(ws.getMaxColumns(), NCOLS - ws.getMaxColumns());

  styleRange(ws.getRange(1,1,1,NCOLS).merge(), {
    value: `👥  MÉTRICAS POR PESSOA  —  ${dataStr}`,
    bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 13, align: 'left',
  });
  ws.setRowHeight(2, 8);
  ws.getRange(2,1,1,NCOLS).setBackground('#1F3864');

  // ── PERFORMANCE INDIVIDUAL
  let row = 3;
  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
    value: '  PERFORMANCE INDIVIDUAL', bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  row++;

  ws.setRowHeight(row, 18);
  const hdrs = ['Responsável','Total','Fechadas','Abertas','Em andamento','% Conclusão','H. Estimadas','H. Rastreadas','H. Méd/Tarefa'];
  const wdts = [200, 60, 70, 70, 100, 90, 100, 110, 110];
  hdrs.forEach((h,i) => { ws.setColumnWidth(i+1, wdts[i]); styleRange(ws.getRange(row,i+1), { value:h, bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:9, align:'center' }); });
  row++;

  pessoas.forEach((p, pi) => {
    const pm  = pessoaMap[p];
    const alt = pi%2===1 ? '#F0F4FA' : '#FFFFFF';
    const pct = pm.total > 0 ? `${Math.round(pm.fechadas/pm.total*100)}%` : '0%';
    const med = pm.total > 0 ? Math.round(pm.horasEst / pm.total * 10) / 10 : 0;
    ws.setRowHeight(row, 20);
    [p, pm.total, pm.fechadas, pm.abertas, pm.andamento, pct, pm.horasEst, pm.horasReal, med]
      .forEach((v, ci) => styleRange(ws.getRange(row, ci+1), { value:v, bg:ci===0?'#1F3864':alt, fg:ci===0?'#FFFFFF':'#000000', bold:ci===0, sz:10, align:ci===0?'left':'center' }));
    row++;
  });

  row++;
  ws.setRowHeight(row - 1, 12);

  // ── MIX DE TIPO POR PESSOA
  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
    value: '  MIX DE TIPO POR PESSOA', bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  row++;

  const tiposSet = new Set();
  dados.forEach(d => tiposSet.add(d.tipo || '(sem tipo)'));
  const tipos = [...tiposSet].sort();
  const tipoNCOLS = 1 + tipos.length;
  if (ws.getMaxColumns() < tipoNCOLS) ws.insertColumnsAfter(ws.getMaxColumns(), tipoNCOLS - ws.getMaxColumns());

  ws.setRowHeight(row, 18);
  styleRange(ws.getRange(row,1), { value:'Responsável', bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:9, align:'center' });
  tipos.forEach((t,i) => { ws.setColumnWidth(i+2, 90); styleRange(ws.getRange(row,i+2), { value:t, bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:9, align:'center' }); });
  row++;

  pessoas.forEach((p, pi) => {
    const pm  = pessoaMap[p];
    const alt = pi%2===1 ? '#F0F4FA' : '#FFFFFF';
    ws.setRowHeight(row, 20);
    styleRange(ws.getRange(row,1), { value:p, bg:'#1F3864', fg:'#FFFFFF', bold:true, sz:10, align:'left' });
    tipos.forEach((t,i) => {
      const cnt = pm.tipos[t] || 0;
      styleRange(ws.getRange(row,i+2), { value:cnt>0?cnt:'—', bg:cnt>0?alt:'#F9F9F9', fg:cnt>0?'#000000':'#CCCCCC', bold:false, sz:10, align:'center' });
    });
    row++;
  });

  row++;
  ws.setRowHeight(row-1, 12);

  // ── PERFORMANCE POR DIFICULDADE (apenas Closed)
  const DIFICULDADE_ORD = ['Fácil','Média fácil','Média','Média difícil','Difícil','(não definida)'];

  // Collect all difficulty levels that appear
  const difsExist = new Set();
  pessoas.forEach(p => Object.keys(pessoaMap[p].difs).forEach(d => difsExist.add(d)));
  const difsOrdered = DIFICULDADE_ORD.filter(d => difsExist.has(d));
  const difsExtra   = [...difsExist].filter(d => !DIFICULDADE_ORD.includes(d));
  const difs        = [...difsOrdered, ...difsExtra];

  const difNCOLS = 2 + difs.length * 2 + 1; // nome | índice | (cnt+h per dif) | total h
  if (ws.getMaxColumns() < difNCOLS) ws.insertColumnsAfter(ws.getMaxColumns(), difNCOLS - ws.getMaxColumns());

  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,difNCOLS).merge(), {
    value: '  PERFORMANCE POR DIFICULDADE — tarefas Closed',
    bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  row++;

  ws.setRowHeight(row, 14);
  styleRange(ws.getRange(row,1,1,difNCOLS).merge(), {
    value: '  Índice ponderado = Σ(peso × H.Estimadas) para Closed — pondera tarefas difíceis e longas',
    bg: '#DEEAF1', fg: '#2E75B6', bold: false, sz: 9, align: 'left', italic: true,
  });
  row++;

  ws.setRowHeight(row, 18);
  ws.setColumnWidth(1, 200); ws.setColumnWidth(2, 100);
  styleRange(ws.getRange(row,1), { value:'Responsável',    bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:9, align:'center' });
  styleRange(ws.getRange(row,2), { value:'Índice Ponderado',bg:'#0B5394', fg:'#FFFFFF', bold:true, sz:9, align:'center' });
  difs.forEach((dif, di) => {
    const col = 3 + di * 2;
    ws.setColumnWidth(col,   70);
    ws.setColumnWidth(col+1, 80);
    styleRange(ws.getRange(row,col),   { value:`${dif} (qtd)`, bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:8, align:'center' });
    styleRange(ws.getRange(row,col+1), { value:`H.Est`,         bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:8, align:'center' });
  });
  row++;

  // Sort pessoas by índice ponderado desc for this section
  const pessoasPorIndice = [...pessoas].sort((a,b) => pessoaMap[b].indicePonderado - pessoaMap[a].indicePonderado);

  pessoasPorIndice.forEach((p, pi) => {
    const pm  = pessoaMap[p];
    const alt = pi%2===1 ? '#F0F4FA' : '#FFFFFF';
    ws.setRowHeight(row, 20);
    const ip = Math.round(pm.indicePonderado * 10) / 10;
    styleRange(ws.getRange(row,1), { value:p,  bg:'#1F3864', fg:'#FFFFFF', bold:true,  sz:10, align:'left'   });
    styleRange(ws.getRange(row,2), { value:ip, bg:'#0B5394', fg:'#FFFFFF', bold:true,  sz:10, align:'center' });
    difs.forEach((dif, di) => {
      const col  = 3 + di * 2;
      const dd   = pm.difs[dif] || { cnt:0, horasEst:0 };
      const hasD = dd.cnt > 0;
      styleRange(ws.getRange(row,col),   { value:hasD?dd.cnt:'—',      bg:hasD?alt:'#F9F9F9', fg:hasD?'#000000':'#CCCCCC', bold:false, sz:10, align:'center' });
      styleRange(ws.getRange(row,col+1), { value:hasD?dd.horasEst:'—', bg:hasD?alt:'#F9F9F9', fg:hasD?'#000000':'#CCCCCC', bold:false, sz:10, align:'center' });
    });
    row++;
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// 3. POR CLIENTE
// ═══════════════════════════════════════════════════════════════════════════════
function gerarPorClienteM(ss, dados) {
  const ws = getOrCreateSheet(ss, '🏢 Por Cliente', '#1F3864');
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const dataStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const NCOLS = 7;
  if (ws.getMaxColumns() < NCOLS) ws.insertColumnsAfter(ws.getMaxColumns(), NCOLS - ws.getMaxColumns());

  ws.setRowHeight(1, 30);
  styleRange(ws.getRange(1,1,1,NCOLS).merge(), {
    value: `🏢  MÉTRICAS POR CLIENTE / PROJETO  —  ${dataStr}`,
    bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 13, align: 'left',
  });
  ws.setRowHeight(2, 8);
  ws.getRange(2,1,1,NCOLS).setBackground('#1F3864');

  // Agrega por cliente
  const clienteMap = {};
  dados.forEach(d => {
    const key = d.projeto ? `${d.cliente} › ${d.projeto}` : d.cliente;
    if (!clienteMap[key]) clienteMap[key] = { total:0, fechadas:0, abertas:0, horasEst:0, horasReal:0 };
    const cm = clienteMap[key];
    cm.total++;
    if (d.status === 'Closed') cm.fechadas++; else cm.abertas++;
    cm.horasEst  += d.horasEst;
    cm.horasReal += d.horasReal;
  });
  const clientes = Object.keys(clienteMap).sort((a,b) => clienteMap[b].total - clienteMap[a].total);

  // ── DEMANDAS E HORAS POR CLIENTE
  let row = 3;
  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
    value: '  DEMANDAS E HORAS POR CLIENTE', bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  row++;

  ws.setRowHeight(row, 18);
  const hdrs = ['Cliente / Projeto','Total','Fechadas','Abertas','% Conclusão','H. Estimadas','H. Rastreadas'];
  const wdts = [220, 60, 70, 70, 90, 100, 110];
  hdrs.forEach((h,i) => { ws.setColumnWidth(i+1, wdts[i]); styleRange(ws.getRange(row,i+1), { value:h, bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:9, align:'center' }); });
  row++;

  clientes.forEach((cl, ci) => {
    const cm  = clienteMap[cl];
    const alt = ci%2===1 ? '#F0F4FA' : '#FFFFFF';
    const pct = cm.total > 0 ? `${Math.round(cm.fechadas/cm.total*100)}%` : '0%';
    ws.setRowHeight(row, 20);
    [cl, cm.total, cm.fechadas, cm.abertas, pct, cm.horasEst, cm.horasReal]
      .forEach((v, vi) => styleRange(ws.getRange(row,vi+1), { value:v, bg:vi===0?'#1F3864':alt, fg:vi===0?'#FFFFFF':'#000000', bold:vi===0, sz:10, align:vi===0?'left':'center' }));
    row++;
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// 4. ALERTAS
// ═══════════════════════════════════════════════════════════════════════════════
function gerarAlertasM(ss, dados) {
  const ws = getOrCreateSheet(ss, '⚠️ Alertas', '#C00000');
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const dataStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const total = dados.length;
  const NCOLS = 7;
  if (ws.getMaxColumns() < NCOLS) ws.insertColumnsAfter(ws.getMaxColumns(), NCOLS - ws.getMaxColumns());

  ws.setRowHeight(1, 30);
  styleRange(ws.getRange(1,1,1,NCOLS).merge(), {
    value: `⚠️  ALERTAS E QUALIDADE DE DADOS  —  ${dataStr}`,
    bg: '#C00000', fg: '#FFFFFF', bold: true, sz: 13, align: 'left',
  });
  ws.setRowHeight(2, 8);
  ws.getRange(2,1,1,NCOLS).setBackground('#C00000');

  // ── QUALIDADE DOS DADOS
  let row = 3;
  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
    value: '  QUALIDADE DOS DADOS', bg: '#7F0000', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  row++;

  ws.setRowHeight(row, 18);
  const qHdrs = ['Problema','Qtd tarefas','% do total','Ação recomendada'];
  const qWdts = [200, 90, 90, 250];
  qHdrs.forEach((h,i) => { ws.setColumnWidth(i+1, qWdts[i]); styleRange(ws.getRange(row,i+1), { value:h, bg:'#C00000', fg:'#FFFFFF', bold:true, sz:9, align:'center' }); });
  row++;

  const qualidade = [
    ['Sem responsável',      dados.filter(d => d.pessoa==='Sem responsável').length,    'Atribuir dono'],
    ['Sem tipo de tarefa',   dados.filter(d => !d.tipo).length,                         'Categorizar'],
    ['Sem tempo estimado',   dados.filter(d => d.horasEst===0).length,                  'Estimar horas'],
    ['Sem rastreamento',     dados.filter(d => d.horasReal===0).length,                 'Registrar horas'],
    ['Sem prioridade',       dados.filter(d => !d.prioridade).length,                   'Definir prioridade'],
    ['Sem cliente vinculado',dados.filter(d => d.cliente==='Sem cliente').length,        'Vincular cliente'],
    ['Sem data final',       dados.filter(d => !d.dataFinal).length,                    'Definir prazo'],
  ];
  qualidade.forEach(([prob, cnt, acao], ri) => {
    const r = row + ri;
    ws.setRowHeight(r, 20);
    const alt = ri%2===1 ? '#FDE9D9' : '#FFFFFF';
    const emph = cnt > 0;
    styleRange(ws.getRange(r,1), { value:prob, bg:emph?'#FDECEA':alt, fg:'#000000', bold:emph, sz:10, align:'left'   });
    styleRange(ws.getRange(r,2), { value:cnt,  bg:emph?'#FDECEA':alt, fg:emph?'#C00000':'#000000', bold:emph, sz:10, align:'center' });
    styleRange(ws.getRange(r,3), { value:pctM(cnt,total), bg:emph?'#FDECEA':alt, fg:emph?'#C00000':'#595959', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(r,4), { value:acao, bg:alt, fg:'#595959', bold:false, sz:9, align:'left', italic:true });
  });
  row += qualidade.length + 1;

  // ── TAREFAS ATRASADAS
  ws.setColumnWidth(1, 260); ws.setColumnWidth(2, 140); ws.setColumnWidth(3, 100);
  ws.setColumnWidth(4, 200); ws.setColumnWidth(5, 90);  ws.setColumnWidth(6, 90); ws.setColumnWidth(7, 130);

  const atrasadas = dados.filter(d => d.atrasado).sort((a,b) => (a.dataFinal||0) - (b.dataFinal||0));
  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
    value: `  TAREFAS ATRASADAS  (${atrasadas.length})`, bg: '#7F0000', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  row++;

  const STATUS_BG_A = {
    'Closed':'#D9EAD3','Em andamento':'#DEEAF1','Em revisão':'#EAD1F7','Em testes':'#E6F3FF',
    'A fazer':'#F4F4F4','Backlog':'#EEEEEE','Aguardando':'#FFF3E0',
    'Aguardando cliente':'#FFF3E0','Aguardando TA':'#FFF3E0','Ajustes de revisão':'#FDE9D9',
  };

  if (atrasadas.length === 0) {
    ws.setRowHeight(row, 22);
    styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
      value: '✓ Nenhuma tarefa atrasada.', bg: '#D9EAD3', fg: '#38761D', bold: false, sz: 11, align: 'center',
    });
    row++;
  } else {
    ws.setRowHeight(row, 18);
    ['Tarefa','Status','Prioridade','Responsável','Vencimento','Dias atraso','Urgência']
      .forEach((h,i) => styleRange(ws.getRange(row,i+1), { value:h, bg:'#C00000', fg:'#FFFFFF', bold:true, sz:9, align:'center' }));
    row++;

    atrasadas.forEach((d, ri) => {
      const r    = row + ri;
      const alt  = ri%2===1 ? '#FDE9D9' : '#FFFFFF';
      const dias = d.dataFinal ? Math.round((hoje - d.dataFinal) / 86400000) : 0;
      const sbg  = STATUS_BG_A[d.status] || alt;
      const urgBg = dias > 30 ? '#C00000' : dias > 14 ? '#E65100' : '#BF8F00';
      const urgTxt = dias > 30 ? '🔴 Crítico' : dias > 14 ? '🟠 Alto' : '🟡 Médio';
      ws.setRowHeight(r, 20);
      const nome = d.nome.length > 60 ? d.nome.substring(0,57)+'...' : d.nome;
      styleRange(ws.getRange(r,1), { value:nome,              bg:alt,  fg:'#000000', bold:false, sz:10, align:'left'   });
      styleRange(ws.getRange(r,2), { value:d.status,          bg:sbg,  fg:'#000000', bold:false, sz:10, align:'center' });
      styleRange(ws.getRange(r,3), { value:d.prioridade||'—', bg:alt,  fg:'#000000', bold:false, sz:10, align:'center' });
      styleRange(ws.getRange(r,4), { value:d.pessoa,          bg:alt,  fg:'#000000', bold:false, sz:10, align:'left'   });
      const dtCell = ws.getRange(r,5);
      styleRange(dtCell, { value:d.dataFinal||'—', bg:alt, fg:'#000000', bold:false, sz:10, align:'center' });
      if (d.dataFinal) dtCell.setNumberFormat('DD/MM/YYYY');
      styleRange(ws.getRange(r,6), { value:dias,    bg:'#FDECEA', fg:'#C00000', bold:true, sz:10, align:'center' });
      styleRange(ws.getRange(r,7), { value:urgTxt,  bg:urgBg,     fg:'#FFFFFF', bold:true, sz:9,  align:'center' });
    });
    row += atrasadas.length;
  }

  row++;
  ws.setRowHeight(row-1, 12);

  // ── TAREFAS BLOQUEADAS (AGUARDANDO)
  const aguardTodas = dados.filter(d => d.status.startsWith('Aguardando'));

  // Enriquece com urgência por data_final
  const comUrgencia = aguardTodas.map(d => {
    const diasVencida = d.dataFinal ? Math.round((hoje - d.dataFinal) / 86400000) : null;
    const diasRestantes = d.dataFinal ? Math.round((d.dataFinal - hoje) / 86400000) : null;
    let urgOrder, situacao, sitBg, sitFg;
    if (!d.dataFinal) {
      urgOrder=3; situacao='Sem prazo definido'; sitBg='#F4F4F4'; sitFg='#757575';
    } else if (diasVencida > 0) {
      urgOrder=0; situacao=`Vencida há ${diasVencida}d`; sitBg='#FDECEA'; sitFg='#C00000';
    } else if (diasRestantes <= 7) {
      urgOrder=1; situacao=`Vence em ${diasRestantes}d`; sitBg='#FFF2CC'; sitFg='#BF8F00';
    } else if (diasRestantes <= 30) {
      urgOrder=2; situacao=`Vence em ${diasRestantes}d`; sitBg='#FFF8E1'; sitFg='#E65100';
    } else {
      urgOrder=4; situacao=`Vence em ${diasRestantes}d`; sitBg='#E8F5E9'; sitFg='#38761D';
    }
    return { ...d, urgOrder, diasVencida, diasRestantes, situacao, sitBg, sitFg };
  }).sort((a,b) => a.urgOrder - b.urgOrder || (a.diasVencida||0) - (b.diasVencida||0));

  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
    value: `  TAREFAS BLOQUEADAS — AGUARDANDO  (${aguardTodas.length})`,
    bg: '#BF8F00', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  row++;

  if (aguardTodas.length === 0) {
    ws.setRowHeight(row, 22);
    styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
      value: '✓ Nenhuma tarefa aguardando.', bg: '#D9EAD3', fg: '#38761D', bold: false, sz: 11, align: 'center',
    });
    return;
  }

  ws.setRowHeight(row, 18);
  ['Tarefa','Sub-status','Prioridade','Responsável','Data Final','Situação','Ação']
    .forEach((h,i) => styleRange(ws.getRange(row,i+1), { value:h, bg:'#BF8F00', fg:'#FFFFFF', bold:true, sz:9, align:'center' }));
  row++;

  comUrgencia.forEach((d, ri) => {
    const r   = row + ri;
    const alt = ri%2===1 ? '#FFF9E6' : '#FFFFFF';
    const sbg = STATUS_BG_A[d.status] || alt;
    ws.setRowHeight(r, 20);
    const nome = d.nome.length > 60 ? d.nome.substring(0,57)+'...' : d.nome;
    const acao = d.urgOrder===0 ? 'Desbloquear urgente' : d.urgOrder===1 ? 'Verificar bloqueio' : d.urgOrder===3 ? 'Definir prazo' : 'Monitorar';
    styleRange(ws.getRange(r,1), { value:nome,              bg:alt,     fg:'#000000', bold:false, sz:10, align:'left'   });
    styleRange(ws.getRange(r,2), { value:d.status,          bg:sbg,     fg:'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(r,3), { value:d.prioridade||'—', bg:alt,     fg:'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(r,4), { value:d.pessoa,          bg:alt,     fg:'#000000', bold:false, sz:10, align:'left'   });
    const dtCell = ws.getRange(r,5);
    styleRange(dtCell, { value:d.dataFinal||'—', bg:alt, fg:'#000000', bold:false, sz:10, align:'center' });
    if (d.dataFinal) dtCell.setNumberFormat('DD/MM/YYYY');
    styleRange(ws.getRange(r,6), { value:d.situacao, bg:d.sitBg, fg:d.sitFg, bold:d.urgOrder<=1, sz:10, align:'center' });
    styleRange(ws.getRange(r,7), { value:acao, bg:alt, fg:'#595959', bold:false, sz:9, align:'left', italic:true });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// 5. HORAS POR SEMANA
// ═══════════════════════════════════════════════════════════════════════════════
function gerarHorasSemanaM(ss, dados) {
  const ws    = getOrCreateSheet(ss, '📅 Horas por Semana', '#0B5394');
  const hoje  = new Date(); hoje.setHours(0, 0, 0, 0);
  const N     = 8; // semanas futuras exibidas (sem contar a semana de referência)

  // Semana de referência = a semana anterior completa (Seg–Sex)
  const monAtual    = weekMonday(hoje);
  const monRef      = new Date(monAtual); monRef.setDate(monRef.getDate() - 7);
  const sexRef      = new Date(monRef);   sexRef.setDate(sexRef.getDate() + 4);
  const refLabel    = `↩ ${weekLabelM(monRef)} (ref.)`;

  // Semanas exibidas: [semana_ref, semana_atual, +1, +2, ... +N-1]
  const semanas = Array.from({length: N + 1}, (_, i) => {
    const m = new Date(monRef); m.setDate(m.getDate() + i * 7); return m;
  });
  // semanas[0] = semana anterior (referência), semanas[1] = semana atual, etc.

  // Pessoas únicas (excluindo "Sem responsável")
  const pessoasSet = new Set();
  dados.forEach(d => splitPessoasM(d.pessoa).forEach(p => pessoasSet.add(p)));
  const pessoas = [...pessoasSet].sort();

  // Monta matriz horas[pessoa][semana_idx] para todas as semanas exibidas
  const horas = {};
  pessoas.forEach(p => horas[p] = Array(N + 1).fill(0));

  dados
    .filter(d => d.status !== 'Closed' && d.dataFinal)
    .forEach(d => {
      const wk  = weekMonday(d.dataFinal);
      const idx = semanas.findIndex(s => s.getTime() === wk.getTime());
      if (idx < 0) return;
      splitPessoasM(d.pessoa).forEach(p => {
        if (horas[p]) horas[p][idx] += d.horasEst;
      });
    });

  const NCOLS = 2 + (N + 1) + 1; // nome | total | (N+1) semanas | situação
  if (ws.getMaxColumns() < NCOLS) ws.insertColumnsAfter(ws.getMaxColumns(), NCOLS - ws.getMaxColumns());

  const dataStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  // Row 1: título
  ws.setRowHeight(1, 30);
  styleRange(ws.getRange(1,1,1,NCOLS).merge(), {
    value: `📅  HORAS POR SEMANA / COLABORADOR  —  Capacidade: ${CAP_SEMANA}h  |  Mínimo: ${MIN_SEMANA}h  |  ${dataStr}`,
    bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 12, align: 'left',
  });

  // Row 2: spacer
  ws.setRowHeight(2, 8);
  ws.getRange(2,1,1,NCOLS).setBackground('#1F3864');

  // Row 3: legenda
  ws.setRowHeight(3, 18);
  styleRange(ws.getRange(3,1,1,NCOLS).merge(), {
    value: '  🟩 ≥44h (capacidade plena)    🟨 36–43h (mínimo ok)    🟥 <36h (precisa de tarefas)    ⬜ 0h (sem tarefas agendadas)',
    bg: '#F8F9FA', fg: '#444444', bold: false, sz: 9, align: 'left',
  });

  // Row 4: spacer
  ws.setRowHeight(4, 8);

  // Row 5: section header
  ws.setRowHeight(5, 22);
  styleRange(ws.getRange(5,1,1,NCOLS).merge(), {
    value: '  PLANEJAMENTO SEMANAL — horas estimadas por colaborador (tarefas abertas, agrupadas por data final)',
    bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });

  // Row 6: cabeçalhos
  ws.setRowHeight(6, 22);
  ws.setColumnWidth(1, 190);
  ws.setColumnWidth(2, 80);
  styleRange(ws.getRange(6,1), { value:'Colaborador',       bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:9, align:'center' });
  styleRange(ws.getRange(6,2), { value:`Total (${N+1} sem.)`, bg:'#2E75B6', fg:'#FFFFFF', bold:true, sz:9, align:'center' });
  semanas.forEach((s,i) => {
    ws.setColumnWidth(3+i, 88);
    // Semana de referência (índice 0) recebe destaque diferente
    const isRef = i === 0;
    styleRange(ws.getRange(6,3+i), {
      value: isRef ? refLabel : weekLabelM(s),
      bg: isRef ? '#0B5394' : '#2E75B6',
      fg: '#FFFFFF', bold: true, sz: 8, align: 'center',
    });
  });
  ws.setColumnWidth(3+(N+1), 150);
  styleRange(ws.getRange(6,3+(N+1)), { value:'Situação (sem. anterior)', bg:'#0B5394', fg:'#FFFFFF', bold:true, sz:8, align:'center' });

  // Rows 7+: por pessoa
  pessoas.forEach((pessoa, pi) => {
    const row  = 7 + pi;
    const hrs  = horas[pessoa];
    const tot  = hrs.reduce((s,h) => s+h, 0);
    const altBg = pi%2===1 ? '#F0F4FA' : '#FFFFFF';

    // Situação baseada SOMENTE na semana anterior (índice 0)
    const hRef = hrs[0];
    let alertTxt, alertBg, alertFg;
    if      (hRef === 0)        { alertTxt = '⬜ Sem tarefas agendadas'; alertBg = '#FDECEA'; alertFg = '#C00000'; }
    else if (hRef < MIN_SEMANA) { alertTxt = `⚠ ${hRef}h — abaixo de ${MIN_SEMANA}h`; alertBg = '#FFF2CC'; alertFg = '#BF8F00'; }
    else if (hRef < CAP_SEMANA) { alertTxt = `✓ ${hRef}h — mínimo ok`; alertBg = '#D9EAD3'; alertFg = '#38761D'; }
    else                        { alertTxt = `✓ ${hRef}h — capacidade plena`; alertBg = '#B7E1CD'; alertFg = '#0D652D'; }

    ws.setRowHeight(row, 22);
    styleRange(ws.getRange(row,1), { value:pessoa, bg:'#1F3864', fg:'#FFFFFF', bold:true,  sz:10, align:'left'   });
    styleRange(ws.getRange(row,2), { value:tot,    bg:altBg,    fg:'#000000', bold:false, sz:10, align:'center' });

    hrs.forEach((h, i) => {
      const isRef = i === 0;
      let cellBg, cellFg;
      if      (h === 0)        { cellBg = isRef ? '#FDECEA' : '#F8F8F8'; cellFg = isRef ? '#C00000' : '#BBBBBB'; }
      else if (h < MIN_SEMANA) { cellBg='#FFF2CC'; cellFg='#BF8F00'; }
      else if (h < CAP_SEMANA) { cellBg='#D9EAD3'; cellFg='#38761D'; }
      else                     { cellBg='#B7E1CD'; cellFg='#0D652D'; }
      styleRange(ws.getRange(row, 3+i), {
        value: h > 0 ? h : '—',
        bg: cellBg, fg: cellFg,
        bold: isRef || (h > 0 && h < MIN_SEMANA),
        sz: 10, align: 'center',
      });
    });

    styleRange(ws.getRange(row, 3+(N+1)), { value:alertTxt, bg:alertBg, fg:alertFg, bold:true, sz:9, align:'left' });
  });

  // Linha de total por semana
  const sumRow = 7 + pessoas.length;
  ws.setRowHeight(sumRow, 22);
  styleRange(ws.getRange(sumRow,1), { value:'TOTAL POR SEMANA', bg:'#1F3864', fg:'#FFFFFF', bold:true,  sz:10, align:'left'   });
  styleRange(ws.getRange(sumRow,2), { value:'',                 bg:'#1F3864', fg:'#FFFFFF', bold:false, sz:10, align:'center' });
  semanas.forEach((_,i) => {
    const totSem = pessoas.reduce((acc,p) => acc + horas[p][i], 0);
    const isRef  = i === 0;
    styleRange(ws.getRange(sumRow, 3+i), { value:totSem, bg: isRef ? '#0B5394' : '#2E75B6', fg:'#FFFFFF', bold:true, sz:10, align:'center' });
  });
  styleRange(ws.getRange(sumRow, 3+(N+1)), { value:'', bg:'#0B5394', fg:'#FFFFFF', bold:false, sz:10, align:'center' });
}

// ═══════════════════════════════════════════════════════════════════════════════
// 6. CONCLUÍDAS — ANÁLISE DE QUALIDADE E DIFICULDADE
// ═══════════════════════════════════════════════════════════════════════════════
function gerarConcluidasM(ss, dados) {
  const ws    = getOrCreateSheet(ss, '✅ Concluídas', '#38761D');
  const hoje  = new Date(); hoje.setHours(0,0,0,0);
  const dataStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const NCOLS = 12;
  if (ws.getMaxColumns() < NCOLS) ws.insertColumnsAfter(ws.getMaxColumns(), NCOLS - ws.getMaxColumns());

  const total   = dados.length;
  const closed  = dados.filter(d => d.status === 'Closed');
  const nClosed = closed.length;

  // ── Row 1: título
  ws.setRowHeight(1, 30);
  styleRange(ws.getRange(1,1,1,NCOLS).merge(), {
    value: `✅  CONCLUÍDAS — ANÁLISE DE QUALIDADE & DIFICULDADE  —  ${dataStr}`,
    bg: '#38761D', fg: '#FFFFFF', bold: true, sz: 13, align: 'left',
  });
  ws.setRowHeight(2, 8);
  ws.getRange(2,1,1,NCOLS).setBackground('#38761D');

  // ── KPIs
  ws.setRowHeight(3, 20);
  styleRange(ws.getRange(3,1,1,NCOLS).merge(), {
    value: '  VISÃO GERAL DAS ENTREGAS', bg: '#274E13', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });

  const closedComH    = closed.filter(d => d.horasEst>0 && d.horasReal>0);
  const closedComReal = closed.filter(d => d.horasReal>0);
  const closedOK      = closedComH.filter(d => Math.abs((d.horasReal-d.horasEst)/d.horasEst)<=0.10);
  const desvioMClosed = closedComH.length
    ? closedComH.reduce((s,d) => s+(d.horasReal-d.horasEst)/d.horasEst,0)/closedComH.length : 0;
  const pesoTotal     = closed.reduce((s,d) => s+(d.peso||0), 0);
  const hRealTotal    = closedComReal.reduce((s,d) => s+d.horasReal, 0);

  const kpis = [
    ['Total Closed',        nClosed,                                                   '#38761D'],
    ['% do total',          pctM(nClosed,total),                                       '#0B5394'],
    ['H. Rastreadas',       hRealTotal,                                                '#2E75B6'],
    ['Taxa rastreamento',   pctM(closedComReal.length, nClosed),                       nClosed>0&&closedComReal.length/nClosed<0.5?'#C00000':'#38761D'],
    ['Precisão ±10%',       pctM(closedOK.length, closedComH.length||1),              closedComH.length>0&&closedOK.length/closedComH.length<0.4?'#BF8F00':'#38761D'],
    ['Peso total entregue', pesoTotal,                                                  '#1F3864'],
  ];
  const kpiW = [120, 110, 110, 120, 110, 130];
  kpis.forEach(([lbl,val,color],ki) => {
    ws.setColumnWidth(ki+1, kpiW[ki]);
    styleRange(ws.getRange(4, ki+1), { value:lbl, bg:color,     fg:'#FFFFFF', bold:true,  sz:9,  align:'center' });
    styleRange(ws.getRange(5, ki+1), { value:val, bg:'#FFFFFF', fg:color,     bold:true,  sz:20, align:'center' });
  });
  ws.setRowHeight(4, 18); ws.setRowHeight(5, 38);
  ws.getRange(4,7,2,NCOLS-6).setBackground('#FFFFFF');
  ws.setRowHeight(6, 14);

  let row = 7;

  // ── TENDÊNCIA MENSAL (últimos 6 meses por data_final)
  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
    value: '  TENDÊNCIA MENSAL DE ENTREGAS', bg: '#274E13', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  row++;

  ws.setRowHeight(row, 18);
  ws.setColumnWidth(1,120); ws.setColumnWidth(2,70); ws.setColumnWidth(3,100);
  ws.setColumnWidth(4,100); ws.setColumnWidth(5,100); ws.setColumnWidth(6,120);
  ['Mês','Qtd','H. Rastreadas','H. Estimadas','Peso entregue','Desvio médio'].forEach((h,i) =>
    styleRange(ws.getRange(row,i+1), { value:h, bg:'#38761D', fg:'#FFFFFF', bold:true, sz:9, align:'center' }));
  row++;

  const mesesMap = {};
  closed.forEach(d => {
    if (!d.dataFinal) return;
    const key = `${d.dataFinal.getFullYear()}-${String(d.dataFinal.getMonth()+1).padStart(2,'0')}`;
    if (!mesesMap[key]) mesesMap[key] = { cnt:0, hReal:0, hEst:0, peso:0, deltas:[] };
    const m = mesesMap[key];
    m.cnt++; m.hReal+=d.horasReal; m.hEst+=d.horasEst; m.peso+=(d.peso||0);
    if (d.horasEst>0&&d.horasReal>0) m.deltas.push((d.horasReal-d.horasEst)/d.horasEst);
  });

  const MESES_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const mesesKeys = Object.keys(mesesMap).sort().reverse().slice(0,6).reverse();
  if (!mesesKeys.length) {
    ws.setRowHeight(row, 22);
    styleRange(ws.getRange(row,1,1,NCOLS).merge(), { value:'Sem tarefas Closed com data final definida.', bg:'#F4F4F4', fg:'#595959', bold:false, sz:10, align:'center' });
    row++;
  } else {
    mesesKeys.forEach((key, mi) => {
      const m   = mesesMap[key];
      const alt = mi%2===1 ? '#EBF5EB' : '#FFFFFF';
      const [yr, mo] = key.split('-');
      const mLabel   = `${MESES_PT[parseInt(mo)-1]}/${yr}`;
      const desvio   = m.deltas.length ? `${(m.deltas.reduce((s,v)=>s+v,0)/m.deltas.length*100).toFixed(1)}%` : '—';
      ws.setRowHeight(row, 20);
      [mLabel, m.cnt, m.hReal, m.hEst, m.peso, desvio].forEach((v,vi) =>
        styleRange(ws.getRange(row,vi+1), { value:v, bg:vi===0?'#1F3864':alt, fg:vi===0?'#FFFFFF':'#000000', bold:vi===0, sz:10, align:vi===0?'left':'center' }));
      row++;
    });
  }

  row++;
  ws.setRowHeight(row-1, 12);

  // ── ENTREGAS POR DIFICULDADE — EQUIPE
  const DIFICULDADE_ORD_C = ['Fácil','Média fácil','Média','Média difícil','Difícil'];
  const difMapC = {};
  closed.forEach(d => {
    const df = d.dificuldade || '(não definida)';
    if (!difMapC[df]) difMapC[df] = { cnt:0, hEst:0, hReal:0, pesoSum:0, deltas:[] };
    const dm = difMapC[df];
    dm.cnt++; dm.hEst+=d.horasEst; dm.hReal+=d.horasReal; dm.pesoSum+=(d.peso||0);
    if (d.horasEst>0&&d.horasReal>0) dm.deltas.push((d.horasReal-d.horasEst)/d.horasEst);
  });
  const difKeys = [
    ...DIFICULDADE_ORD_C.filter(d => difMapC[d]),
    ...Object.keys(difMapC).filter(d => !DIFICULDADE_ORD_C.includes(d)),
  ];

  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
    value: '  ENTREGAS POR DIFICULDADE — EQUIPE',
    bg: '#274E13', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  row++;

  ws.setRowHeight(row, 14);
  styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
    value: '  Peso total = soma de "Dificuldade (Peso)" das Closed — indica esforço real entregue por categoria',
    bg: '#EBF5EB', fg: '#38761D', bold: false, sz: 9, align: 'left', italic: true,
  });
  row++;

  ws.setRowHeight(row, 18);
  ws.setColumnWidth(1,170); ws.setColumnWidth(2,70); ws.setColumnWidth(3,70);
  ws.setColumnWidth(4,90);  ws.setColumnWidth(5,90); ws.setColumnWidth(6,100);
  ws.setColumnWidth(7,110); ws.setColumnWidth(8,110);
  ['Dificuldade','Qtd','% Closed','H. Est.','H. Reais','Peso total','H. méd/tarefa','Desvio médio'].forEach((h,i) =>
    styleRange(ws.getRange(row,i+1), { value:h, bg:'#38761D', fg:'#FFFFFF', bold:true, sz:9, align:'center' }));
  row++;

  const DIF_BG = {'Fácil':'#D9EAD3','Média fácil':'#B6D7A8','Média':'#6AA84F','Média difícil':'#38761D','Difícil':'#274E13'};
  const DIF_FG = {'Fácil':'#000000','Média fácil':'#000000','Média':'#FFFFFF','Média difícil':'#FFFFFF','Difícil':'#FFFFFF'};

  difKeys.forEach((df, di) => {
    const dm  = difMapC[df];
    const alt = di%2===1 ? '#EBF5EB' : '#FFFFFF';
    const med = dm.cnt>0 ? Math.round(dm.hEst/dm.cnt*10)/10 : 0;
    const dev = dm.deltas.length ? `${(dm.deltas.reduce((s,v)=>s+v,0)/dm.deltas.length*100).toFixed(1)}%` : '—';
    ws.setRowHeight(row, 20);
    styleRange(ws.getRange(row,1), { value:df,                   bg:DIF_BG[df]||alt, fg:DIF_FG[df]||'#000000', bold:true,  sz:10, align:'left'   });
    styleRange(ws.getRange(row,2), { value:dm.cnt,               bg:alt, fg:'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(row,3), { value:pctM(dm.cnt,nClosed), bg:alt, fg:'#595959', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(row,4), { value:dm.hEst,              bg:alt, fg:'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(row,5), { value:dm.hReal,             bg:alt, fg:'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(row,6), { value:dm.pesoSum,           bg:alt, fg:'#0B5394', bold:true,  sz:10, align:'center' });
    styleRange(ws.getRange(row,7), { value:med,                  bg:alt, fg:'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(row,8), { value:dev,                  bg:alt, fg:'#595959', bold:false, sz:10, align:'center' });
    row++;
  });

  row++;
  ws.setRowHeight(row-1, 12);

  // ── DISTRIBUIÇÃO DE PRECISÃO DE ESTIMATIVA
  ws.setRowHeight(row, 20);
  styleRange(ws.getRange(row,1,1,NCOLS).merge(), {
    value: '  PRECISÃO DE ESTIMATIVA — DISTRIBUIÇÃO (tarefas Closed com dados)',
    bg: '#274E13', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });
  row++;

  ws.setRowHeight(row, 18);
  ws.setColumnWidth(1,210); ws.setColumnWidth(2,70); ws.setColumnWidth(3,100); ws.setColumnWidth(4,120);
  ['Faixa de precisão','Qtd','% de avaliadas','H. médias rastreadas'].forEach((h,i) =>
    styleRange(ws.getRange(row,i+1), { value:h, bg:'#38761D', fg:'#FFFFFF', bold:true, sz:9, align:'center' }));
  row++;

  const comDadosC = closed.filter(d => d.horasEst>0 && d.horasReal>0);
  const faixas = [
    ['Muito abaixo  (< -30%)', d => (d.horasReal-d.horasEst)/d.horasEst < -0.30,                                   '#1F3864','#FFFFFF'],
    ['Abaixo  (-30% a -11%)',  d => { const r=(d.horasReal-d.horasEst)/d.horasEst; return r>=-0.30&&r<-0.10; },    '#2E75B6','#FFFFFF'],
    ['OK  (±10%)',             d => Math.abs((d.horasReal-d.horasEst)/d.horasEst)<=0.10,                           '#38761D','#FFFFFF'],
    ['Acima  (+11% a +30%)',   d => { const r=(d.horasReal-d.horasEst)/d.horasEst; return r>0.10&&r<=0.30; },      '#BF8F00','#FFFFFF'],
    ['Muito acima  (> +30%)',  d => (d.horasReal-d.horasEst)/d.horasEst > 0.30,                                    '#C00000','#FFFFFF'],
  ];

  faixas.forEach(([label,fn,bg,fg], fi) => {
    const grupo = comDadosC.filter(fn);
    const med   = grupo.length>0 ? Math.round(grupo.reduce((s,d)=>s+d.horasReal,0)/grupo.length*10)/10 : 0;
    const alt   = fi%2===1 ? '#EBF5EB' : '#FFFFFF';
    ws.setRowHeight(row, 20);
    styleRange(ws.getRange(row,1), { value:label,                          bg:bg,  fg:fg,       bold:true,  sz:10, align:'left'   });
    styleRange(ws.getRange(row,2), { value:grupo.length,                   bg:alt, fg:'#000000', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(row,3), { value:pctM(grupo.length,comDadosC.length||1), bg:alt, fg:'#595959', bold:false, sz:10, align:'center' });
    styleRange(ws.getRange(row,4), { value:grupo.length>0?med:'—',         bg:alt, fg:'#000000', bold:false, sz:10, align:'center' });
    row++;
  });
}
