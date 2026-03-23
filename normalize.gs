/**
 * =========================
 * SEQUÊNCIA CANÔNICA DE STATUS
 * =========================
 */
const STATUS_CANONICO = [
  'Backlog', 'A fazer', 'Em andamento', 'Ajustes de revisão',
  'Aguardando', 'Aguardando cliente', 'Aguardando TA',
  'Em revisão', 'Em testes', 'Closed',
];

/**
 * =========================
 * PIPELINE COMPLETO
 * Normaliza dados + Gera Gantt
 * =========================
 */
function pipelineCompleto() {
  normalizarDados();
  gerarGantt();
}

/**
 * =========================
 * NORMALIZAÇÃO PRINCIPAL
 * =========================
 */
function normalizarDados() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Quadro TI");

  if (!rawSheet) throw new Error("Aba 'Quadro TI' não encontrada.");

  const raw     = rawSheet.getDataRange().getValues();
  const headers = raw[0];
  const rows    = raw.slice(1);
  const data    = rows.map(row => mapRow(headers, row));

  const ws   = getOrCreateSheet(ss, "Quadro TI Normalizado", '#0B5394');
  const COLS = 15;
  const today = new Date(); today.setHours(0, 0, 0, 0);

  // ── KPIs
  const total       = data.length;
  const emAndamento = data.filter(d => d.status === 'Em andamento').length;
  const concluidas  = data.filter(d => d.status === 'Closed').length;
  const atrasadas   = data.filter(d => d.atrasado).length;
  const semResp     = data.filter(d => !d.pessoa || d.pessoa === 'Sem responsável').length;

  // ── Row 1: título
  const dataStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  ws.setRowHeight(1, 30);
  styleRange(ws.getRange(1, 1, 1, COLS).merge(), {
    value: `QUADRO TI NORMALIZADO  —  Atualizado em ${dataStr}`,
    bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 13, align: 'left',
  });

  // ── Row 2: spacer
  ws.setRowHeight(2, 8);
  ws.getRange(2, 1, 1, COLS).setBackground('#1F3864');

  // ── Rows 3–4: KPI cards
  const kpis = [
    ['Total tarefas',   total,       '#1F3864'],
    ['Em andamento',    emAndamento, '#0070C0'],
    ['Concluídas',      concluidas,  '#38761D'],
    ['Atrasadas',       atrasadas,   atrasadas > 0 ? '#C00000' : '#38761D'],
    ['Sem responsável', semResp,     semResp  > 0 ? '#BF8F00' : '#595959'],
  ];
  kpis.forEach(([lbl, val, color], ki) => {
    ws.setColumnWidth(ki + 1, 130);
    styleRange(ws.getRange(3, ki + 1), { value: lbl, bg: color,     fg: '#FFFFFF', bold: true,  sz: 9,  align: 'center' });
    styleRange(ws.getRange(4, ki + 1), { value: val, bg: '#FFFFFF', fg: color,     bold: true,  sz: 22, align: 'center' });
  });
  ws.setRowHeight(3, 18);
  ws.setRowHeight(4, 42);
  if (COLS > kpis.length) {
    ws.getRange(3, kpis.length + 1, 2, COLS - kpis.length).setBackground('#FFFFFF');
  }

  // ── Row 5: spacer
  ws.setRowHeight(5, 14);

  // ── Row 6: cabeçalho da seção
  ws.setRowHeight(6, 22);
  styleRange(ws.getRange(6, 1, 1, COLS).merge(), {
    value: '  TAREFAS NORMALIZADAS',
    bg: '#1F3864', fg: '#FFFFFF', bold: true, sz: 10, align: 'left',
  });

  // ── Row 7: cabeçalhos das colunas
  const colHeaders = [
    'Nome', 'Cliente', 'Projeto', 'Responsável', 'Status',
    'Dificuldade', 'Peso', 'H. Estimadas', 'H. Reais', 'Delta (h)',
    'Tipo', 'Prioridade', 'Data Criação', 'Data Final', 'Atrasado?',
  ];
  const colWidths = [200, 120, 150, 120, 130, 80, 60, 90, 80, 80, 120, 100, 100, 100, 80];
  colHeaders.forEach((h, i) => {
    ws.setColumnWidth(i + 1, colWidths[i]);
    styleRange(ws.getRange(7, i + 1), { value: h, bg: '#2E75B6', fg: '#FFFFFF', bold: true, sz: 9, align: 'center' });
  });
  ws.setRowHeight(7, 20);

  // ── Rows 8+: dados
  const STATUS_COLORS = {
    'Closed':             ['#D9EAD3', '#38761D'],
    'Em andamento':       ['#DEEAF1', '#2E75B6'],
    'Em revisão':         ['#EAD1F7', '#7030A0'],
    'Em testes':          ['#E6F3FF', '#0B5394'],
    'Ajustes de revisão': ['#FDE9D9', '#C55A11'],
    'A fazer':            ['#F4F4F4', '#595959'],
    'Backlog':            ['#EEEEEE', '#757575'],
    'Aguardando':         ['#FFF3E0', '#E65100'],
    'Aguardando cliente': ['#FFF3E0', '#BF8F00'],
    'Aguardando TA':      ['#FFF3E0', '#BF8F00'],
  };

  data.forEach((d, ri) => {
    const row   = 8 + ri;
    const altBg = ri % 2 === 1 ? '#F0F4FA' : '#FFFFFF';
    const sc    = STATUS_COLORS[d.status] || [altBg, '#000000'];
    ws.setRowHeight(row, 20);

    [
      { v: d.nome,           bg: altBg,                                fg: '#000000',  bold: false, align: 'left'   },
      { v: d.cliente,        bg: altBg,                                fg: '#000000',  bold: false, align: 'left'   },
      { v: d.projeto,        bg: altBg,                                fg: '#000000',  bold: false, align: 'left'   },
      { v: d.pessoa,         bg: altBg,                                fg: '#000000',  bold: false, align: 'left'   },
      { v: d.status,         bg: sc[0],                                fg: sc[1],      bold: true,  align: 'center' },
      { v: d.dificuldade,    bg: altBg,                                fg: '#000000',  bold: false, align: 'center' },
      { v: d.peso,           bg: altBg,                                fg: '#000000',  bold: false, align: 'center' },
      { v: d.horasEstimadas, bg: altBg,                                fg: '#000000',  bold: false, align: 'center' },
      { v: d.horasReal,      bg: altBg,                                fg: '#000000',  bold: false, align: 'center' },
      { v: d.delta,          bg: d.delta < 0 ? '#FFF2CC' : altBg,     fg: d.delta < 0 ? '#BF8F00' : '#000000', bold: d.delta < 0, align: 'center' },
      { v: d.tipo,           bg: altBg,                                fg: '#000000',  bold: false, align: 'left'   },
      { v: d.prioridade,     bg: altBg,                                fg: '#000000',  bold: false, align: 'center' },
      { v: d.dataCriacao,    bg: altBg,                                fg: '#000000',  bold: false, align: 'center', fmt: 'DD/MM/YYYY' },
      { v: d.dataFinal,      bg: altBg,                                fg: '#000000',  bold: false, align: 'center', fmt: 'DD/MM/YYYY' },
      { v: d.atrasado ? '⚠ SIM' : 'OK', bg: d.atrasado ? '#FDECEA' : '#D9EAD3', fg: d.atrasado ? '#C00000' : '#38761D', bold: true, align: 'center' },
    ].forEach((cell, ci) => {
      const r = ws.getRange(row, ci + 1);
      styleRange(r, { value: cell.v, bg: cell.bg, fg: cell.fg, bold: cell.bold, sz: 10, align: cell.align });
      if (cell.fmt) r.setNumberFormat(cell.fmt);
    });
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${total} tarefas | ${emAndamento} em andamento | ${atrasadas} atrasadas`,
    'Dados normalizados!', 6
  );
  Logger.log("Normalização concluída.");
}

/**
 * =========================
 * MAPEAMENTO DE LINHA
 * =========================
 */
function mapRow(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);

  const horasEstimadas = parseHoras(obj["Tempo estimado"]);
  const horasReal      = parseHoras(obj["Tempo rastreado"]);
  const clienteProjeto = splitClienteProjeto(obj["Cliente - Projeto"]);

  return {
    nome:          removerLink(obj["Nome"] || ""),
    cliente:       clienteProjeto.cliente,
    projeto:       clienteProjeto.projeto,
    pessoa:        obj["Responsável"]    || "Sem responsável",
    status:        normalizarStatus(obj["Status"]),
    dificuldade:   obj["Dificuldade"]    || "",
    peso:          toNumber(obj["Dificuldade (Peso)"]),
    horasEstimadas,
    horasReal,
    delta:         horasReal - horasEstimadas,
    tipo:          obj["Tipo da tarefa"] || "",
    prioridade:    obj["Prioridade"]     || "",
    dataCriacao:   parseDate(obj["Data de criação"]),
    dataFinal:     parseDate(obj["Data final"]),
    atrasado:      isAtrasado(obj),
  };
}

// ── Remove links do Notion: "[texto](url)" → "texto" e URLs soltas ────────────
function removerLink(texto) {
  if (!texto) return texto;
  let s = texto.toString();
  // Markdown link: [texto](url) → texto
  s = s.replace(/\[([^\]]+)\]\([^)]+\)/g, '$1');
  // URL solta: http(s)://... → ''
  s = s.replace(/https?:\/\/\S+/g, '');
  return s.trim();
}

// ── Cliente / Projeto ────────────────────────────────────────────────────────
function splitClienteProjeto(valor) {
  if (!valor) return { cliente: "Sem cliente", projeto: "" };
  const limpo  = removerLink(valor.toString());
  const partes = limpo.split(" - ");
  return { cliente: partes[0].trim() || "Sem cliente", projeto: (partes[1] || "").trim() };
}

// ── Tempo → Horas ────────────────────────────────────────────────────────────
function parseHoras(valor) {
  if (!valor) return 0;
  const str = valor.toString().toLowerCase();
  let horas = 0;
  const hMatch = str.match(/(\d+)\s*h/);
  const mMatch = str.match(/(\d+)\s*m/);
  if (hMatch) horas += parseInt(hMatch[1]);
  if (mMatch) horas += parseInt(mMatch[1]) / 60;
  if (!hMatch && !mMatch && !isNaN(str)) return parseFloat(str);
  return horas;
}

// ── Número seguro ────────────────────────────────────────────────────────────
function toNumber(val) {
  if (!val) return 0;
  if (typeof val === "number") return val;
  const parsed = parseFloat(val);
  return isNaN(parsed) ? 0 : parsed;
}

// ── Data ─────────────────────────────────────────────────────────────────────
const MESES_PT = {
  'janeiro':1,'fevereiro':2,'março':3,'marco':3,'abril':4,
  'maio':5,'junho':6,'julho':7,'agosto':8,'setembro':9,
  'outubro':10,'novembro':11,'dezembro':12
};

function parseDate(val) {
  if (!val) return "";

  // GAS já entregou um Date object (caso normal com getValues())
  if (val instanceof Date) return isNaN(val.getTime()) ? "" : val;

  const str = val.toString().trim();

  // Formato brasileiro DD/MM/YYYY (ou DD/MM/YY)
  const dmyMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (dmyMatch) {
    let year = parseInt(dmyMatch[3]);
    if (year < 100) year += year < 50 ? 2000 : 1900;
    const d = new Date(year, parseInt(dmyMatch[2]) - 1, parseInt(dmyMatch[1]));
    return isNaN(d.getTime()) ? "" : d;
  }

  // Formato por extenso em português: "17 de março de 2026"
  const ptMatch = str.match(/^(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})$/i);
  if (ptMatch) {
    const mes = MESES_PT[ptMatch[2].toLowerCase()];
    if (mes) {
      const d = new Date(parseInt(ptMatch[3]), mes - 1, parseInt(ptMatch[1]));
      return isNaN(d.getTime()) ? "" : d;
    }
  }

  // ISO ou outros formatos reconhecíveis pelo engine
  try {
    const d = new Date(val);
    return isNaN(d.getTime()) ? "" : d;
  } catch { return ""; }
}

// ── Status canônico ───────────────────────────────────────────────────────────
function normalizarStatus(status) {
  if (!status) return 'Backlog';
  const raw = status.toString().trim();
  const s   = raw.toLowerCase();

  // Correspondência exata (case-insensitive)
  const exact = STATUS_CANONICO.find(c => c.toLowerCase() === s);
  if (exact) return exact;

  // Fallbacks por palavras-chave
  if (s === 'done' || s.includes('concluí') || s.includes('feito')) return 'Closed';
  if (s.includes('aguardando cliente'))    return 'Aguardando cliente';
  if (s.includes('aguardando ta'))         return 'Aguardando TA';
  if (s.includes('aguardando'))            return 'Aguardando';
  if (s.includes('ajustes'))               return 'Ajustes de revisão';
  if (s.includes('andamento') || s.includes('progress')) return 'Em andamento';
  if (s.includes('revisão') || s.includes('revisao') || s.includes('review')) return 'Em revisão';
  if (s.includes('testes') || s.includes('test'))  return 'Em testes';
  if (s.includes('a fazer') || s.includes('todo') || s.includes('to do')) return 'A fazer';
  if (s.includes('backlog')) return 'Backlog';

  return raw; // mantém original se não encontrado
}

// ── Atraso ───────────────────────────────────────────────────────────────────
function isAtrasado(obj) {
  if (!obj["Vencimento"]) return false;
  const hoje = new Date();
  const venc = new Date(obj["Vencimento"]);
  return normalizarStatus(obj["Status"]) !== "Closed" && venc < hoje;
}
