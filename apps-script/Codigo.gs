/**
 * ============================================================================
 *  VESTIBULAR IPEDE 2027.1 — PÁGINA DE INSCRIÇÃO + REDAÇÃO ONLINE
 *  Faculdades IPEDE – Codó/MA
 * ----------------------------------------------------------------------------
 *  Fluxo do candidato:
 *   1. Escolhe a forma de ingresso e preenche a inscrição  -> aba "Leads"
 *   2. (Vestibular Online) faz a redação sobre o tema atual -> aba "Redações"
 *   3. É direcionado ao WhatsApp da recepção com mensagem pronta
 *
 *  Publicação: Implantar > Nova implantação > Tipo "App da Web"
 *              Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 * ============================================================================
 */

const CFG = {
  planilhaNome: 'Inscrições – Vestibular IPEDE 2027.1',
  whatsapp: '5599981930161',          // (99) 98193-0161 — recepção
  minPalavras: 120,
  maxPalavras: 700,
  cor: '#006450',
};

const ABA_LEADS = 'Leads';
const ABA_REDACOES = 'Redações';

const CAB_LEADS = ['Data/hora', 'Protocolo', 'Nome completo', 'CPF', 'Nascimento', 'E-mail', 'WhatsApp',
  'Cidade/UF', 'Curso', 'Turno', 'Forma de ingresso', 'Como conheceu a IPEDE', 'Aceite LGPD',
  'Status', 'Redação enviada em', 'Palavras na redação'];
const CAB_REDACOES = ['Data/hora', 'Protocolo', 'Nome', 'Curso', 'Tema', 'Título', 'Texto da redação',
  'Palavras', 'Tempo de escrita (min)', 'Nota (0 a 10)', 'Avaliador', 'Observações'];

/* ============================== WEB APP ================================== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Vestibular IPEDE 2027.1 – Inscrições abertas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Recebe os envios da página hospedada fora do Google (GitHub Pages). */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.acao === 'inscricao') out = { ok: true, r: salvarInscricao(req.dados) };
    else if (req.acao === 'redacao') out = { ok: true, r: salvarRedacao(req.dados) };
    else throw new Error('Ação inválida.');
  } catch (err) {
    out = { ok: false, erro: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================== PLANILHA ================================= */
/** Rode uma vez pelo editor: cria (ou localiza) a planilha de leads. */
function configurar() {
  const ss = planilha_();
  Logger.log('Planilha de inscrições: ' + ss.getUrl());
  return ss.getUrl();
}

function planilha_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(CFG.planilhaNome);
    props.setProperty('SHEET_ID', ss.getId());
  }
  prepararAba_(ss, ABA_LEADS, CAB_LEADS, [150, 120, 240, 120, 100, 220, 140, 150, 140, 100, 170, 170, 90, 150, 150, 90]);
  prepararAba_(ss, ABA_REDACOES, CAB_REDACOES, [150, 120, 220, 140, 260, 220, 520, 80, 90, 90, 140, 220]);
  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 2) ss.deleteSheet(padrao);
  return ss;
}

function prepararAba_(ss, nome, cab, larguras) {
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold')
      .setBackground(CFG.cor).setFontColor('#FFFFFF').setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 32);
    larguras.forEach((w, i) => sh.setColumnWidth(i + 1, w));
    if (nome === ABA_REDACOES) sh.getRange('G:G').setWrap(true).setVerticalAlignment('top');
    sh.getRange('D:D').setNumberFormat('@'); // CPF / Curso como texto
  }
  return sh;
}

/* ======================== CHAMADAS DA PÁGINA ============================= */
function salvarInscricao(d) {
  d = d || {};
  const limpo = s => String(s == null ? '' : s).trim().replace(/\s+/g, ' ').slice(0, 200);
  const nome = limpo(d.nome).toUpperCase();
  const cpf = limpo(d.cpf).replace(/\D/g, '');
  const email = limpo(d.email).toLowerCase();
  const whats = limpo(d.whatsapp).replace(/\D/g, '');

  if (nome.split(' ').length < 2) throw new Error('Informe o nome completo.');
  if (!cpfValido_(cpf)) throw new Error('CPF inválido. Confira os números.');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('E-mail inválido.');
  if (whats.length < 10 || whats.length > 11) throw new Error('WhatsApp inválido. Use DDD + número.');
  if (!limpo(d.curso)) throw new Error('Escolha o curso.');
  if (d.lgpd !== true) throw new Error('É preciso aceitar o uso dos dados para a inscrição.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = planilha_().getSheetByName(ABA_LEADS);
    // Já inscrito com este CPF? Devolve o mesmo protocolo.
    const n = sh.getLastRow() - 1;
    if (n > 0) {
      const dados = sh.getRange(2, 1, n, CAB_LEADS.length).getValues();
      for (let i = 0; i < dados.length; i++) {
        if (String(dados[i][3]).replace(/\D/g, '') === cpf) {
          return { protocolo: dados[i][1], nome: dados[i][2], curso: dados[i][8],
            ingresso: dados[i][10], jaInscrito: true, redacaoEnviada: !!dados[i][14] };
        }
      }
    }
    const protocolo = 'IPEDE27-' + ('00000' + (n + 1)).slice(-5);
    const ingresso = limpo(d.ingresso) || 'Vestibular Online';
    sh.appendRow([new Date(), protocolo, nome, formatarCpf_(cpf), limpo(d.nascimento), email,
      formatarFone_(whats), limpo(d.cidade), limpo(d.curso), limpo(d.turno), ingresso,
      limpo(d.origem), 'Sim', ingresso === 'Vestibular Online' ? 'Aguardando redação' : 'Inscrito – contato WhatsApp',
      '', '']);
    return { protocolo, nome, curso: limpo(d.curso), ingresso, jaInscrito: false, redacaoEnviada: false };
  } finally {
    lock.releaseLock();
  }
}

function salvarRedacao(r) {
  r = r || {};
  const protocolo = String(r.protocolo || '').trim();
  const texto = String(r.texto || '').replace(/\r/g, '').trim().slice(0, 12000);
  const palavras = (texto.match(/[A-Za-zÀ-ÿ0-9]+/g) || []).length;
  if (!/^IPEDE27-\d{5}$/.test(protocolo)) throw new Error('Protocolo inválido. Refaça a inscrição.');
  if (palavras < CFG.minPalavras) throw new Error('A redação precisa ter pelo menos ' + CFG.minPalavras + ' palavras.');
  if (palavras > CFG.maxPalavras) throw new Error('A redação ultrapassou ' + CFG.maxPalavras + ' palavras.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = planilha_();
    const leads = ss.getSheetByName(ABA_LEADS);
    const n = leads.getLastRow() - 1;
    const prot = n > 0 ? leads.getRange(2, 2, n, 1).getValues().map(x => x[0]) : [];
    const idx = prot.indexOf(protocolo);
    if (idx < 0) throw new Error('Inscrição não encontrada. Refaça a inscrição.');
    const linha = idx + 2;
    const lead = leads.getRange(linha, 1, 1, CAB_LEADS.length).getValues()[0];
    if (lead[14]) return { ok: true, jaEnviada: true, nome: lead[2], curso: lead[8] };

    const agora = new Date();
    ss.getSheetByName(ABA_REDACOES).appendRow([agora, protocolo, lead[2], lead[8],
      String(r.tema || '').slice(0, 300), String(r.titulo || '').trim().slice(0, 150), texto,
      palavras, Math.max(0, Math.round(Number(r.minutos) || 0)), '', '', '']);
    leads.getRange(linha, 14, 1, 3).setValues([['Redação enviada – aguardando matrícula', agora, palavras]]);
    return { ok: true, jaEnviada: false, nome: lead[2], curso: lead[8] };
  } finally {
    lock.releaseLock();
  }
}

function contatoWhatsApp() { return CFG.whatsapp; }

/* ============================ UTILITÁRIOS ================================ */
function cpfValido_(c) {
  if (!/^\d{11}$/.test(c) || /^(\d)\1{10}$/.test(c)) return false;
  for (let t = 9; t < 11; t++) {
    let s = 0;
    for (let i = 0; i < t; i++) s += Number(c[i]) * (t + 1 - i);
    if (((s * 10) % 11) % 10 !== Number(c[t])) return false;
  }
  return true;
}
function formatarCpf_(c) { return `${c.slice(0, 3)}.${c.slice(3, 6)}.${c.slice(6, 9)}-${c.slice(9)}`; }
function formatarFone_(f) {
  return f.length === 11 ? `(${f.slice(0, 2)}) ${f.slice(2, 7)}-${f.slice(7)}` : `(${f.slice(0, 2)}) ${f.slice(2, 6)}-${f.slice(6)}`;
}
