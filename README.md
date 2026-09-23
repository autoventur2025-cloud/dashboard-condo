<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>📊 Dashboard - Prestação de Contas do Condomínio</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<style>
  * { margin:0; padding:0; box-sizing:border-box; font-family:'Segoe UI', Arial, sans-serif; }
  body { background:#0f172a; color:#e2e8f0; min-height:100vh; }
  header { background:linear-gradient(135deg,#1e293b,#334155); padding:20px 30px; border-bottom:3px solid #3b82f6; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; }
  header h1 { font-size:1.4rem; color:#fff; }
  header h1 span { color:#3b82f6; }
  .status-legend { display:flex; gap:15px; flex-wrap:wrap; font-size:0.8rem; }
  .dot { display:inline-block; width:12px; height:12px; border-radius:50%; margin-right:5px; vertical-align:middle; }
  .dot-lucro { background:#3b82f6; } .dot-prejuizo { background:#ef4444; }
  .dot-duvida { background:#eab308; } .dot-ok { background:#0a0a0a; border:2px solid #64748b; }
  .dot-receita { background:#22c55e; } .dot-queda { background:#f97316; }

  .container { max-width:1400px; margin:0 auto; padding:20px; }

  .upload-zone { border:3px dashed #3b82f6; border-radius:16px; background:#1e293b; padding:40px 20px; text-align:center; cursor:pointer; transition:all .3s; margin-bottom:25px; }
  .upload-zone:hover, .upload-zone.dragover { background:#273549; border-color:#60a5fa; transform:scale(1.01); }
  .upload-zone .icon { font-size:3.5rem; }
  .upload-zone h2 { margin:10px 0; color:#93c5fd; }
  .upload-zone p { color:#94a3b8; font-size:0.9rem; }
  #fileInput { position:absolute; width:1px; height:1px; opacity:0; overflow:hidden; }
  .file-list { display:flex; flex-wrap:wrap; gap:10px; margin-top:15px; justify-content:center; }
  .file-chip { background:#334155; padding:8px 14px; border-radius:20px; font-size:0.85rem; border:1px solid #475569; }
  #leituraStatus { margin-top:10px; font-size:0.85rem; color:#60a5fa; min-height:20px; }

  .filtros { background:#1e293b; border-radius:14px; padding:18px; margin-bottom:25px; display:flex; flex-wrap:wrap; gap:15px; align-items:end; }
  .filtro-grupo { display:flex; flex-direction:column; gap:5px; }
  .filtro-grupo label { font-size:0.75rem; color:#94a3b8; text-transform:uppercase; font-weight:600; }
  .filtro-grupo select { background:#0f172a; border:1px solid #475569; color:#e2e8f0; border-radius:8px; padding:9px 12px; font-size:0.9rem; min-width:140px; }
  .filtro-grupo select:focus { outline:none; border-color:#3b82f6; }
  .btn-limpar { background:#334155; color:#cbd5e1; border:1px solid #475569; border-radius:8px; padding:10px 18px; cursor:pointer; font-size:0.85rem; }
  .btn-limpar:hover { background:#ef4444; color:#fff; border-color:#ef4444; }
  .btn-buscar { background:#3b82f6; color:#fff; border:none; border-radius:8px; padding:10px 22px; cursor:pointer; font-size:0.85rem; font-weight:600; }
  .btn-buscar:hover { background:#2563eb; }
  .filtro-info { width:100%; font-size:0.8rem; color:#64748b; }

  .cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:15px; margin-bottom:25px; }
  .card { background:#1e293b; border-radius:14px; padding:20px; border-left:5px solid #475569; box-shadow:0 4px 12px rgba(0,0,0,.3); }
  .card h3 { font-size:0.8rem; color:#94a3b8; text-transform:uppercase; margin-bottom:8px; }
  .card .valor { font-size:1.5rem; font-weight:700; }
  .card .sub { font-size:0.72rem; margin-top:5px; color:#64748b; }
  .border-lucro { border-left-color:#3b82f6; } .border-prejuizo { border-left-color:#ef4444; }
  .border-duvida { border-left-color:#eab308; } .border-ok { border-left-color:#64748b; }
  .border-receita { border-left-color:#22c55e; } .border-queda { border-left-color:#f97316; }
  .text-lucro { color:#60a5fa; } .text-prejuizo { color:#f87171; }
  .text-duvida { color:#facc15; } .text-ok { color:#e2e8f0; }
  .text-receita { color:#4ade80; } .text-queda { color:#fb923c; }

  .charts { display:grid; grid-template-columns:repeat(auto-fit,minmax(380px,1fr)); gap:20px; margin-bottom:25px; }
  .chart-box { background:#1e293b; border-radius:14px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,.3); }
  .chart-box h3 { margin-bottom:15px; color:#cbd5e1; font-size:0.95rem; }
  .chart-wrap { position:relative; height:320px; }
  .full { grid-column:1/-1; }

  .tabs { display:flex; gap:5px; margin-bottom:0; flex-wrap:wrap; }
  .tab-btn { background:#334155; color:#94a3b8; border:none; padding:12px 22px; cursor:pointer; border-radius:10px 10px 0 0; font-size:0.87rem; font-weight:600; }
  .tab-btn.active { background:#1e293b; color:#60a5fa; border-bottom:2px solid #3b82f6; }
  .tab-content { display:none; background:#1e293b; padding:25px; border-radius:0 14px 14px 14px; margin-bottom:25px; }
  .tab-content.active { display:block; }

  table { width:100%; border-collapse:collapse; font-size:0.88rem; }
  th { background:#334155; padding:12px; text-align:left; color:#cbd5e1; position:sticky; top:0; }
  td { padding:11px; border-bottom:1px solid #334155; }
  tr:hover td { background:#273549; }
  .tbl-scroll { overflow-x:auto; max-height:500px; overflow-y:auto; border-radius:8px; }
  .badge { padding:4px 10px; border-radius:12px; font-size:0.73rem; font-weight:600; white-space:nowrap; }
  .b-lucro { background:rgba(59,130,246,.2); color:#60a5fa; }
  .b-prejuizo { background:rgba(239,68,68,.2); color:#f87171; }
  .b-duvida { background:rgba(234,179,8,.2); color:#facc15; }
  .b-ok { background:rgba(15,15,15,.6); color:#e2e8f0; border:1px solid #475569; }
  .b-receita { background:rgba(34,197,94,.2); color:#4ade80; }
  .b-queda { background:rgba(249,115,22,.2); color:#fb923c; }
  select.status-sel { background:#334155; color:#fff; border:1px solid #475569; border-radius:6px; padding:5px; font-size:0.78rem; }
  .obs-input { width:100%; min-width:150px; background:#0f172a; border:1px solid #475569; color:#e2e8f0; border-radius:6px; padding:8px; font-size:0.78rem; }
  .obs-input:focus { outline:none; border-color:#3b82f6; }

  .com-ok { color:#4ade80; font-weight:700; }
  .com-falta { color:#f87171; font-weight:700; }
  .msg-verificar { background:rgba(239,68,68,.15); color:#f87171; padding:3px 8px; border-radius:6px; font-size:0.72rem; font-weight:700; }

  .veredito { padding:15px; border-radius:10px; margin-bottom:15px; font-size:0.92rem; line-height:1.6; }
  .v-ok { background:rgba(59,130,246,.12); border:1px solid #3b82f6; }
  .v-alerta { background:rgba(234,179,8,.12); border:1px solid #eab308; }
  .v-perda { background:rgba(239,68,68,.12); border:1px solid #ef4444; }
  .v-receita { background:rgba(34,197,94,.12); border:1px solid #22c55e; }
  .v-queda { background:rgba(249,115,22,.12); border:1px solid #f97316; }

  .anexo-zone { border:2px dashed #475569; border-radius:12px; padding:35px; text-align:center; cursor:pointer; }
  .anexo-zone:hover { border-color:#3b82f6; }
  .anexos-lista { display:grid; grid-template-columns:repeat(auto-fill,minmax(260px,1fr)); gap:15px; margin-top:20px; }
  .anexo-card { background:#273549; border-radius:12px; padding:15px; display:flex; align-items:center; gap:12px; }
  .anexo-card .arq-icon { font-size:2rem; }
  .anexo-card .nome { font-size:0.85rem; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .anexo-card .data { font-size:0.7rem; color:#64748b; }
  .anexo-card a { color:#60a5fa; text-decoration:none; font-size:0.75rem; }

  .sim-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:15px; margin-bottom:20px; }
  .sim-card { background:#273549; border-radius:12px; padding:18px; border-top:4px solid; }
  .sim-card h4 { font-size:0.85rem; margin-bottom:10px; }
  .sim-card .rend { font-size:1.2rem; font-weight:700; }
  .sim-card .tx { font-size:0.75rem; color:#94a3b8; margin-top:5px; }
  .sim-card .ganho { font-size:0.78rem; margin-top:8px; padding-top:8px; border-top:1px solid #475569; }
  .table-mini th, .table-mini td { padding:8px 12px; font-size:0.8rem; }
  .vencimento-box { background:rgba(34,197,94,.08); border:1px solid #22c55e; border-radius:10px; padding:15px; margin-top:15px; font-size:0.88rem; }
  @media (max-width:500px){ .charts{grid-template-columns:1fr;} }
</style>
</head>
<body>

<header>
  <h1>🏢 Dashboard <span>Condomínio</span> — Prestação de Contas</h1>
  <div class="status-legend">
    <span><span class="dot dot-lucro"></span>Lucro</span>
    <span><span class="dot dot-prejuizo"></span>Prejuízo</span>
    <span><span class="dot dot-duvida"></span>Duvidoso</span>
    <span><span class="dot dot-ok"></span>Verificado OK</span>
    <span><span class="dot dot-receita"></span>Receita Total</span>
    <span><span class="dot dot-queda"></span>Queda</span>
  </div>
</header>

<div class="container">

  <div class="upload-zone" id="uploadZone">
    <div class="icon">📁</div>
    <h2>Arraste o arquivo da prestação de contas aqui ou clique para selecionar</h2>
    <p>Formatos: <b>PDF</b> • <b>XML</b> • <b>Excel (XLSX/CSV)</b></p>
    <input type="file" id="fileInput" accept=".pdf,.xml,.xlsx,.xls,.csv" multiple>
    <div class="file-list" id="fileList"></div>
    <div id="leituraStatus"></div>
  </div>

  <div class="filtros">
    <div class="filtro-grupo"><label>📅 Dia</label><select id="fDia"><option value="">Todos</option></select></div>
    <div class="filtro-grupo"><label>🗓️ Mês</label><select id="fMes"><option value="">Todos</option></select></div>
    <div class="filtro-grupo"><label>📆 Ano</label><select id="fAno"><option value="">Todos</option></select></div>
    <div class="filtro-grupo"><label>🏷️ Rubrica / Conta</label><select id="fRubrica"><option value="">Todas</option></select></div>
    <button class="btn-buscar" onclick="aplicarFiltros()">🔍 Buscar</button>
    <button class="btn-limpar" onclick="limparFiltros()">✖ Limpar</button>
    <div class="filtro-info" id="filtroInfo">Filtros ativos: nenhum — exibindo todos os lançamentos.</div>
  </div>

  <div class="cards" id="cards"></div>
  <div id="veredito"></div>

  <div class="charts">
    <div class="chart-box"><h3>💰 Distribuição — Receitas × Despesas × Desvios</h3><div class="chart-wrap"><canvas id="gRosca"></canvas></div></div>
    <div class="chart-box"><h3>📊 Saldos Disponíveis por Conta (R$)</h3><div class="chart-wrap"><canvas id="gSaldos"></canvas></div></div>
    <div class="chart-box full"><h3>📈 Evolução do Fundo de Reserva — Comparativo Mensal</h3><div class="chart-wrap"><canvas id="gLinha"></canvas></div></div>
  </div>

  <div class="tab-content active" style="padding:20px;">
    <h3 style="margin-bottom:15px;">🔍 Análise Detalhada de Lançamentos</h3>
    <div class="tbl-scroll">
      <table id="tabela">
        <thead><tr><th>Descrição</th><th>Data</th><th>Rubrica</th><th>Tipo</th><th>Valor (R$)</th><th>Status</th><th>Observações</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <div class="tabs">
    <button class="tab-btn active" onclick="abrirTab('tabDemonstrativo',this)">📑 Demonstrativo Receitas × Despesas</button>
    <button class="tab-btn" onclick="abrirTab('tabSaldos',this)">🏦 Saldos das Contas</button>
    <button class="tab-btn" onclick="abrirTab('tabComprovantes',this)">🧾 Controle de Comprovantes</button>
    <button class="tab-btn" onclick="abrirTab('tabInvestimento',this)">💼 Aplicação do Fundo de Reserva</button>
    <button class="tab-btn" onclick="abrirTab('tabAnaliseFinal',this)">✅ Análise Final Verificada</button>
    <button class="tab-btn" onclick="abrirTab('tabAnexos',this)">📎 Anexos</button>
  </div>

  <div class="tab-content active" id="tabDemonstrativo">
    <h3 style="margin-bottom:15px;">📑 Demonstrativo de Receitas e Despesas</h3>
    <div id="demonstrativoResumo"></div>
    <div class="tbl-scroll" style="margin-top:15px;">
      <table id="tblDemo">
        <thead><tr><th>Data</th><th>Descrição</th><th>Rubrica</th><th>Classificação</th><th>Valor (R$)</th><th>Situação</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <div class="tab-content" id="tabSaldos">
    <h3 style="margin-bottom:15px;">🏦 Análise de Saldos Disponíveis por Conta</h3>
    <div id="saldosResumo"></div>
    <div class="charts" style="margin-top:15px;">
      <div class="chart-box"><h3>💰 Saldo Atual por Conta</h3><div class="chart-wrap"><canvas id="gSaldosAba"></canvas></div></div>
      <div class="chart-box"><h3>📊 Comparativo — Mês Anterior × Mês Atual</h3><div class="chart-wrap"><canvas id="gSaldosComp"></canvas></div></div>
    </div>
    <div class="tbl-scroll" style="margin-top:15px;">
      <table id="tblSaldos">
        <thead><tr><th>Conta / Fundo</th><th>Mês Anterior (R$)</th><th>Mês Atual (R$)</th><th>Variação (R$)</th><th>Situação</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <div class="tab-content" id="tabComprovantes">
    <h3 style="margin-bottom:15px;">🧾 Informe de Comprovantes por Compra</h3>
    <div id="comprovantesResumo"></div>
    <div class="tbl-scroll" style="margin-top:15px;">
      <table id="tblComp">
        <thead><tr><th>Compra / Despesa</th><th>Data</th><th>Rubrica</th><th>Valor (R$)</th><th>Comprovante?</th><th>Ação Necessária</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <div class="tab-content" id="tabInvestimento">
    <h3 style="margin-bottom:15px;">💼 Projeção — Aplicação do Fundo de Reserva</h3>
    <div id="investResumo"></div>
    <div class="sim-grid" id="simCards"></div>
    <div class="tbl-scroll">
      <table class="table-mini" id="tblInv">
        <thead><tr><th>Aplicação</th><th>Rendimento /mês</th><th>6 meses</th><th>12 meses</th><th>24 meses</th><th>Liquidez</th><th>Risco</th><th>Proteção IPCA</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
    <div class="vencimento-box" id="vencimentoBox"></div>
  </div>

  <div class="tab-content" id="tabAnaliseFinal">
    <h3 style="margin-bottom:15px;">✅ Gráfico Consolidado — Todas as Análises Verificadas</h3>
    <div id="analiseFinalResumo"></div>
    <div class="chart-box full" style="margin-top:15px; background:#273549;"><div class="chart-wrap" style="height:360px;"><canvas id="gFinal"></canvas></div></div>
  </div>

  <div class="tab-content" id="tabAnexos">
    <h3 style="margin-bottom:15px;">📎 Arquivos Anexados ao Dashboard</h3>
    <div class="anexo-zone" id="anexoZone">
      <div style="font-size:2.5rem;">➕</div>
      <p style="color:#94a3b8;">Clique para anexar documentos (boletos, notas, extratos, comprovantes...)</p>
      <input type="file" id="anexoInput" multiple style="display:none;">
    </div>
    <div class="anexos-lista" id="anexosLista"></div>
  </div>

</div>

<script>
'use strict';

/* ===== Verifica se as bibliotecas carregaram ===== */
window.addEventListener('DOMContentLoaded', function() {
  if (typeof pdfjsLib === 'undefined') {
    document.getElementById('leituraStatus').textContent = '⚠️ Biblioteca PDF não carregou — verifique sua conexão com a internet e recarregue a página (F5).';
  }
});
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const MESES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const MESES_FULL = ['janeiro','fevereiro','marco','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
const RUBRICAS = ['Fundo Ordinário','Fundo de Reserva','Energia Elétrica','Água e Esgoto','Fundo Fixo','Outras'];
const CORES = { lucro:'#3b82f6', prejuizo:'#ef4444', duvida:'#eab308', ok:'#0a0a0a', receita:'#22c55e', queda:'#f97316' };
const CORES_CONTA = { 'Fundo Ordinário':'#3b82f6', 'Fundo de Reserva':'#22c55e', 'Energia Elétrica':'#eab308', 'Água e Esgoto':'#06b6d4', 'Fundo Fixo':'#a855f7', 'Outras':'#64748b' };

let lancamentos = [], saldos = [], anexos = [], charts = {}, filtroAtivo = {};

/* ===== UTILITÁRIOS ===== */
function fmt(v) { return (v || 0).toLocaleString('pt-BR', { style:'currency', currency:'BRL' }); }

function normalizarAcc(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function parseValorBR(s) {
  if (s === null || s === undefined) return null;
  s = String(s).replace(/[^\d.,-]/g, '');
  if (!s) return null;
  if (s.includes(',')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if ((s.match(/\./g) || []).length === 1 && /\.\d{1,2}$/.test(s) && !/^\d{1,3}\.\d{3}$/.test(s)) {
    // mantém como decimal "1234.56"
  } else {
    s = s.replace(/\./g, '');
  }
  const v = parseFloat(s);
  return isNaN(v) ? null : v;
}

const RX_MONETARIO = /(-)?\s*\(?\s*(?:R\$|RS)\s*(-)?\s*((?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d{2})?)\s*\)?/g;
const RX_NUMERO_SOLTO = /\(?\s*(-)?\s*((?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d{2})?)\s*\)?(?!\s*[%\/])/g;

function extrairValoresLinha(linha) {
  const out = [];
  let m;
  // 1º tenta valores com R$ (mais confiáveis)
  RX_MONETARIO.lastIndex = 0;
  while ((m = RX_MONETARIO.exec(linha)) !== null) {
    const neg = !!(m[1] || m[2]) || /^\(/.test(m[0].trim());
    const v = parseValorBR(m[3]);
    if (v !== null && v !== 0) out.push({ valor: v, negativo: neg });
  }
  // Se nada com R$, tenta números soltos no padrão brasileiro (1.234,56)
  if (!out.length) {
    RX_NUMERO_SOLTO.lastIndex = 0;
    while ((m = RX_NUMERO_SOLTO.exec(linha)) !== null) {
      const raw = m[2];
      // Só aceita se tiver formato claramente monetário (com centavos ou milhar)
      if (!raw.includes(',') && !raw.includes('.')) continue;
      if (!/\d{3}/.test(raw) && !raw.includes(',')) continue;
      const neg = !!m[1] || /^\(/.test(m[0].trim());
      const v = parseValorBR(raw);
      if (v !== null && v !== 0) out.push({ valor: v, negativo: neg });
    }
  }
  return out;
}

/* ===== LEITURA DE ARQUIVOS ===== */
const zone = document.getElementById('uploadZone');
const fileInput = document.getElementById('fileInput');

zone.addEventListener('click', function() { fileInput.click(); });
zone.addEventListener('dragover', function(e) { e.preventDefault(); zone.classList.add('dragover'); });
zone.addEventListener('dragleave', function() { zone.classList.remove('dragover'); });
zone.addEventListener('drop', function(e) { e.preventDefault(); zone.classList.remove('dragover'); processarArquivos(e.dataTransfer.files); });
fileInput.addEventListener('change', function(e) { processarArquivos(e.target.files); e.target.value = ''; });

async function processarArquivos(files) {
  const status = document.getElementById('leituraStatus');
  for (const f of files) {
    const ext = f.name.split('.').pop().toLowerCase();
    try {
      status.textContent = '⏳ Processando ' + f.name + '...';
      if (['xlsx','xls','csv'].includes(ext)) await lerExcel(f);
      else if (ext === 'xml') await lerXML(f);
      else if (ext === 'pdf') await lerPDF(f);
      else status.textContent = '⚠️ Formato não suportado: ' + ext;
      adicionarChip(f.name);
    } catch (err) {
      console.error(err);
      status.textContent = '❌ Erro ao processar ' + f.name + ': ' + err.message;
    }
  }
  if (lancamentos.length || saldos.length) {
    status.textContent = '✅ ' + lancamentos.length + ' lançamento(s) e ' + saldos.length + ' conta(s) carregados com sucesso.';
  } else {
    status.textContent = '⚠️ Nenhum valor reconhecido. Se o PDF for digitalizado (imagem), o texto não é selecionável — nesse caso use a planilha Excel.';
  }
  analisar();
  preencherFiltros();
  renderAll();
}

function adicionarChip(nome) {
  const chip = document.createElement('span');
  chip.className = 'file-chip';
  chip.textContent = '📄 ' + nome;
  document.getElementById('fileList').appendChild(chip);
}

/* ===== EXCEL ===== */
function lerExcel(file) {
  return new Promise(function(res, rej) {
    const r = new FileReader();
    r.onload = function(e) {
      try {
        const wb = XLSX.read(e.target.result, { type:'array' });
        wb.SheetNames.forEach(function(nome) {
          const dados = XLSX.utils.sheet_to_json(wb.Sheets[nome], { defval:'' });
          dados.forEach(function(l) {
            const lanc = normalizar(l, nome);
            if (lanc) lancamentos.push(lanc);
          });
          const s = extrairSaldos(dados, nome);
          if (s) saldos = saldos.concat(s);
        });
        res();
      } catch (err) { rej(err); }
    };
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });
}

/* ===== XML ===== */
function lerXML(file) {
  return new Promise(function(res, rej) {
    const r = new FileReader();
    r.onload = function(e) {
      try {
        const doc = new DOMParser().parseFromString(e.target.result, 'text/xml');
        doc.querySelectorAll('lancamento, lancamento, item, linha, det').forEach(function(n) {
          const v = lerValor(n.querySelector('valor, vl, vProd, vLiq')?.textContent);
          const d = n.querySelector('descricao, xProduto, historico')?.textContent || 'Lançamento XML';
          const rub = n.querySelector('rubrica, conta, categoria')?.textContent || 'Outras';
          if (v !== null) lancamentos.push({ descricao:d, dia:'', mes:0, ano:new Date().getFullYear(), rubrica:identRubrica(rub), tipo:v>=0?'receita':'despesa', valor:Math.abs(v), status:'duvida', comprovante:v>=0?'sim':'nao', obs:'' });
        });
        res();
      } catch (err) { rej(err); }
    };
    r.onerror = rej;
    r.readAsText(file);
  });
}

/* ===== PDF ===== */
async function lerPDF(file) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

  /* Reconstrói as linhas reais do PDF */
  const linhas = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const c = await page.getTextContent();
    let linha = '';
    c.items.forEach(function(it) {
      if (typeof it.str !== 'string') return;
      linha += it.str;
      if (it.hasEOL) {
        linhas.push(linha.trim());
        linha = '';
      } else if (linha.length && !linha.endsWith(' ')) {
        linha += ' ';
      }
    });
    if (linha.trim()) linhas.push(linha.trim());
  }

  const texto = linhas.join('\n');
  const textoNorm = normalizarAcc(texto);

  /* 1) Saldos das contas (sem duplicar) */
  RUBRICAS.forEach(function(rub) {
    const rubN = normalizarAcc(rub);
    const rx = new RegExp(rubN.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '[^\\n]{0,150}?((?:\\d{1,3}(?:\\.\\d{3})+|\\d+)(?:,\\d{2})?)', 'gi');
    const achados = [];
    let m;
    while ((m = rx.exec(textoNorm)) !== null) {
      const v = parseValorBR(m[1]);
      if (v !== null && v !== 0) achados.push(v);
      if (achados.length > 6) break;
    }
    if (achados.length) {
      const atual = achados[achados.length - 1];
      const anterior = achados.length > 1 ? achados[achados.length - 2] : 0;
      const idx = saldos.findIndex(function(s) { return s.conta === rub; });
      const novo = { conta: rub, anterior: anterior, atual: atual };
      if (idx >= 0) saldos[idx] = novo; else saldos.push(novo);
    }
  });

  /* 2) Lançamentos linha por linha */
  let secao = null;

  linhas.forEach(function(l) {
    const trim = l.trim();
    if (trim.length < 4) return;
    const tn = normalizarAcc(trim).toLowerCase();

    /* Detecta seção do demonstrativo */
    if (trim.length < 70) {
      if (/^receitas?\b/.test(tn)) { secao = 'receita'; return; }
      if (/^despesas?\b|^saidas?\b|^pagamentos?\b|^gastos?\b/.test(tn)) { secao = 'despesa'; return; }
    }

    if (/^total|^saldo|^saldo anterior|^saldo atual/.test(tn)) return;

    const valores = extrairValoresLinha(trim);
    if (!valores.length) return;

    const data = extrairData(trim);
    const ultimo = valores[valores.length - 1];

    let desc = trim
      .replace(RX_MONETARIO, ' ')
      .replace(RX_NUMERO_SOLTO, ' ')
      .replace(/\d{1,2}[\/\-.]\d{1,2}([\/\-.]\d{2,4})?/g, ' ')
      .replace(/\s{2,}/g, ' ')
      .trim();
    if (desc.length < 3) desc = 'Lançamento PDF';
    desc = desc.slice(0, 80);

    const despKw = ['pagamento','pago ','despesa','debito','saida','compra','fatura','boleto','energia','luz','agua','esgoto','salario','manutencao','limpeza','porteiro','jardim','seguranca','elevador','obra','reforma','taxa'];
    const recKw = ['recebimento','recebido','receita','credito','entrada','cota','repasse','rendimento','deposito','condominio'];

    let tipo;
    if (ultimo.negativo) tipo = 'despesa';
    else if (recKw.some(function(k) { return tn.includes(k); })) tipo = 'receita';
    else if (despKw.some(function(k) { return tn.includes(k); })) tipo = 'despesa';
    else tipo = secao || 'despesa';

    lancamentos.push({
      descricao: desc,
      dia: data ? data.dia : '',
      mes: data ? data.mes : 0,
      ano: data ? data.ano : new Date().getFullYear(),
      rubrica: identRubrica(desc),
      tipo: tipo,
      valor: Math.abs(ultimo.valor),
      status: 'duvida',
      comprovante: detectarComprovante(tn),
      obs: ''
    });
  });
}

/* Datas: dd/mm/aaaa, mm/aaaa, Janeiro/2024, Jan/24 */
function extrairData(l) {
  const tn = normalizarAcc(l).toLowerCase();
  let m = tn.match(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/);
  if (m) {
    return { dia: m[1], mes: Math.min(Math.max(parseInt(m[2]) - 1, 0), 11),
             ano: m[3].length === 2 ? 2000 + parseInt(m[3]) : parseInt(m[3]) };
  }
  m = tn.match(/\b(\d{1,2})[\/\-.](20\d{2})\b/);
  if (m) return { dia: '', mes: Math.min(Math.max(parseInt(m[1]) - 1, 0), 11), ano: parseInt(m[2]) };
  for (let i = 0; i < 12; i++) {
    if (tn.includes(MESES_FULL[i]) || new RegExp('\\b' + MESES[i].toLowerCase() + '\\b').test(tn)) {
      const y = tn.match(/(20\d{2})/);
      return { dia: '', mes: i, ano: y ? parseInt(y[1]) : new Date().getFullYear() };
    }
  }
  return null;
}

function lerValor(txt) {
  if (txt === null || txt === undefined || txt === '') return null;
  return parseValorBR(txt);
}

function identRubrica(txt) {
  const t = normalizarAcc(String(txt)).toLowerCase();
  if (t.includes('ordin')) return 'Fundo Ordinário';
  if (t.includes('reserv')) return 'Fundo de Reserva';
  if (t.includes('energ') || t.includes('luz') || t.includes('eletr')) return 'Energia Elétrica';
  if (t.includes('agua') || t.includes('esgoto') || t.includes('sane')) return 'Água e Esgoto';
  if (t.includes('fixo')) return 'Fundo Fixo';
  return 'Outras';
}

function detectarComprovante(tn) {
  if (/\b(nf|nfe|nf-e|nota fiscal|recibo|comprovante|anexo)\b/.test(tn)) return 'sim';
  if (/sem compro|obscuro|ileg|n\/d|nao localizado/.test(tn)) return 'nao';
  return 'desconhecido';
}

/* ===== EXCEL: saldos ===== */
function extrairSaldos(dados, sheetNome) {
  const out = [];
  RUBRICAS.forEach(function(rub) {
    let atual = null, anterior = null;
    dados.forEach(function(l) {
      const keys = Object.keys(l);
      const getV = function() {
        for (let k = 0; k < arguments.length; k++) {
          const f = keys.find(function(x) { return x.toLowerCase().includes(arguments[k]); });
          if (f && l[f] !== '') return l[f];
        }
        return null;
      };
      const conta = normalizarAcc(String(getV('conta','rubrica','fundo','descricao') || '')).toLowerCase();
      if (identRubrica(conta) !== rub) return;
      const a = lerValor(getV('saldo atual','atual','saldo'));
      const p = lerValor(getV('anterior','saldo anterior'));
      if (a !== null) atual = a;
      if (p !== null) anterior = p;
    });
    if (atual !== null) out.push({ conta: rub, anterior: anterior ?? 0, atual: atual });
  });
  return out.length ? out : null;
}

/* ===== EXCEL: lançamentos ===== */
function normalizar(l, sheetNome) {
  const keys = Object.keys(l);
  const get = function() {
    for (let k = 0; k < arguments.length; k++) {
      const f = keys.find(function(x) { return x.toLowerCase().includes(arguments[k]); });
      if (f && l[f] !== '') return l[f];
    }
    return '';
  };
  const valor = lerValor(get('valor','vlr','total','credito','debito')) || 0;
  if (valor === 0) return null;
  const tipoStr = normalizarAcc(String(get('tipo','natureza','operacao'))).toLowerCase();
  let tipo = valor < 0 ? 'despesa' : 'receita';
  if (tipoStr.includes('desp') || tipoStr.includes('debit')) tipo = 'despesa';
  else if (tipoStr.includes('rec') || tipoStr.includes('cred')) tipo = 'receita';
  const dataStr = String(get('data','competencia'));
  let dia = '', mes = 0, ano = new Date().getFullYear();
  const dm = dataStr.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (dm) { dia = dm[1]; mes = parseInt(dm[2]) - 1; ano = dm[3].length === 2 ? 2000 + parseInt(dm[3]) : parseInt(dm[3]); }
  else {
    const dn = extrairData(dataStr);
    if (dn) { dia = dn.dia; mes = dn.mes; ano = dn.ano; }
  }
  const desc = String(get('descricao','historico','lancamento','documento') || sheetNome).slice(0, 80);
  const rubricaRaw = get('rubrica','conta','fundo','categoria');
  const compRaw = normalizarAcc(String(get('comprovante','documento'))).toLowerCase();
  return {
    descricao: desc, dia: dia, mes: mes, ano: ano,
    rubrica: rubricaRaw ? identRubrica(rubricaRaw) : identRubrica(desc),
    tipo: tipo, valor: Math.abs(valor), status: 'duvida',
    comprovante: /\b(sim|comprovante|anexo)\b/.test(compRaw) ? 'sim' : /(nao|sem compro|n\/d)/.test(compRaw) ? 'nao' : 'desconhecido',
    obs: ''
  };
}

/* ===== ANÁLISE AUTOMÁTICA ===== */
function analisar() {
  ['receita','despesa'].forEach(function(tipo) {
    const itens = lancamentos.filter(function(l) { return l.tipo === tipo && l.valor > 0; });
    if (itens.length < 3) { itens.forEach(function(l) { if (l.status === 'duvida') l.status = 'ok'; }); return; }
    const media = itens.reduce(function(s,l) { return s + l.valor; }, 0) / itens.length;
    itens.forEach(function(l) {
      if (l.status !== 'duvida') return;
      l.status = l.valor > media * 2.5 ? 'duvida' : 'ok';
    });
  });
  lancamentos.forEach(function(l) {
    if (l.tipo === 'despesa' && l.comprovante === 'nao' && l.status === 'ok') l.status = 'duvida';
  });
  if (!saldos.length) {
    RUBRICAS.forEach(function(rub) {
      const rec = lancamentos.filter(function(l) { return l.rubrica === rub && l.tipo === 'receita'; }).reduce(function(s,l) { return s + l.valor; }, 0);
      const desp = lancamentos.filter(function(l) { return l.rubrica === rub && l.tipo === 'despesa'; }).reduce(function(s,l) { return s + l.valor; }, 0);
      if (rec || desp) saldos.push({ conta: rub, anterior: 0, atual: rec - desp });
    });
  }
  saldos.forEach(function(s) { if (s.anterior === null || s.anterior === undefined) s.anterior = 0; });
}

/* ===== FILTROS ===== */
function preencherFiltros() {
  const setAno = [...new Set(lancamentos.map(function(l) { return l.ano; }))].sort(function(a,b) { return a - b; });
  const setRub = [...new Set(lancamentos.map(function(l) { return l.rubrica; }))];
  fill('fAno', setAno);
  fill('fRubrica', setRub);
  fill('fMes', MESES);
  const dias = [...new Set(lancamentos.map(function(l) { return l.dia; }).filter(Boolean))].sort(function(a,b) { return Number(a) - Number(b); });
  fill('fDia', dias);
}
function fill(id, arr) {
  const sel = document.getElementById(id);
  if (!sel) return;
  sel.innerHTML = '<option value="">Todos</option>';
  arr.forEach(function(v) { sel.innerHTML += '<option>' + v + '</option>'; });
}
function aplicarFiltros() {
  filtroAtivo = {
    dia: document.getElementById('fDia').value,
    mes: document.getElementById('fMes').value,
    ano: document.getElementById('fAno').value,
    rub: document.getElementById('fRubrica').value
  };
  renderAll();
  const ativos = Object.values(filtroAtivo).filter(Boolean).join(' • ');
  document.getElementById('filtroInfo').textContent = ativos ? 'Filtros ativos: ' + ativos : 'Filtros ativos: nenhum — exibindo todos os lançamentos.';
}
function limparFiltros() {
  ['fDia','fMes','fAno','fRubrica'].forEach(function(id) { document.getElementById(id).value = ''; });
  filtroAtivo = {};
  renderAll();
  document.getElementById('filtroInfo').textContent = 'Filtros ativos: nenhum — exibindo todos os lançamentos.';
}
function filtrados() {
  return lancamentos.filter(function(l) {
    return (!filtroAtivo.dia || String(l.dia) === filtroAtivo.dia) &&
           (!filtroAtivo.mes || MESES[l.mes] === filtroAtivo.mes) &&
           (!filtroAtivo.ano || String(l.ano) === filtroAtivo.ano) &&
           (!filtroAtivo.rub || l.rubrica === filtroAtivo.rub);
  });
}

/* ===== TOTAIS ===== */
function totais() {
  const ls = filtrados();
  const rec = ls.filter(function(l) { return l.tipo === 'receita'; }).reduce(function(s,l) { return s + l.valor; }, 0);
  const desp = ls.filter(function(l) { return l.tipo === 'despesa'; }).reduce(function(s,l) { return s + l.valor; }, 0);
  const duv = ls.filter(function(l) { return l.status === 'duvida'; });
  return { rec: rec, desp: desp, resultado: rec - desp, duvidosos: duv,
           valorDuvidoso: duv.reduce(function(s,l) { return s + l.valor; }, 0),
           reserva: rec - desp, total: ls.length };
}

function saldoReserva() {
  const s = saldos.find(function(x) { return x.conta === 'Fundo de Reserva'; });
  return s ? s.atual : totais().reserva;
}

/* ===== CARDS ===== */
function renderCards() {
  const t = totais();
  const semComp = filtrados().filter(function(l) { return l.tipo === 'despesa' && l.comprovante === 'nao'; });
  const queda = t.resultado < 0 ? t.resultado : 0;
  document.getElementById('cards').innerHTML =
    '<div class="card border-receita"><h3>🟢 Receita Total</h3><div class="valor text-receita">' + fmt(t.rec) + '</div><div class="sub">' + filtrados().filter(function(l){return l.tipo==='receita';}).length + ' recebimentos</div></div>' +
    '<div class="card border-prejuizo"><h3>🔴 Total de Prejuízo / Despesas</h3><div class="valor text-prejuizo">' + fmt(t.desp) + '</div><div class="sub">' + filtrados().filter(function(l){return l.tipo==='despesa';}).length + ' despesas</div></div>' +
    '<div class="card border-queda"><h3>🟠 Queda / Prejuízo do Período</h3><div class="valor text-queda">' + fmt(queda) + '</div><div class="sub">' + (t.resultado < 0 ? '⚠️ Consumo de reservas' : 'Sem queda no período') + '</div></div>' +
    '<div class="card border-duvida"><h3>🟡 Sem Comprovante / Obscuro</h3><div class="valor text-duvida">' + fmt(semComp.reduce(function(s,l) { return s + l.valor; }, 0)) + '</div><div class="sub">' + semComp.length + ' compra(s) — verificar</div></div>' +
    '<div class="card ' + (t.resultado >= 0 ? 'border-lucro' : 'border-prejuizo') + '"><h3>📊 Resultado</h3><div class="valor ' + (t.resultado >= 0 ? 'text-lucro' : 'text-prejuizo') + '">' + fmt(t.resultado) + '</div><div class="sub">' + (t.resultado >= 0 ? '✅ Superávit' : '⚠️ Déficit') + '</div></div>' +
    '<div class="card border-ok"><h3>🏦 Fundo de Reserva</h3><div class="valor text-ok">' + fmt(saldoReserva()) + '</div><div class="sub">Saldo disponível</div></div>';
}

/* ===== VEREDITO ===== */
function renderVeredito() {
  const t = totais();
  if (!t.total) { document.getElementById('veredito').innerHTML = ''; return; }
  const semComp = filtrados().filter(function(l) { return l.tipo === 'despesa' && l.comprovante === 'nao'; });
  let html = '';
  if (semComp.length) html += '<div class="veredito v-alerta">🟡 <b>ATENÇÃO:</b> <b>' + semComp.length + '</b> compra(s) <b>sem comprovante</b> somando <b>' + fmt(semComp.reduce(function(s,l) { return s + l.valor; }, 0)) + '</b>. Confira a aba <b>🧾 Controle de Comprovantes</b>.</div>';
  if (t.duvidosos.length) html += '<div class="veredito v-alerta">⚠️ <b>' + t.duvidosos.length + '</b> lançamento(s) com valores <b>acima de 2,5x a média</b> (' + fmt(t.valorDuvidoso) + '). Registre observações na tabela de análise.</div>';
  html += t.resultado >= 0
    ? '<div class="veredito v-receita">🟢 <b>Receita total:</b> ' + fmt(t.rec) + ' — o período fechou com <b class="text-receita">superávit de ' + fmt(t.resultado) + '</b>.</div>'
    : '<div class="veredito v-queda">🟠 <b>Queda/Prejuízo:</b> o período fechou com <b class="text-queda">déficit de ' + fmt(Math.abs(t.resultado)) + '</b>. Fundo de reserva consumido — requer atenção.</div>';
  document.getElementById('veredito').innerHTML = html;
}

/* ===== TABELA DETALHADA ===== */
function renderTabela() {
  const tb = document.querySelector('#tabela tbody');
  tb.innerHTML = '';
  const ls = filtrados();
  if (!ls.length) { tb.innerHTML = '<tr><td colspan="7" style="text-align:center;color:#64748b;">Nenhum lançamento. Carregue um arquivo ou ajuste os filtros.</td></tr>'; return; }
  ls.forEach(function(l) {
    const i = lancamentos.indexOf(l);
    const tr = document.createElement('tr');
    tr.style.borderLeft = '4px solid ' + corStatus(l.status);
    tr.innerHTML =
      '<td>' + l.descricao + '</td>' +
      '<td>' + (l.dia ? l.dia + '/' : '') + MESES[l.mes] + '/' + l.ano + '</td>' +
      '<td><span class="badge" style="background:' + (CORES_CONTA[l.rubrica] || '#64748b') + '33;color:' + (CORES_CONTA[l.rubrica] || '#e2e8f0') + ';">' + l.rubrica + '</span></td>' +
      '<td><span class="badge ' + (l.tipo === 'receita' ? 'b-receita' : 'b-prejuizo') + '">' + (l.tipo === 'receita' ? '🟢 Receita' : '🔴 Despesa') + '</span></td>' +
      '<td class="' + (l.tipo === 'receita' ? 'text-receita' : 'text-prejuizo') + '" style="font-weight:700;">' + (l.tipo === 'receita' ? '+' : '-') + ' ' + fmt(l.valor) + '</td>' +
      '<td><select class="status-sel" onchange="mudarStatus(' + i + ', this.value)">' +
        '<option value="ok"' + (l.status === 'ok' ? ' selected' : '') + '>✅ OK — Verificado</option>' +
        '<option value="duvida"' + (l.status === 'duvida' ? ' selected' : '') + '>🟡 Duvidoso — Analisar</option>' +
        '<option value="prejuizo"' + (l.status === 'prejuizo' ? ' selected' : '') + '>🔴 Prejuízo</option>' +
        '<option value="lucro"' + (l.status === 'lucro' ? ' selected' : '') + '>🔵 Lucro</option>' +
      '</select></td>' +
      '<td><input class="obs-input" placeholder="Digite observações..." value="' + (l.obs || '').replace(/"/g, '&quot;') + '" onchange="lancamentos[' + i + '].obs=this.value"></td>';
    tb.appendChild(tr);
  });
}
function corStatus(s) { return { ok:'#64748b', duvida:'#eab308', prejuizo:'#ef4444', lucro:'#3b82f6' }[s] || '#475569'; }
function mudarStatus(i, v) { lancamentos[i].status = v; renderAll(); }

/* ===== ABA DEMONSTRATIVO ===== */
function renderDemonstrativo() {
  const ls = filtrados();
  const rec = ls.filter(function(l) { return l.tipo === 'receita'; }).reduce(function(s,l) { return s + l.valor; }, 0);
  const desp = ls.filter(function(l) { return l.tipo === 'despesa'; }).reduce(function(s,l) { return s + l.valor; }, 0);
  const res = rec - desp;
  document.getElementById('demonstrativoResumo').innerHTML =
    '<div class="veredito v-receita">🟢 <b>RECEITA TOTAL (recebimentos):</b> <b style="font-size:1.2rem;color:#4ade80;">' + fmt(rec) + '</b></div>' +
    '<div class="veredito v-perda">🔴 <b>TOTAL DE PREJUÍZO (despesas):</b> <b style="font-size:1.2rem;color:#f87171;">' + fmt(desp) + '</b></div>' +
    (res >= 0
      ? '<div class="veredito v-receita">🟢 <b>Saldo positivo:</b> <b style="color:#4ade80;">+' + fmt(res) + '</b></div>'
      : '<div class="veredito v-queda">🟠 <b>Queda / Prejuízo:</b> <b style="color:#fb923c;">' + fmt(res) + '</b> — consumo das reservas</div>');
  const tb = document.querySelector('#tblDemo tbody');
  tb.innerHTML = '';
  if (!ls.length) { tb.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#64748b;">Sem dados.</td></tr>'; return; }
  ls.forEach(function(l) {
    let cor, classe, situacao;
    if (l.tipo === 'despesa' && (l.comprovante === 'nao' || l.status === 'duvida')) {
      cor = 'text-duvida'; classe = 'b-duvida';
      situacao = l.comprovante === 'nao' ? '🟡 Sem comprovante — obscuro' : '🟡 Valor obscuro — analisar';
    } else if (l.tipo === 'receita') {
      cor = 'text-lucro'; classe = 'b-lucro'; situacao = '🔵 Recebimento';
    } else {
      cor = 'text-prejuizo'; classe = 'b-prejuizo'; situacao = '🔴 Prejuízo / Despesa';
    }
    tb.innerHTML += '<tr style="border-left:4px solid ' + corStatus(l.status) + ';">' +
      '<td>' + (l.dia ? l.dia + '/' : '') + MESES[l.mes] + '/' + l.ano + '</td>' +
      '<td>' + l.descricao + '</td>' +
      '<td><span class="badge" style="background:' + (CORES_CONTA[l.rubrica] || '#64748b') + '33;color:' + (CORES_CONTA[l.rubrica] || '#e2e8f0') + ';">' + l.rubrica + '</span></td>' +
      '<td class="' + cor + '" style="font-weight:700;">' + (l.tipo === 'receita' ? '+' : '-') + ' ' + fmt(l.valor) + '</td>' +
      '<td class="' + cor + '" style="font-weight:700;">' + (l.tipo === 'receita' ? '🟢 Recebimento' : '🔴 Prejuízo') + '</td>' +
      '<td><span class="badge ' + classe + '">' + situacao + '</span></td></tr>';
  });
}

/* ===== ABA SALDOS ===== */
function renderSaldos() {
  if (!saldos.length) { document.getElementById('saldosResumo').innerHTML = '<p style="color:#64748b;">Carregue um arquivo com saldos por conta/rubrica.</p>'; return; }
  const totalAtual = saldos.reduce(function(s,x) { return s + x.atual; }, 0);
  const totalAnt = saldos.reduce(function(s,x) { return s + x.anterior; }, 0);
  const varTotal = totalAtual - totalAnt;
  document.getElementById('saldosResumo').innerHTML =
    '<div class="veredito v-ok">🏦 <b>Total disponível em todas as contas:</b> <b style="font-size:1.15rem;color:#e2e8f0;">' + fmt(totalAtual) + '</b></div>' +
    (varTotal >= 0
      ? '<div class="veredito v-receita">🟢 <b>Evolução vs mês anterior:</b> <b style="color:#4ade80;">+' + fmt(varTotal) + '</b> — os saldos <b>cresceram</b>.</div>'
      : '<div class="veredito v-queda">🟠 <b>Queda vs mês anterior:</b> <b style="color:#fb923c;">' + fmt(varTotal) + '</b> — os saldos <b>diminuíram</b>.</div>');
  const tb = document.querySelector('#tblSaldos tbody');
  tb.innerHTML = '';
  saldos.forEach(function(s) {
    const v = s.atual - s.anterior;
    const situ = v > 0 ? '<span class="badge b-receita">🟢 Cresceu</span>' : v < 0 ? '<span class="badge b-queda">🟠 Caiu</span>' : '<span class="badge b-ok">⚪ Estável</span>';
    tb.innerHTML += '<tr><td><b style="color:' + (CORES_CONTA[s.conta] || '#e2e8f0') + '">' + s.conta + '</b></td>' +
      '<td>' + fmt(s.anterior) + '</td><td style="font-weight:700;">' + fmt(s.atual) + '</td>' +
      '<td class="' + (v >= 0 ? 'text-receita' : 'text-queda') + '" style="font-weight:700;">' + (v >= 0 ? '+' : '') + fmt(v) + '</td>' +
      '<td>' + situ + '</td></tr>';
  });
  if (charts.saldosAba) charts.saldosAba.destroy();
  charts.saldosAba = new Chart(document.getElementById('gSaldosAba'), {
    type: 'bar',
    data: { labels: saldos.map(function(s) { return s.conta; }),
      datasets: [{ label: 'Saldo Atual (R$)', data: saldos.map(function(s) { return s.atual; }),
        backgroundColor: saldos.map(function(s) { return CORES_CONTA[s.conta] || '#64748b'; }), borderRadius: 8 }] },
    options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(c) { return fmt(c.raw); } } } },
      scales: { x: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } },
                y: { ticks: { color: '#e2e8f0' }, grid: { display: false } } } }
  });
  if (charts.saldosComp) charts.saldosComp.destroy();
  charts.saldosComp = new Chart(document.getElementById('gSaldosComp'), {
    type: 'bar',
    data: { labels: saldos.map(function(s) { return s.conta; }), datasets: [
      { label: 'Mês Anterior', data: saldos.map(function(s) { return s.anterior; }), backgroundColor: '#64748b', borderRadius: 6 },
      { label: 'Mês Atual', data: saldos.map(function(s) { return s.atual; }), backgroundColor: saldos.map(function(s) { return CORES_CONTA[s.conta] || '#3b82f6'; }), borderRadius: 6 }] },
    options: { responsive: true, maintainAspectRatio: false,
      plugins: { legend: { labels: { color: '#cbd5e1' } }, tooltip: { callbacks: { label: function(c) { return c.dataset.label + ': ' + fmt(c.raw); } } } },
      scales: { x: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } } } }
  });
}

/* ===== ABA COMPROVANTES ===== */
function renderComprovantes() {
  const ls = filtrados().filter(function(l) { return l.tipo === 'despesa'; });
  const ok = ls.filter(function(l) { return l.comprovante === 'sim'; }).length;
  const falta = ls.filter(function(l) { return l.comprovante === 'nao'; });
  const desconhecido = ls.filter(function(l) { return l.comprovante === 'desconhecido'; });
  document.getElementById('comprovantesResumo').innerHTML =
    '<div class="veredito v-receita">✅ <b>Com comprovante:</b> ' + ok + ' compra(s) — situação regular.</div>' +
    '<div class="veredito v-perda">🔴 <b>SEM comprovante:</b> ' + falta.length + ' compra(s) somando <b>' + fmt(falta.reduce(function(s,l) { return s + l.valor; }, 0)) + '</b> — <b>VERIFICAR COMPROVANTE!</b></div>' +
    (desconhecido.length ? '<div class="veredito v-alerta">🟡 <b>Não identificado:</b> ' + desconhecido.length + ' compra(s) — confirmar anexação do comprovante.</div>' : '');
  const tb = document.querySelector('#tblComp tbody');
  tb.innerHTML = '';
  if (!ls.length) { tb.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#64748b;">Sem despesas carregadas.</td></tr>'; return; }
  ls.forEach(function(l) {
    const tem = l.comprovante === 'sim';
    const acao = tem ? '<span class="com-ok">✅ Regular</span>'
      : l.comprovante === 'nao' ? '<span class="msg-verificar">🔴 VERIFICAR COMPROVANTE</span>'
      : '<span class="com-falta">🟡 Confirmar comprovante</span>';
    tb.innerHTML += '<tr style="border-left:4px solid ' + (tem ? '#22c55e' : l.comprovante === 'nao' ? '#ef4444' : '#eab308') + ';">' +
      '<td>' + l.descricao + '</td>' +
      '<td>' + (l.dia ? l.dia + '/' : '') + MESES[l.mes] + '/' + l.ano + '</td>' +
      '<td><span class="badge" style="background:' + (CORES_CONTA[l.rubrica] || '#64748b') + '33;color:' + (CORES_CONTA[l.rubrica] || '#e2e8f0') + ';">' + l.rubrica + '</span></td>' +
      '<td class="' + (tem ? 'text-receita' : 'text-prejuizo') + '" style="font-weight:700;">' + fmt(l.valor) + '</td>' +
      '<td>' + (tem ? '<span class="com-ok">✅ SIM</span>' : l.comprovante === 'nao' ? '<span class="com-falta">🔴 NÃO</span>' : '<span class="text-duvida">❓ Não identificado</span>') + '</td>' +
      '<td>' + acao + '</td></tr>';
  });
}

/* ===== ABA INVESTIMENTO ===== */
function renderInvestimento() {
  const fr = saldoReserva();
  if (fr <= 0) { document.getElementById('investResumo').innerHTML = '<div class="veredito v-perda">⚠️ Fundo de Reserva sem saldo positivo para projeção de aplicação.</div>'; return; }
  const t = totais();
  document.getElementById('investResumo').innerHTML =
    '<div class="veredito v-alerta">💡 <b>Diagnóstico:</b> o Fundo de Reserva de <b>' + fmt(fr) + '</b> está <b>parado (renda parada)</b> e perde valor para a inflação todos os meses. Abaixo, simulação de aplicações bancárias seguras para o dinheiro <b>render e não desvalorizar</b>.</div>' +
    '<div class="veredito v-ok">🎯 <b>Recomendação:</b> manter em reserva de liquidez imediata cerca de <b>' + fmt(Math.max(t.desp, 0) * 2) + '</b> (2 meses de despesas) e aplicar o excedente em <b>CDB 100%+ do CDI com liquidez diária</b> ou <b>Tesouro Selic</b>.</div>';
  const cdiMes = 0.0085;
  const aplicacoes = [
    { nome: '💳 Conta remunerada (100% CDI)', tx: cdiMes, liquidez: 'Diária', risco: 'Mínimo', ipca: 'Parcial', cor: '#3b82f6' },
    { nome: '🏦 CDB 100% CDI liquidez diária', tx: cdiMes, liquidez: 'Diária', risco: 'Baixo (FGC)', ipca: 'Parcial', cor: '#22c55e' },
    { nome: '📜 Tesouro Selic 2029', tx: cdiMes * 0.98, liquidez: 'D+1', risco: 'Baixíssimo', ipca: 'Parcial', cor: '#eab308' },
    { nome: '📈 CDB 110% CDI (12 meses)', tx: cdiMes * 1.10, liquidez: 'No vencimento', risco: 'Baixo (FGC)', ipca: 'Parcial', cor: '#a855f7' },
    { nome: '🛡️ Tesouro IPCA+ (proteção inflação)', tx: 0.0065, liquidez: 'No vencimento', risco: 'Baixíssimo', ipca: '✅ Total', cor: '#06b6d4' }
  ];
  document.getElementById('simCards').innerHTML = aplicacoes.map(function(a) {
    return '<div class="sim-card" style="border-top-color:' + a.cor + ';">' +
      '<h4>' + a.nome + '</h4>' +
      '<div class="rend" style="color:' + a.cor + ';">+' + fmt(fr * (Math.pow(1 + a.tx, 12) - 1)) + '</div>' +
      '<div class="tx">em 12 meses • taxa ' + (a.tx * 100).toFixed(2) + '% a.m.</div>' +
      '<div class="ganho">6m: ' + fmt(fr * (Math.pow(1 + a.tx, 6) - 1)) + '<br>24m: ' + fmt(fr * (Math.pow(1 + a.tx, 24) - 1)) + '<br>Mês 1: ' + fmt(fr * a.tx) + '</div></div>';
  }).join('');
  const tb = document.querySelector('#tblInv tbody');
  tb.innerHTML = '';
  aplicacoes.forEach(function(a) {
    tb.innerHTML += '<tr>' +
      '<td><b style="color:' + a.cor + '">' + a.nome + '</b></td>' +
      '<td class="text-receita">' + (a.tx * 100).toFixed(2) + '% a.m.</td>' +
      '<td class="text-receita">+' + fmt(fr * (Math.pow(1 + a.tx, 6) - 1)) + '</td>' +
      '<td class="text-receita" style="font-weight:700;">+' + fmt(fr * (Math.pow(1 + a.tx, 12) - 1)) + '</td>' +
      '<td class="text-receita">+' + fmt(fr * (Math.pow(1 + a.tx, 24) - 1)) + '</td>' +
      '<td>' + a.liquidez + '</td><td>' + a.risco + '</td><td>' + a.ipca + '</td></tr>';
  });
  document.getElementById('vencimentoBox').innerHTML =
    '📌 <b>Como o dinheiro para de desvalorizar:</b> aplicado a ~0,85% a.m. (CDI), o Fundo de Reserva de <b>' + fmt(fr) + '</b> rende cerca de <b style="color:#4ade80;">' + fmt(fr * cdiMes) + '/mês</b> — acima da inflação média (~0,4% a.m.). Em 12 meses, o ganho real protegido fica em torno de <b style="color:#4ade80;">' + fmt(fr * (Math.pow(1 + cdiMes, 12) - 1) - fr * 0.004 * 12) + '</b> acima da correção inflacionária. ⚠️ Prefira instituições com <b>garantia FGC</b> e liquidez diária para emergências do condomínio.';
}

/* ===== ABA ANÁLISE FINAL ===== */
function renderAnaliseFinal() {
  const ls = filtrados();
  const ok = ls.filter(function(l) { return l.status === 'ok'; }).reduce(function(s,l) { return s + l.valor; }, 0);
  const nOk = ls.filter(function(l) { return l.status === 'ok'; }).length;
  const duv = ls.filter(function(l) { return l.status === 'duvida'; }).reduce(function(s,l) { return s + l.valor; }, 0);
  const nDuv = ls.filter(function(l) { return l.status === 'duvida'; }).length;
  const prej = ls.filter(function(l) { return l.status === 'prejuizo'; }).reduce(function(s,l) { return s + l.valor; }, 0);
  const recTotal = ls.filter(function(l) { return l.tipo === 'receita'; }).reduce(function(s,l) { return s + l.valor; }, 0);
  const despTotal = ls.filter(function(l) { return l.tipo === 'despesa'; }).reduce(function(s,l) { return s + l.valor; }, 0);
  const semComp = ls.filter(function(l) { return l.tipo === 'despesa' && l.comprovante === 'nao'; }).reduce(function(s,l) { return s + l.valor; }, 0);
  document.getElementById('analiseFinalResumo').innerHTML =
    '<div class="veredito v-ok">✅ <b>Verificados OK:</b> ' + nOk + ' lançamento(s) — ' + fmt(ok) + '</div>' +
    '<div class="veredito v-alerta">🟡 <b>Duvidosos (precisam análise):</b> ' + nDuv + ' lançamento(s) — ' + fmt(duv) + '</div>' +
    '<div class="veredito v-perda">🔴 <b>Prejuízos confirmados:</b> ' + fmt(prej) + '</div>' +
    '<div class="veredito ' + (despTotal > recTotal ? 'v-perda' : 'v-receita') + '">📊 <b>Balanço geral:</b> Receitas ' + fmt(recTotal) + ' × Despesas ' + fmt(despTotal) + ' → ' +
      (recTotal >= despTotal ? '<b class="text-receita">Saldo positivo ' + fmt(recTotal - despTotal) + '</b>' : '<b class="text-prejuizo">Déficit ' + fmt(recTotal - despTotal) + '</b>') + '</div>';
  if (charts.final) charts.final.destroy();
  charts.final = new Chart(document.getElementById('gFinal'), {
    type: 'bar',
    data: { labels: ['Receitas','Despesas','Verificados OK','Duvidosos','Prejuízos','Sem Comprovante'],
      datasets: [{ label: 'Valor (R$)', data: [recTotal, despTotal, ok, duv, prej, semComp],
        backgroundColor: [CORES.receita, CORES.prejuizo, '#64748b', CORES.duvida, CORES.prejuizo, '#f97316'],
        borderRadius: 10 }] },
    options: { responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(c) { return fmt(c.raw); } } } },
      scales: { x: { ticks: { color: '#e2e8f0', font: { size: 11 } }, grid: { display: false } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } } } }
  });
}

/* ===== GRÁFICOS PRINCIPAIS ===== */
function renderGraficos() {
  ['rosca','linha','saldos'].forEach(function(k) { if (charts[k]) charts[k].destroy(); });
  const t = totais();
  if (!t.total) return;
  charts.rosca = new Chart(document.getElementById('gRosca'), {
    type: 'doughnut',
    data: { labels: ['🟢 Receitas','🔴 Despesas','🟡 Desvios Suspeitos'],
      datasets: [{ data: [t.rec, t.desp, t.valorDuvidoso],
        backgroundColor: [CORES.receita, CORES.prejuizo, CORES.duvida], borderWidth: 3, borderColor: '#1e293b' }] },
    options: { responsive: true, maintainAspectRatio: false, cutout: '62%',
      plugins: { legend: { position: 'bottom', labels: { color: '#cbd5e1', padding: 15 } },
        tooltip: { callbacks: { label: function(c) { return c.label + ': ' + fmt(c.raw); } } } } }
  });
  if (saldos.length) {
    charts.saldos = new Chart(document.getElementById('gSaldos'), {
      type: 'bar',
      data: { labels: saldos.map(function(s) { return s.conta; }),
        datasets: [{ data: saldos.map(function(s) { return s.atual; }), backgroundColor: saldos.map(function(s) { return CORES_CONTA[s.conta] || '#64748b'; }), borderRadius: 8 }] },
      options: { responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(c) { return fmt(c.raw); } } } },
        scales: { x: { ticks: { color: '#94a3b8', font: { size: 9 }, maxRotation: 45 }, grid: { color: '#334155' } },
                  y: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } } } }
    });
  }
  const ls = filtrados();
  const porMes = MESES.map(function(_, i) {
    return {
      rec: ls.filter(function(l) { return l.mes === i && l.tipo === 'receita'; }).reduce(function(s,l) { return s + l.valor; }, 0),
      desp: ls.filter(function(l) { return l.mes === i && l.tipo === 'despesa'; }).reduce(function(s,l) { return s + l.valor; }, 0)
    };
  });
  const ativos = [];
  porMes.forEach(function(m, i) { if (m.rec || m.desp) ativos.push(i); });
  const mesesL = ativos.map(function(i) { return MESES[i]; });
  const porMesA = ativos.map(function(i) { return porMes[i]; });
  let acum = 0;
  const fundo = porMesA.map(function(m) { acum += (m.rec - m.desp); return acum; });
  charts.linha = new Chart(document.getElementById('gLinha'), {
    type: 'line',
    data: { labels: mesesL.length ? mesesL : MESES,
      datasets: [
        { label: 'Fundo de Reserva Acumulado', data: fundo.length ? fundo : [0], borderColor: CORES.receita, backgroundColor: 'rgba(34,197,94,.15)', fill: true, tension: 0.35, borderWidth: 3, pointRadius: 5 },
        { label: 'Resultado Mensal', data: porMesA.map(function(m) { return m.rec - m.desp; }), borderColor: CORES.duvida, tension: 0.35, borderWidth: 2, borderDash: [6,4], pointRadius: 4 }] },
    options: { responsive: true, maintainAspectRatio: false,
      plugins: { legend: { labels: { color: '#cbd5e1' } } },
      scales: { x: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: '#334155' } } } }
  });
}

/* ===== ANEXOS ===== */
const anexoZone = document.getElementById('anexoZone');
const anexoInput = document.getElementById('anexoInput');
anexoZone.addEventListener('click', function() { anexoInput.click(); });
anexoInput.addEventListener('change', function(e) { [...e.target.files].forEach(addAnexo); e.target.value = ''; });
function addAnexo(f) {
  const icons = { pdf:'📕', xml:'🔗', xlsx:'📗', xls:'📗', csv:'📗', jpg:'🖼️', png:'🖼️', jpeg:'🖼️', doc:'📘', docx:'📘' };
  anexos.push({ nome: f.name, url: URL.createObjectURL(f), data: new Date().toLocaleString('pt-BR'), icon: icons[f.name.split('.').pop().toLowerCase()] || '📄' });
  renderAnexos();
}
function renderAnexos() {
  document.getElementById('anexosLista').innerHTML = anexos.length ? anexos.map(function(a) {
    return '<div class="anexo-card"><span class="arq-icon">' + a.icon + '</span>' +
      '<div><div class="nome">' + a.nome + '</div><div class="data">📅 ' + a.data + '</div>' +
      '<a href="' + a.url + '" target="_blank">Visualizar arquivo →</a></div></div>';
  }).join('') : '<p style="color:#64748b;">Nenhum anexo ainda.</p>';
}

/* ===== TABS / RENDER ===== */
function abrirTab(id, btn) {
  document.querySelectorAll('.tab-content, .tab-btn').forEach(function(el) { el.classList.remove('active'); });
  document.getElementById(id).classList.add('active');
  btn.classList.add('active');
}
function renderAll() {
  renderCards();
  renderVeredito();
  renderTabela();
  renderGraficos();
  renderDemonstrativo();
  renderSaldos();
  renderComprovantes();
  renderInvestimento();
  renderAnaliseFinal();
}
</script>
</body>
</html>
