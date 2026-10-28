<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Pré-indução de hangar · materiais</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --bg:#E9EDF1; --surface:#FFFFFF; --surface-2:#F5F7F9; --ink:#15202B; --muted:#5A6775; --line:#D2D9E0;
  --navy:#1D3D5E; --disp:#23845A; --atende:#C99318; --nao:#C23A2F; --focus:#2F6FB3;
  --font:"Barlow",system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;
  --cond:"Barlow Condensed","Arial Narrow",var(--font);
  color-scheme:light;
  box-sizing:border-box;
  padding-top:env(safe-area-inset-top,0px); padding-bottom:env(safe-area-inset-bottom,0px);
}
@media (prefers-color-scheme:dark){
  :root:not([data-theme="light"]){
    --bg:#0F1720; --surface:#16212C; --surface-2:#1C2935; --ink:#E6ECF1; --muted:#97A6B4; --line:#2A3A49;
    --navy:#8DB4DA; --disp:#3BAF7C; --atende:#E0AE3A; --nao:#E5594D; --focus:#7FB0E6; color-scheme:dark;
  }
}
:root[data-theme="dark"]{
  --bg:#0F1720; --surface:#16212C; --surface-2:#1C2935; --ink:#E6ECF1; --muted:#97A6B4; --line:#2A3A49;
  --navy:#8DB4DA; --disp:#3BAF7C; --atende:#E0AE3A; --nao:#E5594D; --focus:#7FB0E6; color-scheme:dark;
}
*,*::before,*::after{box-sizing:inherit}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--font);font-size:15px;line-height:1.45;font-variant-numeric:tabular-nums}
button,input,select{font:inherit;color:inherit}
:focus-visible{outline:2px solid var(--focus);outline-offset:2px}
.wrap{max-width:1320px;margin:0 auto;padding:20px 20px 48px}

/* topo */
.top{display:flex;flex-wrap:wrap;align-items:flex-end;justify-content:space-between;gap:16px;padding-bottom:18px;border-bottom:2px solid var(--ink)}
.top h1{font-family:var(--cond);font-weight:700;font-size:clamp(28px,4vw,40px);line-height:1;margin:0;letter-spacing:.01em}
.top .sub{color:var(--muted);margin-top:6px}
.actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.btn{border:1px solid var(--line);background:var(--surface);padding:8px 14px;border-radius:6px;cursor:pointer;font-weight:600}
.btn:hover{border-color:var(--muted)}
.btn.primary{background:var(--navy);border-color:var(--navy);color:var(--surface)}
.btn:disabled{opacity:.5;cursor:default}

/* estado vazio */
.drop{margin-top:28px;border:2px dashed var(--line);border-radius:10px;background:var(--surface);padding:56px 24px;text-align:center;cursor:pointer}
.drop.over{border-color:var(--navy);background:var(--surface-2)}
.drop h2{font-family:var(--cond);font-size:26px;margin:0 0 8px}
.drop p{margin:4px auto;color:var(--muted);max-width:60ch}
.status-msg{margin-top:14px;font-weight:600}
.status-msg.err{color:var(--nao)}

/* resumo */
.summary{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:0;margin-top:20px;background:var(--surface);border:1px solid var(--line);border-radius:8px}
.summary div{padding:14px 18px;border-left:1px solid var(--line)}
.summary div:first-child{border-left:0}
.summary b{display:block;font-family:var(--cond);font-size:32px;line-height:1.05;font-weight:700}
.summary span{color:var(--muted);font-size:13px}
.summary .k-nao b{color:var(--nao)}
.summary .k-disp b{color:var(--disp)}

/* filtros */
.filters{display:flex;flex-wrap:wrap;gap:10px;margin-top:18px;align-items:flex-end}
.filters label{display:flex;flex-direction:column;font-size:12px;color:var(--muted);gap:3px}
.filters select,.filters input{background:var(--surface);border:1px solid var(--line);border-radius:6px;padding:7px 9px;min-width:150px}
.filters input{min-width:220px}

/* layout principal */
.grid{display:grid;grid-template-columns:minmax(0,1.7fr) minmax(0,1fr);gap:20px;margin-top:20px}
.panel{background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:16px 18px}
.panel h3{font-family:var(--cond);font-size:21px;margin:0 0 2px;font-weight:600}
.panel .hint{color:var(--muted);font-size:13px;margin:0 0 12px}
.legend{display:flex;gap:14px;font-size:13px;color:var(--muted);margin-bottom:10px;flex-wrap:wrap}
.legend i{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:5px;vertical-align:-1px}

/* fila de entradas no hangar */
.queue{display:flex;flex-direction:column}
.ev{display:grid;grid-template-columns:62px 72px minmax(0,1fr) 56px;gap:10px;align-items:center;padding:7px 6px;border-top:1px solid var(--line);cursor:pointer;background:none;border-left:0;border-right:0;border-bottom:0;text-align:left;width:100%}
.ev:hover{background:var(--surface-2)}
.ev.sel{background:var(--surface-2);box-shadow:inset 3px 0 0 var(--navy)}
.ev .reg{font-family:var(--cond);font-weight:700;font-size:20px;line-height:1}
.ev .reg small{display:block;font-family:var(--font);font-weight:500;font-size:11px;color:var(--muted);margin-top:2px}
.ev .when{font-size:13px;line-height:1.2}
.ev .when em{display:block;font-style:normal;color:var(--muted);font-size:12px}
.ev .when.soon em{color:var(--nao);font-weight:600}
.bar{display:flex;height:16px;border-radius:3px;overflow:hidden;background:var(--surface-2)}
.bar span{display:block;height:100%}
.ev .pct{font-family:var(--cond);font-weight:700;font-size:19px;text-align:right}
.ev .meta{grid-column:3/5;font-size:12px;color:var(--muted);margin-top:-4px}

/* barras simples */
.hbars{display:flex;flex-direction:column;gap:8px}
.hb{display:grid;grid-template-columns:120px minmax(0,1fr) 54px;gap:10px;align-items:center;font-size:13px}
.hb .track{height:12px;background:var(--surface-2);border-radius:3px;overflow:hidden}
.hb .track span{display:block;height:100%;background:var(--nao)}
.hb b{text-align:right;font-weight:600}
.side{display:flex;flex-direction:column;gap:20px}

/* tabela */
.tablebox{margin-top:20px}
.tablehead{display:flex;justify-content:space-between;align-items:flex-end;gap:12px;flex-wrap:wrap;margin-bottom:10px}
.scroll{overflow-x:auto;border:1px solid var(--line);border-radius:6px}
table{border-collapse:collapse;width:100%;font-size:13px;min-width:1050px}
th,td{padding:7px 9px;text-align:left;border-bottom:1px solid var(--line);vertical-align:top}
th{background:var(--surface-2);font-weight:600;cursor:pointer;white-space:nowrap;position:sticky;top:0}
th[aria-sort="ascending"]::after{content:" ▲";font-size:10px}
th[aria-sort="descending"]::after{content:" ▼";font-size:10px}
td.num{text-align:right}
td.desc{max-width:260px}
td.obs{max-width:280px;color:var(--muted)}
.tag{display:inline-block;padding:1px 7px;border-radius:10px;font-size:12px;font-weight:600;white-space:nowrap;color:#fff}
.tag.disp{background:var(--disp)} .tag.atende{background:var(--atende);color:#2a1f00} .tag.nao{background:var(--nao)}
.pager{display:flex;gap:8px;align-items:center;justify-content:flex-end;margin-top:10px;font-size:13px;color:var(--muted)}

.overlay{position:fixed;inset:0;background:rgba(10,18,26,.55);display:none;align-items:center;justify-content:center;z-index:20}
.overlay.on{display:flex}
.overlay div{background:var(--surface);padding:22px 28px;border-radius:8px;font-weight:600;max-width:90vw}
.hidden{display:none!important}
.note{color:var(--muted);font-size:12px;margin-top:22px}

@media (max-width:980px){
  .grid{grid-template-columns:1fr}
  .summary{grid-template-columns:repeat(2,minmax(0,1fr))}
  .summary div{border-left:0;border-top:1px solid var(--line)}
  .summary div:nth-child(-n+2){border-top:0}
}
@media (max-width:560px){
  .ev{grid-template-columns:54px 64px minmax(0,1fr) 48px;gap:8px}
  .hb{grid-template-columns:96px minmax(0,1fr) 44px}
  .filters select,.filters input{min-width:0;width:100%}
  .filters label{flex:1 1 140px}
}
@media (prefers-reduced-motion:no-preference){ .bar span,.hb .track span{transition:width .35s ease} }
.badge{display:inline-block;font-size:12px;font-weight:600;padding:2px 8px;border-radius:10px;border:1px solid var(--line);color:var(--muted);margin-left:6px;vertical-align:1px}
.badge.cloud{border-color:var(--disp);color:var(--disp)}
.banner{margin-top:14px;padding:10px 14px;border:1px solid var(--line);border-left:4px solid var(--atende);background:var(--surface);border-radius:6px;font-size:14px}
.banner button{margin-left:10px}
.tag.gone{background:var(--muted)}
td.hist{font-size:12px;color:var(--muted);white-space:nowrap}
.hist-list{display:flex;flex-direction:column;font-size:13px}
.hist-list div{display:grid;grid-template-columns:86px minmax(0,1fr) auto;gap:8px;padding:6px 0;border-top:1px solid var(--line)}
.hist-list div span:nth-child(2){color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.trend{width:100%;height:120px;display:block;margin:4px 0 6px}
.trend .l1{stroke:var(--disp);fill:none;stroke-width:2}
.trend .l2{stroke:var(--nao);fill:none;stroke-width:2}
.trend text{fill:var(--muted);font-size:10px;font-family:var(--font)}
.trend line{stroke:var(--line)}
.modal{position:fixed;inset:0;background:rgba(10,18,26,.55);display:none;align-items:center;justify-content:center;z-index:30;padding:16px}
.modal.on{display:flex}
.modal .box{background:var(--surface);border-radius:10px;padding:22px 24px;max-width:520px;width:100%}
.modal h3{font-family:var(--cond);font-size:24px;margin:0 0 6px}
.modal dl{display:grid;grid-template-columns:1fr auto;gap:6px 16px;margin:14px 0 18px}
.modal dt{color:var(--muted)} .modal dd{margin:0;font-weight:700;text-align:right}
.modal .row{display:flex;gap:8px;justify-content:flex-end}
.modal p{margin:6px 0;color:var(--muted);font-size:14px}
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <div>
      <h1>Pré-indução de hangar</h1>
      <div class="sub"><span id="fileInfo">Nenhuma carga salva ainda</span><span class="badge" id="storeBadge">Conectando…</span></div>
    </div>
    <div class="actions">
      <button class="btn" id="themeBtn" type="button">Tema</button>
      <button class="btn primary" id="loadBtn" type="button">Carregar planilha do dia</button>
      <input type="file" id="fileInput" accept=".xlsx,.xlsm" class="hidden">
    </div>
  </header>
  <div class="banner hidden" id="newBanner">Uma nova carga foi salva por outra pessoa.<button class="btn" id="reloadBtn" type="button">Atualizar painel</button></div>
  <div class="banner hidden" id="localBanner">Este painel está aberto fora do Claude, então as cargas ficam guardadas só neste navegador. Para a equipe compartilhar o mesmo histórico, use o link publicado.</div>

  <section id="empty" class="drop" tabindex="0" aria-label="Área para soltar a planilha">
    <h2>Arraste a planilha de pré-indução para cá</h2>
    <p>Ou clique para escolher o arquivo. O painel lê a aba <strong>Caminho Crítico</strong>, compara com o histórico salvo e guarda a nova carga sem apagar as anteriores.</p>
    <p>Planilhas grandes podem levar de 10 a 30 segundos para abrir.</p>
    <div class="status-msg" id="emptyMsg"></div>
  </section>

  <main id="dash" class="hidden">
    <section class="summary" aria-label="Resumo">
      <div><b id="kLinhas">–</b><span>linhas ativas</span></div>
      <div class="k-disp"><b id="kDisp">–</b><span>material disponível</span></div>
      <div><b id="kAtend">–</b><span>atendido (disponível + chega a tempo)</span></div>
      <div class="k-nao"><b id="kNao">–</b><span>linhas que não atendem</span></div>
      <div><b id="kEventos">–</b><span>entradas no hangar</span></div>
    </section>

    <section class="filters" aria-label="Filtros">
      <label>Evento<select id="fEvento"></select></label>
      <label>Responsável<select id="fResp"></select></label>
      <label>Marco temporal<select id="fMarco"></select></label>
      <label>Atividade<select id="fAtiv"></select></label>
      <label>Na última carga<select id="fMud">
        <option value="">Todas as linhas</option>
        <option value="mudou">Mudaram de status</option>
        <option value="nova">Entraram agora</option>
      </select></label>
      <label>Buscar PN, descrição ou WO<input id="fBusca" type="search" placeholder="Ex.: SEALANT, B00130"></label>
      <button class="btn" id="resetBtn" type="button">Limpar filtros</button>
    </section>

    <div class="grid">
      <section class="panel">
        <h3>Fila de entradas no hangar</h3>
        <p class="hint">Ordenada pela data de programação. Clique em um evento para filtrar o painel.</p>
        <div class="legend">
          <span><i style="background:var(--disp)"></i>Disponível</span>
          <span><i style="background:var(--atende)"></i>Chega a tempo</span>
          <span><i style="background:var(--nao)"></i>Não atende</span>
        </div>
        <div class="queue" id="queue"></div>
      </section>
      <div class="side">
        <section class="panel">
          <h3>Evolução das cargas</h3>
          <p class="hint">% disponível (verde) e linhas que não atendem (vermelho) a cada carga.</p>
          <svg class="trend" id="trend" viewBox="0 0 400 120" preserveAspectRatio="none" role="img" aria-label="Evolução das cargas"></svg>
          <div class="hist-list" id="histList"></div>
        </section>
        <section class="panel">
          <h3>Não atende por responsável</h3>
          <p class="hint">Quem precisa agir sobre as linhas críticas.</p>
          <div class="hbars" id="byResp"></div>
        </section>
        <section class="panel">
          <h3>Não atende por marco temporal</h3>
          <p class="hint">Quanto mais perto de T0, mais tarde a demanda entrou.</p>
          <div class="hbars" id="byMarco"></div>
        </section>
      </div>
    </div>

    <section class="panel tablebox">
      <div class="tablehead">
        <div>
          <h3>Linhas de demanda</h3>
          <p class="hint" id="tableCount" style="margin:0"></p>
        </div>
        <div class="filters" style="margin:0">
          <label>Status<select id="fStatus">
            <option value="nao">Não atende</option>
            <option value="atende">Chega a tempo</option>
            <option value="disp">Disponível</option>
            <option value="ativas">Todas as ativas</option>
            <option value="gone">Saíram da planilha</option>
          </select></label>
          <button class="btn" id="csvBtn" type="button">Exportar CSV</button>
        </div>
      </div>
      <div class="scroll">
        <table>
          <thead><tr id="thead"></tr></thead>
          <tbody id="tbody"></tbody>
        </table>
      </div>
      <div class="pager">
        <button class="btn" id="prevBtn" type="button">Anterior</button>
        <span id="pageInfo"></span>
        <button class="btn" id="nextBtn" type="button">Próxima</button>
      </div>
    </section>
    <p class="note" id="note"></p>
  </main>
</div>

<div class="overlay" id="overlay" role="status" aria-live="polite"><div id="overlayMsg">Lendo planilha…</div></div>

<div class="modal" id="confirm" role="dialog" aria-modal="true" aria-labelledby="cfTitle">
  <div class="box">
    <h3 id="cfTitle">Salvar a carga de hoje?</h3>
    <p id="cfFile"></p>
    <dl id="cfStats"></dl>
    <p>Linhas que saíram da planilha não são apagadas: ficam guardadas com a data em que saíram.</p>
    <div class="row" style="margin-top:16px">
      <button class="btn" id="cfCancel" type="button">Cancelar</button>
      <button class="btn primary" id="cfSave" type="button">Salvar carga</button>
    </div>
  </div>
</div>

<script>
(function(){
"use strict";
var SHEET = "Caminho Crítico";
var PAGE = 100, DAY = 86400000, CHUNK_BYTES = 190000;
var $ = function(id){ return document.getElementById(id); };

var state = { rows: [], cargas: [], meta: null, gen: null, sortKey: "nec", sortDir: 1, page: 0, tableRows: [] };
var store = null, downloads = null, pending = null;

/* ---------- utilidades ---------- */
function norm(s){ return String(s == null ? "" : s).normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s+/g," ").trim().toLowerCase(); }
function isSerialDate(v){ return typeof v === "number" && v > 30000 && v < 80000; }
function serialToDay(v){ return Math.floor(v) - 25569; }            /* dias desde 01/01/1970 */
function toDay(v){ return isSerialDate(v) ? serialToDay(v) : null; }
function todayDay(){ var n = new Date(); return Math.floor(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) / DAY); }
function fmtDay(d, short){
  if (d == null) return "";
  var x = new Date(d * DAY), dd = String(x.getUTCDate()).padStart(2,"0"), mm = String(x.getUTCMonth()+1).padStart(2,"0");
  return short ? dd + "/" + mm : dd + "/" + mm + "/" + x.getUTCFullYear();
}
function pct(a,b){ return b ? a/b*100 : 0; }
function fmtPct(a,b){ return b ? pct(a,b).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1}) + "%" : "–"; }
function fmtInt(n){ return Number(n || 0).toLocaleString("pt-BR"); }
function esc(s){ return String(s == null ? "" : s).replace(/[&<>"]/g, function(c){ return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]; }); }
function plural(n, a, b){ return fmtInt(n) + " " + (n === 1 ? a : b); }
function statusOf(v){
  var s = norm(v);
  if (s.charAt(0) === "1" || s.indexOf("disponivel") >= 0) return "d";
  if (s.charAt(0) === "2" || s === "atende") return "a";
  return "n";
}
var ST = { d: "disp", a: "atende", n: "nao" };
var STATUS_LABEL = { d: "Disponível", a: "Chega a tempo", n: "Não atende", g: "Saiu da planilha" };
var STATUS_SHORT = { d: "Disp.", a: "Chega", n: "Não at." };
function marcoOrder(m){ var x = /^T(\d+)$/i.exec(m || ""); return x ? -parseInt(x[1],10) : 1; }
function sleep(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }

/* ---------- leitura da planilha ---------- */
var COLS = {
  prog: ["programacao"], marco: ["demanda"], wo: ["wo"], acft: ["acft"], pn: ["pn"], desc: ["description","descricao"],
  qty: ["qty","quantidade"], nec: ["data necessidade"], prev: ["previsao na cia"], atend: ["atendimento"],
  resp: ["responsavel"], obs: ["observacoes"], order: ["order no. (oh_orderno)"], pacote: ["pacote"],
  ativ: ["atividade"], acao: ["acao confirmed"]
};
function parseWorkbook(buf){
  var wb = XLSX.read(buf, { type: "array", sheets: [SHEET], dense: true, cellDates: false, cellFormula: false, cellHTML: false, cellText: false, cellStyles: false });
  var ws = wb.Sheets[SHEET];
  if (!ws) { var alt = wb.SheetNames.filter(function(n){ return norm(n) === norm(SHEET); })[0]; ws = alt ? wb.Sheets[alt] : null; }
  if (!ws) throw new Error("A aba “" + SHEET + "” não foi encontrada. Confira se é a planilha de pré-indução.");
  var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  var h = -1;
  for (var i = 0; i < Math.min(15, aoa.length); i++) { var hs = (aoa[i] || []).map(norm); if (hs.indexOf("atendimento") >= 0 && hs.indexOf("pacote") >= 0) { h = i; break; } }
  if (h < 0) throw new Error("Não encontrei o cabeçalho da aba “" + SHEET + "” (colunas Atendimento e Pacote).");
  var head = aoa[h].map(norm), map = {};
  Object.keys(COLS).forEach(function(k){
    var idx = -1; COLS[k].some(function(n){ idx = head.indexOf(n); return idx >= 0; });
    if (idx < 0 && k === "order") idx = head.findIndex(function(x){ return x.indexOf("order no") === 0; });
    map[k] = idx;
  });
  ["pacote","wo","pn","atend"].forEach(function(k){ if (map[k] < 0) throw new Error("Coluna obrigatória não encontrada: " + COLS[k][0]); });
  var str = function(v){ return v == null ? "" : String(v).trim(); };
  var out = [], seen = {};
  for (var r = h + 1; r < aoa.length; r++) {
    var a = aoa[r]; if (!a) continue;
    var g = function(k){ return map[k] >= 0 ? a[map[k]] : null; };
    if (g("pn") == null && g("pacote") == null) continue;
    var prev = g("prev");
    var row = {
      pacote: str(g("pacote")), acft: str(g("acft")), wo: str(g("wo")), pn: str(g("pn")), desc: str(g("desc")),
      qty: typeof g("qty") === "number" ? g("qty") : str(g("qty")), nec: toDay(g("nec")), prog: toDay(g("prog")),
      marco: str(g("marco")) || "Sem marco", prev: isSerialDate(prev) ? fmtDay(serialToDay(prev)) : str(prev),
      st: statusOf(g("atend")), resp: str(g("resp")) || "Sem responsável", obs: str(g("obs")).slice(0, 300),
      order: str(g("order")), ativ: str(g("ativ")).replace(/^[:\s]+/, ""), acao: str(g("acao"))
    };
    var base = row.pacote + "|" + row.wo + "|" + row.pn;
    seen[base] = (seen[base] || 0) + 1;
    row.k = seen[base] > 1 ? base + "#" + seen[base] : base;
    out.push(row);
  }
  if (!out.length) throw new Error("A aba “" + SHEET + "” está vazia.");
  return out;
}

/* ---------- mesclagem com o histórico ---------- */
var FIELDS = ["pacote","acft","wo","pn","desc","qty","nec","prog","marco","prev","st","resp","obs","order","ativ","acao"];
function merge(oldRows, newRows, day){
  var byKey = {}, res = [], stats = { lidas: newRows.length, novas: 0, mudaram: 0, voltaram: 0, sairam: 0, iguais: 0 };
  oldRows.forEach(function(r){ byKey[r.k] = r; });
  var inNew = {};
  newRows.forEach(function(n){
    inNew[n.k] = 1;
    var o = byKey[n.k], r = { k: n.k };
    FIELDS.forEach(function(f){ r[f] = n[f]; });
    if (!o) {
      r.first = day; r.last = day; r.gone = null; r.chg = day; r.hist = [[day, n.st]];
      stats.novas++;
    } else {
      r.first = o.first; r.last = day; r.gone = null; r.hist = (o.hist || []).slice(); r.chg = o.chg;
      if (o.gone != null) { stats.voltaram++; }
      var lastSt = r.hist.length ? r.hist[r.hist.length-1][1] : null;
      if (lastSt !== n.st) {
        if (r.hist.length && r.hist[r.hist.length-1][0] === day) r.hist[r.hist.length-1] = [day, n.st];
        else r.hist.push([day, n.st]);
        r.chg = day; stats.mudaram++;
      } else if (o.gone == null) stats.iguais++;
    }
    res.push(r);
  });
  oldRows.forEach(function(o){
    if (inNew[o.k]) return;
    var r = {}; Object.keys(o).forEach(function(f){ r[f] = o[f]; });
    if (r.gone == null) { r.gone = day; stats.sairam++; }
    res.push(r);
  });
  return { rows: res, stats: stats };
}
function summarize(rows, day, file, stats){
  var s = { ts: Date.now(), dia: day, arquivo: file, ativas: 0, d: 0, a: 0, n: 0, ev: {} };
  rows.forEach(function(r){
    if (r.gone != null) return;
    s.ativas++; s[r.st]++;
    var e = s.ev[r.pacote] = s.ev[r.pacote] || [0,0,0];
    e[r.st === "d" ? 0 : r.st === "a" ? 1 : 2]++;
  });
  ["lidas","novas","mudaram","voltaram","sairam"].forEach(function(k){ s[k] = stats[k]; });
  return s;
}

/* ---------- formato compacto para gravar ---------- */
var PACK = ["k"].concat(FIELDS).concat(["first","last","gone","chg","hist"]);
function pack(r){ return PACK.map(function(f){ return r[f] === undefined ? null : r[f]; }); }
function unpack(a){ var r = {}; PACK.forEach(function(f, i){ r[f] = a[i]; }); if (!r.hist) r.hist = []; return r; }
function chunk(rows){
  var sorted = rows.slice().sort(function(a,b){ return a.pacote < b.pacote ? -1 : a.pacote > b.pacote ? 1 : 0; });
  var chunks = [], cur = [], size = 0;
  sorted.forEach(function(r){
    var p = pack(r), len = JSON.stringify(p).length * 1.15 + 4;
    if (size + len > CHUNK_BYTES && cur.length) { chunks.push(cur); cur = []; size = 0; }
    cur.push(p); size += len;
  });
  if (cur.length) chunks.push(cur);
  return chunks;
}

/* ---------- armazenamento: nuvem (db do artefato) ---------- */
function retry(fn){
  var tries = 0;
  return (function go(){
    return fn().catch(function(e){
      tries++;
      var code = e && e.code;
      if (tries < 4 && (code === "unavailable" || code === "resource_exhausted")) return sleep(800 * tries + Math.random() * 400).then(go);
      throw e;
    });
  })();
}
function cloudStore(db){
  return {
    kind: "cloud",
    load: function(){
      return retry(function(){ return db.doc("estado/atual").get(); }).then(function(snap){
        if (!snap.exists) return { meta: null, rows: [], cargas: [] };
        var meta = snap.data(), ids = [];
        for (var i = 0; i < meta.n; i++) ids.push(meta.gen + "-" + i);
        var rows = [], idx = 0;
        function next(){
          var batch = ids.slice(idx, idx + 6); idx += 6;
          if (!batch.length) return Promise.resolve();
          return Promise.all(batch.map(function(id){ return retry(function(){ return db.doc("linhas/" + id).get(); }); })).then(function(snaps){
            snaps.forEach(function(sn){ if (sn.exists) (sn.data().rows || []).forEach(function(a){ rows.push(unpack(a)); }); });
            return next();
          });
        }
        return next().then(function(){
          return retry(function(){ return db.collection("cargas").orderBy("ts","desc").limit(120).get(); });
        }).then(function(q){
          return { meta: meta, rows: rows, cargas: q.docs.map(function(d){ return d.data(); }).reverse() };
        });
      });
    },
    currentGen: function(){ return retry(function(){ return db.doc("estado/atual").get(); }).then(function(s){ return s.exists ? s.data().gen : null; }); },
    save: function(rows, carga, progress){
      var gen = "g" + Date.now(), chunks = chunk(rows), i = 0;
      function writeNext(){
        if (i >= chunks.length) return Promise.resolve();
        progress("Salvando parte " + (i + 1) + " de " + chunks.length + "…");
        var id = gen + "-" + i, body = { gen: gen, i: i, rows: chunks[i] };
        i++;
        return retry(function(){ return db.doc("linhas/" + id).set(body); }).then(writeNext);
      }
      return writeNext().then(function(){
        progress("Registrando a carga…");
        return retry(function(){ return db.doc("cargas/c" + carga.ts).set(carga); });
      }).then(function(){
        return retry(function(){ return db.doc("estado/atual").set({ gen: gen, n: chunks.length, arquivo: carga.arquivo, ts: carga.ts, dia: carga.dia }); });
      }).then(function(){
        progress("Limpando versões antigas…");
        return retry(function(){ return db.collection("linhas").where("gen","!=",gen).limit(1000).get(); }).then(function(q){
          return q.docs.reduce(function(p, d){ return p.then(function(){ return retry(function(){ return db.doc("linhas/" + d.id).delete(); }); }); }, Promise.resolve());
        }).catch(function(){ /* limpeza é opcional */ });
      }).then(function(){ return gen; });
    },
    watch: function(cb){ return db.doc("estado/atual").onSnapshot(function(s){ cb(s.exists ? s.data().gen : null); }, function(){}); }
  };
}

/* ---------- armazenamento: só neste navegador ---------- */
function localStore(){
  function open(){ return new Promise(function(res, rej){ try { var rq = indexedDB.open("preinducao-historico", 1); rq.onupgradeneeded = function(){ rq.result.createObjectStore("s"); }; rq.onsuccess = function(){ res(rq.result); }; rq.onerror = function(){ rej(rq.error); }; } catch(e){ rej(e); } }); }
  function get(){ return open().then(function(db){ return new Promise(function(res){ try { var rq = db.transaction("s").objectStore("s").get("estado"); rq.onsuccess = function(){ res(rq.result || null); }; rq.onerror = function(){ res(null); }; } catch(e){ res(null); } }); }).catch(function(){ return null; }); }
  return {
    kind: "local",
    load: function(){ return get().then(function(v){ return v ? { meta: v.meta, rows: v.rows.map(unpack), cargas: v.cargas || [] } : { meta: null, rows: [], cargas: [] }; }); },
    currentGen: function(){ return get().then(function(v){ return v && v.meta ? v.meta.gen : null; }); },
    save: function(rows, carga, progress){
      progress("Salvando neste navegador…");
      return get().then(function(v){
        var cargas = ((v && v.cargas) || []).concat([carga]);
        var gen = "g" + Date.now();
        var payload = { meta: { gen: gen, arquivo: carga.arquivo, ts: carga.ts, dia: carga.dia }, rows: rows.map(pack), cargas: cargas };
        return open().then(function(db){ return new Promise(function(res, rej){ var tx = db.transaction("s","readwrite"); tx.objectStore("s").put(payload, "estado"); tx.oncomplete = function(){ res(gen); }; tx.onerror = function(){ rej(tx.error); }; }); });
      });
    },
    watch: function(){ return function(){}; }
  };
}

/* ---------- inicialização ---------- */
function initStore(){
  var c = window.claude;
  if (!c || typeof c.use !== "function") return Promise.resolve(localStore());
  c.use("downloads").then(function(d){ downloads = d; }).catch(function(){});
  return c.use("db").then(function(db){ return db ? cloudStore(db) : localStore(); }).catch(function(){ return localStore(); });
}
function refreshFromStore(){
  showOverlay("Carregando histórico…");
  return store.load().then(function(data){
    hideOverlay();
    state.meta = data.meta; state.gen = data.meta ? data.meta.gen : null; state.rows = data.rows; state.cargas = data.cargas;
    $("newBanner").classList.add("hidden");
    if (state.rows.length) showDash(); else showEmpty("");
  }).catch(function(e){
    hideOverlay();
    showEmpty("Não foi possível carregar o histórico salvo (" + ((e && (e.message || e.code)) || "erro") + "). Tente recarregar a página.");
  });
}
initStore().then(function(s){
  store = s;
  var b = $("storeBadge");
  if (s.kind === "cloud") { b.textContent = "Histórico salvo na nuvem"; b.classList.add("cloud"); }
  else { b.textContent = "Histórico só neste navegador"; $("localBanner").classList.remove("hidden"); }
  refreshFromStore().then(function(){
    var first = true;
    s.watch(function(gen){
      if (first) { first = false; return; }
      if (gen && gen !== state.gen) $("newBanner").classList.remove("hidden");
    });
  });
});

/* ---------- carregar arquivo ---------- */
function showOverlay(msg){ $("overlayMsg").textContent = msg; $("overlay").classList.add("on"); }
function hideOverlay(){ $("overlay").classList.remove("on"); }
function showEmpty(msg){
  $("dash").classList.add("hidden"); $("empty").classList.remove("hidden");
  $("emptyMsg").className = "status-msg" + (msg ? " err" : ""); $("emptyMsg").textContent = msg || "";
}
function alertMsg(msg){ if (state.rows.length) { window.alert(msg); } else showEmpty(msg); }

function handleFile(file){
  if (!file) return;
  if (!store) { alertMsg("O painel ainda está carregando. Aguarde alguns segundos e tente de novo."); return; }
  if (typeof XLSX === "undefined") { alertMsg("A biblioteca de leitura de Excel não carregou. Verifique a conexão e recarregue a página."); return; }
  showOverlay("Lendo “" + file.name + "”…");
  var reader = new FileReader();
  reader.onerror = function(){ hideOverlay(); alertMsg("Não foi possível ler o arquivo."); };
  reader.onload = function(e){
    setTimeout(function(){
      var newRows;
      try { newRows = parseWorkbook(new Uint8Array(e.target.result)); }
      catch(err){ hideOverlay(); alertMsg(err.message || String(err)); return; }
      hideOverlay();
      var day = todayDay(), m = merge(state.rows, newRows, day);
      pending = { file: file.name, day: day, newRows: newRows, merged: m };
      openConfirm();
    }, 50);
  };
  reader.readAsArrayBuffer(file);
}
function openConfirm(){
  var s = pending.merged.stats;
  $("cfFile").textContent = pending.file + " · carga de " + fmtDay(pending.day);
  $("cfStats").innerHTML = [
    ["Linhas lidas", s.lidas], ["Linhas novas", s.novas], ["Mudaram de status", s.mudaram],
    ["Voltaram para a planilha", s.voltaram], ["Saíram da planilha", s.sairam]
  ].map(function(x){ return "<dt>" + x[0] + "</dt><dd>" + fmtInt(x[1]) + "</dd>"; }).join("");
  $("confirm").classList.add("on");
  $("cfSave").focus();
}
function closeConfirm(){ $("confirm").classList.remove("on"); }

function saveCarga(){
  closeConfirm();
  var p = pending; pending = null;
  showOverlay("Conferindo se houve outra carga…");
  store.currentGen().then(function(gen){
    if (gen === state.gen) return p.merged;
    /* alguém salvou depois que o painel abriu: recarrega e mescla de novo para não perder nada */
    return store.load().then(function(data){ state.rows = data.rows; state.cargas = data.cargas; state.gen = gen; return merge(data.rows, p.newRows, p.day); });
  }).then(function(m){
    var carga = summarize(m.rows, p.day, p.file, m.stats);
    return store.save(m.rows, carga, showOverlay).then(function(gen){
      state.rows = m.rows; state.gen = gen; state.cargas = state.cargas.concat([carga]);
      state.meta = { gen: gen, arquivo: p.file, ts: carga.ts, dia: p.day };
      hideOverlay(); state.page = 0; showDash();
    });
  }).catch(function(e){
    hideOverlay();
    var code = e && e.code;
    var msg = code === "invalid_argument" || code === "not_granted" ? "Você não tem permissão para salvar cargas neste painel. Peça ao dono do painel acesso de colaborador ou editor."
      : code === "quota_exceeded" ? "O espaço de armazenamento do painel está cheio. Fale com o dono do painel para arquivar eventos antigos."
      : "A carga não foi salva (" + ((e && (e.message || e.code)) || "erro desconhecido") + "). Nada foi alterado no histórico; tente de novo.";
    alertMsg(msg);
  });
}

/* ---------- montar painel ---------- */
function uniq(rows, k){ var s = {}; rows.forEach(function(r){ if (r[k]) s[r[k]] = 1; }); return Object.keys(s); }
function mostCommon(arr){ var c = {}, best = "", n = 0; arr.forEach(function(x){ c[x] = (c[x]||0)+1; if (c[x] > n){ n = c[x]; best = x; } }); return best; }
function fillSelect(id, opts){
  var el = $(id), keep = el.value;
  el.innerHTML = opts.map(function(o){ return '<option value="' + esc(o[0]) + '">' + esc(o[1]) + "</option>"; }).join("");
  if (opts.some(function(o){ return o[0] === keep; })) el.value = keep;
}
function lastCargaDay(){ return state.cargas.length ? state.cargas[state.cargas.length-1].dia : null; }

function showDash(){
  var active = state.rows.filter(function(r){ return r.gone == null; });
  var byPac = {};
  active.forEach(function(r){ (byPac[r.pacote] = byPac[r.pacote] || []).push(r); });
  state.events = Object.keys(byPac).map(function(p){
    var rs = byPac[p];
    var prog = rs.reduce(function(m,r){ return r.prog != null && (m == null || r.prog < m) ? r.prog : m; }, null);
    return { pacote: p, prog: prog, acft: mostCommon(rs.map(function(r){ return r.acft; })), ativ: mostCommon(rs.map(function(r){ return r.ativ; })) };
  }).sort(function(a,b){ return (a.prog||0) - (b.prog||0) || a.pacote.localeCompare(b.pacote); });

  fillSelect("fEvento", [["","Todos os eventos"]].concat(state.events.map(function(e){ return [e.pacote, e.acft + " · " + fmtDay(e.prog) + " · " + e.pacote]; })));
  fillSelect("fResp", [["","Todos"]].concat(uniq(state.rows,"resp").sort().map(function(x){ return [x,x]; })));
  fillSelect("fMarco", [["","Todos"]].concat(uniq(state.rows,"marco").sort(function(a,b){ return marcoOrder(a) - marcoOrder(b); }).map(function(x){ return [x,x]; })));
  fillSelect("fAtiv", [["","Todas"]].concat(uniq(state.rows,"ativ").sort().map(function(x){ return [x,x]; })));

  var m = state.meta;
  $("fileInfo").textContent = m ? "Última carga: " + m.arquivo + " · " + new Date(m.ts).toLocaleString("pt-BR") : "Nenhuma carga salva ainda";
  var gone = state.rows.length - active.length;
  $("note").textContent = "Fonte: aba “" + SHEET + "”. " + plural(state.cargas.length, "carga salva", "cargas salvas") + ", " + plural(active.length, "linha ativa", "linhas ativas") +
    " e " + plural(gone, "linha que já saiu da planilha", "linhas que já saíram da planilha") + " guardadas no histórico.";
  $("empty").classList.add("hidden"); $("dash").classList.remove("hidden");
  render();
}

function filters(){ return { ev: $("fEvento").value, resp: $("fResp").value, marco: $("fMarco").value, ativ: $("fAtiv").value, mud: $("fMud").value, q: norm($("fBusca").value), st: $("fStatus").value }; }
function match(r, f, ignoreEvent){
  if (!ignoreEvent && f.ev && r.pacote !== f.ev) return false;
  if (f.resp && r.resp !== f.resp) return false;
  if (f.marco && r.marco !== f.marco) return false;
  if (f.ativ && r.ativ !== f.ativ) return false;
  if (f.mud) {
    var d = lastCargaDay();
    if (f.mud === "mudou" && !(r.chg === d && r.first !== d && r.hist.length > 1)) return false;
    if (f.mud === "nova" && r.first !== d) return false;
  }
  if (f.q && norm(r.pn + " " + r.desc + " " + r.wo + " " + r.order).indexOf(f.q) < 0) return false;
  return true;
}

function render(){
  var f = filters();
  var all = state.rows.filter(function(r){ return match(r, f, false); });
  var base = all.filter(function(r){ return r.gone == null; });
  var c = { d: 0, a: 0, n: 0 }; base.forEach(function(r){ c[r.st]++; });
  $("kLinhas").textContent = fmtInt(base.length);
  $("kDisp").textContent = fmtPct(c.d, base.length);
  $("kAtend").textContent = fmtPct(c.d + c.a, base.length);
  $("kNao").textContent = fmtInt(c.n);
  $("kEventos").textContent = fmtInt(uniq(base, "pacote").length);
  renderQueue(f);
  var nao = base.filter(function(r){ return r.st === "n"; });
  renderBars("byResp", nao, "resp", function(a,b){ return b[1]-a[1]; });
  renderBars("byMarco", nao, "marco", function(a,b){ return marcoOrder(a[0]) - marcoOrder(b[0]); });
  renderHistory();
  renderTable(all, f);
}

function renderQueue(f){
  var agg = {};
  state.rows.forEach(function(r){ if (r.gone != null || !match(r, f, true)) return; var a = agg[r.pacote] = agg[r.pacote] || { d:0, a:0, n:0, t:0 }; a[r.st]++; a.t++; });
  var today = todayDay();
  $("queue").innerHTML = state.events.filter(function(e){ return agg[e.pacote]; }).map(function(e){
    var a = agg[e.pacote], days = e.prog != null ? e.prog - today : null;
    var when = days == null ? "" : days < 0 ? "em andamento" : days === 0 ? "hoje" : "em " + plural(days, "dia", "dias");
    var soon = days != null && days <= 30 && a.n > 0, pd = pct(a.d,a.t), pa = pct(a.a,a.t), pn = pct(a.n,a.t);
    return '<button type="button" class="ev' + (f.ev === e.pacote ? " sel" : "") + '" data-p="' + esc(e.pacote) + '" aria-pressed="' + (f.ev === e.pacote) + '">' +
      '<span class="reg">' + esc(e.acft) + "<small>" + esc(e.pacote.split("/")[1] || "") + "</small></span>" +
      '<span class="when' + (soon ? " soon" : "") + '">' + fmtDay(e.prog, true) + "<em>" + when + "</em></span>" +
      '<span class="bar" title="Disponível ' + a.d + " · Chega a tempo " + a.a + " · Não atende " + a.n + '"><span style="width:' + pd + '%;background:var(--disp)"></span><span style="width:' + pa + '%;background:var(--atende)"></span><span style="width:' + pn + '%;background:var(--nao)"></span></span>' +
      '<span class="pct">' + Math.round(pd + pa) + "%</span>" +
      '<span class="meta">' + esc(e.ativ) + " · " + plural(a.t, "linha", "linhas") + " · " + (a.n ? '<strong style="color:var(--nao)">' + plural(a.n, "não atende", "não atendem") + "</strong>" : "tudo atendido") + "</span></button>";
  }).join("") || '<p class="hint">Nenhum evento com os filtros atuais.</p>';
}

function renderBars(id, rows, key, sorter){
  var c = {}; rows.forEach(function(r){ c[r[key]] = (c[r[key]]||0)+1; });
  var arr = Object.keys(c).map(function(k){ return [k, c[k]]; }).sort(sorter);
  var max = arr.reduce(function(m,x){ return Math.max(m,x[1]); }, 0);
  $(id).innerHTML = arr.length ? arr.map(function(x){ return '<div class="hb"><span>' + esc(x[0]) + '</span><span class="track"><span style="width:' + pct(x[1],max) + '%"></span></span><b>' + fmtInt(x[1]) + "</b></div>"; }).join("")
    : '<p class="hint">Nenhuma linha crítica com os filtros atuais.</p>';
}

function renderHistory(){
  var cs = state.cargas.slice(-30), svg = $("trend");
  if (cs.length < 2) {
    svg.innerHTML = '<text x="200" y="64" text-anchor="middle">O gráfico aparece a partir da segunda carga.</text>';
  } else {
    var W = 400, H = 120, L = 6, R = 6, T = 10, B = 18;
    var x = function(i){ return L + i * (W - L - R) / (cs.length - 1); };
    var pd = cs.map(function(c){ return pct(c.d, c.ativas); });
    var mn = Math.min.apply(null, pd), mx = Math.max.apply(null, pd); if (mx - mn < 2) { mn -= 1; mx += 1; }
    var nn = cs.map(function(c){ return c.n; }), nmax = Math.max.apply(null, nn) || 1;
    var y1 = function(v){ return T + (1 - (v - mn) / (mx - mn)) * (H - T - B); };
    var y2 = function(v){ return T + (1 - v / nmax) * (H - T - B); };
    var path = function(vals, y){ return vals.map(function(v,i){ return (i ? "L" : "M") + x(i).toFixed(1) + " " + y(v).toFixed(1); }).join(" "); };
    svg.innerHTML = '<line x1="0" x2="400" y1="' + (H - B) + '" y2="' + (H - B) + '"/>' +
      '<path class="l1" d="' + path(pd, y1) + '"/><path class="l2" d="' + path(nn, y2) + '"/>' +
      '<text x="' + L + '" y="' + (H - 4) + '">' + fmtDay(cs[0].dia, true) + '</text><text x="' + (W - R) + '" y="' + (H - 4) + '" text-anchor="end">' + fmtDay(cs[cs.length-1].dia, true) + "</text>";
  }
  $("histList").innerHTML = state.cargas.slice(-8).reverse().map(function(c){
    var det = [c.novas ? plural(c.novas, "nova", "novas") : "", c.mudaram ? c.mudaram + " mudaram" : "", c.sairam ? c.sairam + " saíram" : ""].filter(Boolean).join(" · ") || "sem mudanças";
    return "<div><span>" + fmtDay(c.dia) + '</span><span title="' + esc(c.arquivo) + '">' + det + "</span><b>" + fmtPct(c.d, c.ativas) + "</b></div>";
  }).join("");
}

var TCOLS = [["acft","Acft"],["pacote","Pacote"],["nec","Necessidade"],["pn","PN"],["desc","Descrição"],["qty","Qtd"],["st","Status"],["resp","Responsável"],["prev","Previsão"],["order","Pedido"],["marco","Marco"],["first","Desde"],["hist","Histórico de status"],["obs","Observações"]];
function histText(r){
  var h = r.hist.map(function(x){ return fmtDay(x[0], true) + " " + STATUS_SHORT[x[1]]; });
  if (r.gone != null) h.push(fmtDay(r.gone, true) + " saiu");
  return h.join(" → ");
}
function renderTable(all, f){
  var rows = all.filter(function(r){
    if (f.st === "gone") return r.gone != null;
    if (r.gone != null) return false;
    if (f.st === "ativas") return true;
    return ST[r.st] === f.st;
  });
  var k = state.sortKey, d = state.sortDir;
  rows.sort(function(a,b){
    var x = k === "hist" ? a.hist.length : a[k], y = k === "hist" ? b.hist.length : b[k];
    if (x == null || x === "") return 1; if (y == null || y === "") return -1;
    if (typeof x === "number" && typeof y === "number") return (x - y) * d;
    return String(x).localeCompare(String(y), "pt-BR", { numeric: true }) * d;
  });
  state.tableRows = rows;
  var pages = Math.max(1, Math.ceil(rows.length / PAGE)); if (state.page >= pages) state.page = pages - 1;
  $("thead").innerHTML = TCOLS.map(function(c){ return '<th scope="col" tabindex="0" data-k="' + c[0] + '" aria-sort="' + (c[0] === k ? (d > 0 ? "ascending" : "descending") : "none") + '">' + c[1] + "</th>"; }).join("");
  $("tbody").innerHTML = rows.slice(state.page * PAGE, state.page * PAGE + PAGE).map(function(r){
    var stc = r.gone != null ? "gone" : ST[r.st], stl = r.gone != null ? "Saiu em " + fmtDay(r.gone, true) : STATUS_LABEL[r.st];
    return "<tr><td>" + esc(r.acft) + "</td><td>" + esc(r.pacote) + "</td><td>" + fmtDay(r.nec) + "</td><td>" + esc(r.pn) + '</td><td class="desc">' + esc(r.desc) +
      '</td><td class="num">' + esc(typeof r.qty === "number" ? r.qty.toLocaleString("pt-BR") : r.qty) + '</td><td><span class="tag ' + stc + '">' + stl + "</span></td><td>" + esc(r.resp) +
      "</td><td>" + esc(r.prev) + "</td><td>" + esc(r.order) + "</td><td>" + esc(r.marco) + "</td><td>" + fmtDay(r.first, true) + '</td><td class="hist">' + esc(histText(r)) + '</td><td class="obs">' + esc(r.obs) + "</td></tr>";
  }).join("") || '<tr><td colspan="14">Nenhuma linha com os filtros atuais.</td></tr>';
  $("tableCount").textContent = plural(rows.length, "linha", "linhas");
  $("pageInfo").textContent = "Página " + (state.page + 1) + " de " + pages;
  $("prevBtn").disabled = state.page === 0; $("nextBtn").disabled = state.page >= pages - 1;
}

function exportCSV(){
  var cell = function(v){ v = v == null ? "" : String(v); return /[;"\n]/.test(v) ? '"' + v.replace(/"/g,'""') + '"' : v; };
  var lines = [TCOLS.map(function(c){ return c[1]; }).concat(["Saiu da planilha em"]).join(";")];
  state.tableRows.forEach(function(r){
    lines.push([r.acft, r.pacote, fmtDay(r.nec), r.pn, r.desc, typeof r.qty === "number" ? String(r.qty).replace(".", ",") : r.qty, STATUS_LABEL[r.st], r.resp, r.prev, r.order, r.marco, fmtDay(r.first), histText(r), r.obs, fmtDay(r.gone)].map(cell).join(";"));
  });
  var text = "\ufeff" + lines.join("\r\n"), name = "pre-inducao-linhas-" + new Date().toISOString().slice(0,10) + ".csv";
  if (downloads) { downloads.save({ filename: name, data: new Blob([text], { type: "text/csv;charset=utf-8" }) }).catch(function(){}); return; }
  var a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([text], { type: "text/csv;charset=utf-8" })); a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(function(){ URL.revokeObjectURL(a.href); }, 1000);
}

/* ---------- eventos da interface ---------- */
$("loadBtn").addEventListener("click", function(){ $("fileInput").click(); });
$("empty").addEventListener("click", function(){ $("fileInput").click(); });
$("empty").addEventListener("keydown", function(e){ if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("fileInput").click(); } });
$("fileInput").addEventListener("change", function(e){ handleFile(e.target.files[0]); e.target.value = ""; });
["dragenter","dragover"].forEach(function(t){ document.addEventListener(t, function(e){ e.preventDefault(); $("empty").classList.add("over"); }); });
["dragleave","drop"].forEach(function(t){ document.addEventListener(t, function(e){ e.preventDefault(); $("empty").classList.remove("over"); }); });
document.addEventListener("drop", function(e){ var f = e.dataTransfer && e.dataTransfer.files[0]; if (f) handleFile(f); });
$("cfCancel").addEventListener("click", function(){ pending = null; closeConfirm(); });
$("cfSave").addEventListener("click", saveCarga);
document.addEventListener("keydown", function(e){ if (e.key === "Escape" && $("confirm").classList.contains("on")) { pending = null; closeConfirm(); } });
$("reloadBtn").addEventListener("click", refreshFromStore);
["fEvento","fResp","fMarco","fAtiv","fMud","fStatus"].forEach(function(id){ $(id).addEventListener("change", function(){ state.page = 0; render(); }); });
var tmr; $("fBusca").addEventListener("input", function(){ clearTimeout(tmr); tmr = setTimeout(function(){ state.page = 0; render(); }, 200); });
$("resetBtn").addEventListener("click", function(){ ["fEvento","fResp","fMarco","fAtiv","fMud"].forEach(function(id){ $(id).value = ""; }); $("fBusca").value = ""; state.page = 0; render(); });
$("queue").addEventListener("click", function(e){ var b = e.target.closest(".ev"); if (!b) return; var p = b.getAttribute("data-p"); $("fEvento").value = $("fEvento").value === p ? "" : p; state.page = 0; render(); });
function sortBy(k){ if (state.sortKey === k) state.sortDir *= -1; else { state.sortKey = k; state.sortDir = 1; } render(); }
$("thead").addEventListener("click", function(e){ var th = e.target.closest("th"); if (th) sortBy(th.getAttribute("data-k")); });
$("thead").addEventListener("keydown", function(e){ if (e.key === "Enter") { var th = e.target.closest("th"); if (th) sortBy(th.getAttribute("data-k")); } });
$("prevBtn").addEventListener("click", function(){ state.page--; render(); });
$("nextBtn").addEventListener("click", function(){ state.page++; render(); });
$("csvBtn").addEventListener("click", exportCSV);
$("themeBtn").addEventListener("click", function(){
  var cur = document.documentElement.getAttribute("data-theme");
  var dark = cur ? cur === "dark" : window.matchMedia("(prefers-color-scheme: dark)").matches;
  var next = dark ? "light" : "dark"; document.documentElement.setAttribute("data-theme", next);
  try { localStorage.setItem("preinducao-tema", next); } catch(e){}
});
try { var t = localStorage.getItem("preinducao-tema"); if (t) document.documentElement.setAttribute("data-theme", t); } catch(e){}

window.__preInducao = { merge: merge, parseWorkbook: parseWorkbook, state: state };
})();
</script>
</body>

</html>
