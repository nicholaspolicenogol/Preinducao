<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Pré-indução de hangar</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --bg:#E9EDF1; --surface:#FFFFFF; --surface-2:#F4F6F8; --ink:#15202B; --muted:#586574; --line:#D2D9E0;
  --navy:#1D3D5E; --disp:#23845A; --atende:#C08A12; --nao:#C23A2F; --focus:#2F6FB3; --zone:#E3ECF5;
  --font:"Barlow",system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;
  --cond:"Barlow Condensed","Arial Narrow",var(--font);
  color-scheme:light; box-sizing:border-box;
  padding-top:env(safe-area-inset-top,0px); padding-bottom:env(safe-area-inset-bottom,0px);
}
@media (prefers-color-scheme:dark){
  :root:not([data-theme="light"]){
    --bg:#0F1720; --surface:#16212C; --surface-2:#1C2935; --ink:#E6ECF1; --muted:#9AA9B7; --line:#2A3A49;
    --navy:#8DB4DA; --disp:#3BAF7C; --atende:#E0AE3A; --nao:#E5594D; --focus:#7FB0E6; --zone:#1B2C3D; color-scheme:dark;
  }
}
:root[data-theme="dark"]{
  --bg:#0F1720; --surface:#16212C; --surface-2:#1C2935; --ink:#E6ECF1; --muted:#9AA9B7; --line:#2A3A49;
  --navy:#8DB4DA; --disp:#3BAF7C; --atende:#E0AE3A; --nao:#E5594D; --focus:#7FB0E6; --zone:#1B2C3D; color-scheme:dark;
}
*,*::before,*::after{box-sizing:inherit}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--font);font-size:15px;line-height:1.45;font-variant-numeric:tabular-nums}
button,input,select{font:inherit;color:inherit}
:focus-visible{outline:2px solid var(--focus);outline-offset:2px}
.wrap{max-width:1360px;margin:0 auto;padding:20px 20px 56px}
.hidden{display:none!important}

/* topo */
.top{display:flex;flex-wrap:wrap;align-items:flex-end;justify-content:space-between;gap:16px;padding-bottom:14px}
.top h1{font-family:var(--cond);font-weight:700;font-size:clamp(28px,4vw,42px);line-height:1;margin:0}
.top .sub{color:var(--muted);margin-top:6px;font-size:14px}
.actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.btn{border:1px solid var(--line);background:var(--surface);padding:8px 14px;border-radius:6px;cursor:pointer;font-weight:600}
.btn:hover{border-color:var(--muted)}
.btn.primary{background:var(--navy);border-color:var(--navy);color:var(--surface)}
.btn.small{padding:5px 10px;font-size:13px}
.btn:disabled{opacity:.5;cursor:default}
.menu{position:relative}
.menu-list{position:absolute;right:0;top:calc(100% + 6px);background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:6px;min-width:230px;z-index:15;box-shadow:0 8px 24px rgba(15,23,32,.15)}
.menu-list button{display:block;width:100%;text-align:left;background:none;border:0;padding:8px 10px;border-radius:5px;cursor:pointer}
.menu-list button:hover{background:var(--surface-2)}

/* abas */
.tabs{display:flex;gap:4px;border-bottom:2px solid var(--ink);overflow-x:auto}
.tabs a{font-family:var(--cond);font-size:19px;font-weight:600;color:var(--muted);text-decoration:none;padding:8px 14px 7px;border-radius:6px 6px 0 0;white-space:nowrap}
.tabs a:hover{color:var(--ink)}
.tabs a[aria-current="page"]{background:var(--ink);color:var(--bg)}

/* fontes de dados */
.sources{display:flex;flex-wrap:wrap;gap:10px;margin-top:14px}
.src{flex:1 1 300px;display:flex;gap:12px;align-items:center;background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:10px 14px;font-size:14px}
.src .dot{width:10px;height:10px;border-radius:50%;background:var(--line);flex:none}
.src.ok .dot{background:var(--disp)}
.src.old .dot{background:var(--atende)}
.src b{display:block;font-weight:600}
.src span{color:var(--muted)}
.src .btn{margin-left:auto;flex:none}

/* estado vazio */
.drop{margin-top:22px;border:2px dashed var(--line);border-radius:10px;background:var(--surface);padding:48px 24px;text-align:center;cursor:pointer}
.drop.over{border-color:var(--navy);background:var(--surface-2)}
.drop h2{font-family:var(--cond);font-size:28px;margin:0 0 8px}
.drop p{margin:6px auto;color:var(--muted);max-width:62ch}
.msg{margin-top:14px;font-weight:600}
.msg.err{color:var(--nao)}

/* filtros */
.filters{display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;margin-top:16px}
.filters label{display:flex;flex-direction:column;font-size:12px;color:var(--muted);gap:3px}
.filters select,.filters input{background:var(--surface);border:1px solid var(--line);border-radius:6px;padding:7px 9px;min-width:150px}
.filters input[type=search]{min-width:220px}

/* indicadores */
.kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));margin-top:16px;background:var(--surface);border:1px solid var(--line);border-radius:8px}
.kpis>div{padding:13px 16px;border-left:1px solid var(--line)}
.kpis>div:first-child{border-left:0}
.kpis b{display:block;font-family:var(--cond);font-size:31px;line-height:1.05;font-weight:700}
.kpis span{color:var(--muted);font-size:13px;line-height:1.3;display:block}
.kpis .good b{color:var(--disp)} .kpis .bad b{color:var(--nao)} .kpis .warn b{color:var(--atende)}

/* painéis */
.grid2{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:18px;margin-top:18px}
.grid21{display:grid;grid-template-columns:minmax(0,1.6fr) minmax(0,1fr);gap:18px;margin-top:18px}
.panel{background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:16px 18px;min-width:0}
.panel.full{margin-top:18px}
.panel h3{font-family:var(--cond);font-size:22px;margin:0;font-weight:600;line-height:1.15}
.panel .hint{color:var(--muted);font-size:13px;margin:3px 0 12px}
.panel-head{display:flex;justify-content:space-between;align-items:flex-end;gap:12px;flex-wrap:wrap}
.panel-head .filters{margin-top:0}
.empty{color:var(--muted);font-size:14px;padding:8px 0}

/* linha do tempo das induções (destaque) */
.tl{position:relative}
.tl-axis,.tl-row{display:grid;grid-template-columns:150px minmax(0,1fr) 64px;gap:12px;align-items:center}
.tl-axis{font-size:12px;color:var(--muted);padding-bottom:4px}
.tl-scale{position:relative;height:30px}
.tl-scale span{position:absolute;transform:translateX(-50%);white-space:nowrap;bottom:0;line-height:1.1;text-align:center}
.tl-scale span b{display:block;font-family:var(--cond);font-size:14px;color:var(--ink)}
.tl-body{position:relative}
.tl-guides{position:absolute;left:162px;right:76px;top:0;bottom:0;pointer-events:none}
.tl-guides .zone{position:absolute;top:0;bottom:0;left:0;background:var(--zone);border-right:2px solid var(--navy)}
.tl-guides i{position:absolute;top:0;bottom:0;border-left:1px dashed var(--line)}
.tl-row{position:relative;width:100%;background:none;border:0;border-top:1px solid var(--line);padding:5px 0;cursor:pointer;text-align:left}
.tl-row:hover .tl-name{text-decoration:underline}
.tl-row.sel{box-shadow:inset 3px 0 0 var(--navy)}
.tl-name{font-family:var(--cond);font-size:18px;font-weight:700;line-height:1.05;padding-left:8px}
.tl-name small{display:block;font-family:var(--font);font-weight:500;font-size:11.5px;color:var(--muted)}
.tl-track{position:relative;height:26px}
.tl-pill{position:absolute;top:3px;height:20px;min-width:20px;transform:translateX(-50%);border-radius:10px;color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;padding:0 7px;white-space:nowrap}
.tl-pill.ok{background:var(--disp)} .tl-pill.mid{background:var(--atende);color:#2a1f00} .tl-pill.bad{background:var(--nao)}
.tl-pill.past{border-radius:4px}
.tl-pct{font-family:var(--cond);font-weight:700;font-size:18px;text-align:right}
.tl-legend{display:flex;flex-wrap:wrap;gap:16px;font-size:13px;color:var(--muted);margin-bottom:8px}
.tl-legend i{display:inline-block;width:12px;height:12px;border-radius:6px;margin-right:6px;vertical-align:-2px}
.tl-legend .z{border-radius:2px;background:var(--zone);border-right:2px solid var(--navy)}

/* barras */
.hbars{display:flex;flex-direction:column;gap:8px}
.hb{display:grid;grid-template-columns:minmax(90px,150px) minmax(0,1fr) 60px;gap:10px;align-items:center;font-size:13px}
.hb>span:first-child{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.hb .track{height:12px;background:var(--surface-2);border-radius:3px;overflow:hidden;display:flex}
.hb .track span{display:block;height:100%}
.hb b{text-align:right;font-weight:600}
button.hb{background:none;border:0;padding:2px 0;cursor:pointer;text-align:left;width:100%}
button.hb:hover>span:first-child{text-decoration:underline}

/* lista de alertas */
.alerts{display:flex;flex-direction:column}
.alert{display:grid;grid-template-columns:64px minmax(0,1fr) auto;gap:10px;align-items:center;padding:8px 0;border-top:1px solid var(--line);font-size:14px;background:none;border-left:0;border-right:0;border-bottom:0;text-align:left;cursor:pointer;width:100%}
.alert:first-child{border-top:0}
.alert b.reg{font-family:var(--cond);font-size:19px}
.alert .why{color:var(--muted);font-size:13px}
.alert .n{font-family:var(--cond);font-size:20px;font-weight:700;color:var(--nao)}
.marco{display:inline-block;font-size:11.5px;font-weight:700;padding:1px 6px;border-radius:4px;background:var(--surface-2);border:1px solid var(--line);color:var(--ink);margin-right:6px}

/* gráfico de tendência */
.trend{width:100%;height:150px;display:block}
.trend .l1{stroke:var(--disp);fill:none;stroke-width:2}
.trend .l2{stroke:var(--nao);fill:none;stroke-width:2}
.trend text{fill:var(--muted);font-size:10px;font-family:var(--font)}
.trend line{stroke:var(--line)}
.trend-legend{display:flex;gap:16px;font-size:13px;color:var(--muted)}
.trend-legend i{display:inline-block;width:14px;height:3px;margin-right:6px;vertical-align:3px}

/* colunas por semana */
.cols{display:flex;align-items:flex-end;gap:8px;height:150px;padding-top:8px}
.cols div{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;font-size:12px;color:var(--muted);min-width:0}
.cols div span.bar{width:100%;max-width:44px;background:var(--navy);border-radius:3px 3px 0 0;min-height:2px}
.cols div b{font-size:12px;color:var(--ink);margin-bottom:3px}
.cols div em{font-style:normal;margin-top:4px;white-space:nowrap}

/* tabelas */
.scroll{overflow-x:auto;border:1px solid var(--line);border-radius:6px}
table{border-collapse:collapse;width:100%;font-size:13px}
table.wide{min-width:1100px}
th,td{padding:7px 9px;text-align:left;border-bottom:1px solid var(--line);vertical-align:top}
th{background:var(--surface-2);font-weight:600;white-space:nowrap}
th[data-k]{cursor:pointer}
th[aria-sort="ascending"]::after{content:" ▲";font-size:10px}
th[aria-sort="descending"]::after{content:" ▼";font-size:10px}
td.num,th.num{text-align:right}
td.desc{max-width:240px}
td.obs{max-width:260px;color:var(--muted)}
td.hist{font-size:12px;color:var(--muted);white-space:nowrap}
td.late{color:var(--nao);font-weight:700}
tr.click{cursor:pointer}
tr.click:hover td{background:var(--surface-2)}
.tag{display:inline-block;padding:1px 7px;border-radius:10px;font-size:12px;font-weight:600;white-space:nowrap;color:#fff}
.tag.disp{background:var(--disp)} .tag.atende{background:var(--atende);color:#2a1f00} .tag.nao{background:var(--nao)} .tag.gone{background:var(--muted)}
.tag.line{background:none;border:1px solid var(--line);color:var(--ink)}
.pager{display:flex;gap:8px;align-items:center;justify-content:flex-end;margin-top:10px;font-size:13px;color:var(--muted)}
.stockcell{font-size:12px;line-height:1.5}
.stockcell b{font-weight:700}

/* aviso, modal, carregando */
.banner{margin-top:12px;padding:9px 14px;border:1px solid var(--line);border-left:4px solid var(--atende);background:var(--surface);border-radius:6px;font-size:14px}
.overlay{position:fixed;inset:0;background:rgba(10,18,26,.55);display:none;align-items:center;justify-content:center;z-index:40}
.overlay.on{display:flex}
.overlay div{background:var(--surface);padding:22px 28px;border-radius:8px;font-weight:600;max-width:90vw}
.modal{position:fixed;inset:0;background:rgba(10,18,26,.55);display:none;align-items:center;justify-content:center;z-index:30;padding:16px}
.modal.on{display:flex}
.modal .box{background:var(--surface);border-radius:10px;padding:22px 24px;max-width:520px;width:100%}
.modal h3{font-family:var(--cond);font-size:24px;margin:0 0 6px}
.modal dl{display:grid;grid-template-columns:1fr auto;gap:6px 16px;margin:14px 0 14px}
.modal dt{color:var(--muted)} .modal dd{margin:0;font-weight:700;text-align:right}
.modal p{margin:6px 0;color:var(--muted);font-size:14px}
.modal .row{display:flex;gap:8px;justify-content:flex-end;margin-top:16px}
.note{color:var(--muted);font-size:12px;margin-top:22px}

@media (max-width:1100px){
  .kpis{grid-template-columns:repeat(3,minmax(0,1fr))}
  .kpis>div:nth-child(4){border-left:0}
  .kpis>div:nth-child(n+4){border-top:1px solid var(--line)}
}
@media (max-width:980px){
  .grid2,.grid21{grid-template-columns:1fr}
}
@media (max-width:640px){
  .kpis{grid-template-columns:repeat(2,minmax(0,1fr))}
  .kpis>div{border-left:0!important;border-top:1px solid var(--line)}
  .kpis>div:nth-child(-n+2){border-top:0}
  .kpis>div:nth-child(even){border-left:1px solid var(--line)!important}
  .tl-axis,.tl-row{grid-template-columns:92px minmax(0,1fr) 44px;gap:8px}
  .tl-guides{left:100px;right:52px}
  .tl-name{font-size:16px}
  .hb{grid-template-columns:90px minmax(0,1fr) 46px}
  .filters label{flex:1 1 140px}
  .filters select,.filters input,.filters input[type=search]{min-width:0;width:100%}
}
.tl-scale span:first-child{transform:none;text-align:left}
.tl-scale span:last-child{transform:translateX(-100%);text-align:right}
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <div>
      <h1>Pré-indução de hangar</h1>
      <div class="sub" id="subInfo">Planejamento de Suprimentos</div>
    </div>
    <div class="actions">
      <button class="btn" id="themeBtn" type="button">Tema</button>
      <div class="menu">
        <button class="btn" id="menuBtn" type="button" aria-haspopup="true" aria-expanded="false">Histórico</button>
        <div class="menu-list hidden" id="menuList" role="menu">
          <button type="button" role="menuitem" id="exportBtn">Baixar cópia do histórico</button>
          <button type="button" role="menuitem" id="importBtn">Restaurar histórico de um arquivo</button>
          <button type="button" role="menuitem" id="clearBtn">Apagar dados deste navegador</button>
        </div>
      </div>
      <button class="btn primary" id="loadBtn" type="button">Carregar planilhas</button>
      <input type="file" id="fileInput" accept=".xlsx,.xlsm" multiple class="hidden">
      <input type="file" id="importInput" accept=".json" class="hidden">
    </div>
  </header>

  <nav class="tabs" aria-label="Seções">
    <a href="#visao" data-tab="visao">Visão geral</a>
    <a href="#preinducao" data-tab="preinducao">Pré-indução</a>
    <a href="#pedidos" data-tab="pedidos">Pedidos e backorders</a>
    <a href="#estoque" data-tab="estoque">Posição de estoque</a>
  </nav>

  <section class="sources" aria-label="Planilhas carregadas">
    <div class="src" id="srcPre"><i class="dot"></i><div><b>Planilha de pré-indução</b><span id="srcPreTxt">Ainda não carregada</span></div></div>
    <div class="src" id="srcEst"><i class="dot"></i><div><b>Posição de estoque</b><span id="srcEstTxt">Ainda não carregada (opcional)</span></div></div>
  </section>
  <div class="banner">O histórico fica salvo neste navegador. Para não perder nada ao trocar de computador, use <b>Histórico › Baixar cópia do histórico</b> de vez em quando.</div>

  <section id="empty" class="drop" tabindex="0" aria-label="Área para soltar as planilhas">
    <h2>Arraste as planilhas para cá</h2>
    <p>Solte a planilha de pré-indução (aba Caminho Crítico) e, se quiser, a Posição de estoque diária. Pode soltar as duas juntas: o painel reconhece cada uma sozinho.</p>
    <p>Tudo é processado no seu navegador; nenhum dado é enviado para a internet.</p>
    <div class="msg" id="emptyMsg"></div>
  </section>

  <main id="dash" class="hidden">
    <section class="filters" aria-label="Filtros gerais">
      <label>Aeronave / evento<select id="fEvento"></select></label>
      <label>Responsável<select id="fResp"></select></label>
      <label>Atividade<select id="fAtiv"></select></label>
      <label>Buscar PN, descrição, WO ou pedido<input id="fBusca" type="search" placeholder="Ex.: SEALANT, 4500403728"></label>
      <button class="btn" id="resetBtn" type="button">Limpar filtros</button>
    </section>

    <!-- VISÃO GERAL -->
    <section data-view="visao">
      <div class="kpis" id="kVisao"></div>
      <div class="grid21">
        <section class="panel">
          <h3>Linha do tempo das induções</h3>
          <p class="hint">Cada aeronave aparece na distância em dias até a entrada no hangar. Dentro da faixa azul (T-60) a meta é ter 100% do material atendido.</p>
          <div class="tl-legend">
            <span><i style="background:var(--disp)"></i>100% atendido</span>
            <span><i style="background:var(--atende)"></i>Com pendências, fora da meta</span>
            <span><i style="background:var(--nao)"></i>Com pendências dentro de T-60</span>
            <span><i class="z"></i>Zona da meta (T-60)</span>
          </div>
          <div class="tl" id="timeline"></div>
        </section>
        <div style="display:flex;flex-direction:column;gap:18px;min-width:0">
          <section class="panel">
            <h3>Fora da meta</h3>
            <p class="hint">Aeronaves a 60 dias ou menos da indução com linhas que não atendem.</p>
            <div class="alerts" id="alerts"></div>
          </section>
          <section class="panel">
            <h3>Evolução das cargas</h3>
            <div class="trend-legend"><span><i style="background:var(--disp)"></i>% atendido</span><span><i style="background:var(--nao)"></i>Linhas que não atendem</span></div>
            <svg class="trend" id="trend" viewBox="0 0 400 150" preserveAspectRatio="none" role="img" aria-label="Evolução das cargas"></svg>
            <p class="hint" id="trendNote" style="margin:0"></p>
          </section>
        </div>
      </div>
      <div class="grid2">
        <section class="panel">
          <h3>Top 10 itens críticos</h3>
          <p class="hint">PNs com mais linhas que não atendem nas próximas induções.</p>
          <div id="topPn"></div>
        </section>
        <section class="panel">
          <h3>Top 10 aeronaves com pendências</h3>
          <p class="hint">Linhas que não atendem por aeronave. Clique para filtrar.</p>
          <div class="hbars" id="topAcft"></div>
        </section>
      </div>
      <div class="grid2">
        <section class="panel">
          <h3>Quem precisa agir</h3>
          <p class="hint">Linhas que não atendem por área responsável.</p>
          <div class="hbars" id="byResp"></div>
        </section>
        <section class="panel">
          <h3>Situação da previsão</h3>
          <p class="hint">Como cada linha ativa está sendo atendida.</p>
          <div class="hbars" id="byPrev"></div>
        </section>
      </div>
    </section>

    <!-- PRÉ-INDUÇÃO -->
    <section data-view="preinducao" class="hidden">
      <div class="kpis" id="kPre"></div>
      <section class="panel full">
        <div class="panel-head">
          <div><h3>Linhas de demanda</h3><p class="hint" id="preCount"></p></div>
          <div class="filters">
            <label>Status<select id="fStatus">
              <option value="n">Não atende</option>
              <option value="a">Chega a tempo</option>
              <option value="d">Disponível</option>
              <option value="ativas">Todas as ativas</option>
              <option value="gone">Saíram da planilha</option>
            </select></label>
            <label>Marco<select id="fMarco"></select></label>
            <label>Na última carga<select id="fMud">
              <option value="">Todas as linhas</option>
              <option value="mudou">Mudaram de status</option>
              <option value="nova">Entraram agora</option>
            </select></label>
            <button class="btn" type="button" data-csv="tPre">Exportar CSV</button>
          </div>
        </div>
        <div id="tPre"></div>
      </section>
    </section>

    <!-- PEDIDOS -->
    <section data-view="pedidos" class="hidden">
      <div class="kpis" id="kPed"></div>
      <div class="grid2">
        <section class="panel">
          <h3>Chegadas previstas por semana</h3>
          <p class="hint">Linhas com pedido e previsão de chegada nas próximas 8 semanas.</p>
          <div class="cols" id="weeks"></div>
        </section>
        <section class="panel">
          <h3>Top 10 pedidos em backorder</h3>
          <p class="hint">Pedidos atrasados ou sem previsão que travam mais linhas.</p>
          <div id="topPed"></div>
        </section>
      </div>
      <section class="panel full">
        <div class="panel-head">
          <div><h3>Linhas com pedido</h3><p class="hint" id="pedCount"></p></div>
          <div class="filters">
            <label>Situação<select id="fSit">
              <option value="back">Backorder (atrasado ou sem previsão)</option>
              <option value="atrasado">Atrasado</option>
              <option value="sem">Sem previsão</option>
              <option value="prazo">No prazo</option>
              <option value="">Todos</option>
            </select></label>
            <label>Tipo de pedido<select id="fTipo">
              <option value="">Todos</option>
              <option value="PO de compra">PO de compra</option>
              <option value="Requisição (RC)">Requisição (RC)</option>
              <option value="Reparo (RO)">Reparo (RO)</option>
              <option value="Produção (MP)">Produção (MP)</option>
              <option value="Outro">Outro</option>
            </select></label>
            <button class="btn" type="button" data-csv="tPed">Exportar CSV</button>
          </div>
        </div>
        <div id="tPed"></div>
      </section>
      <section class="panel full">
        <h3>Não atendem sem pedido</h3>
        <p class="hint">Linhas que não atendem e ainda não têm PO, requisição, reparo ou ordem de produção.</p>
        <div class="panel-head" style="justify-content:flex-end"><button class="btn small" type="button" data-csv="tSemPed">Exportar CSV</button></div>
        <div id="tSemPed" style="margin-top:8px"></div>
      </section>
    </section>

    <!-- ESTOQUE -->
    <section data-view="estoque" class="hidden">
      <div id="estEmpty" class="drop" tabindex="0">
        <h2>Carregue a Posição de estoque</h2>
        <p>Solte aqui o arquivo Posição de estoque.xlsx (o mesmo que a automação salva na pasta Relatórios automatizados). O painel cruza o estoque de todas as bases com as linhas que não atendem.</p>
      </div>
      <div id="estDash" class="hidden">
        <div class="kpis" id="kEst"></div>
        <div class="grid2">
          <section class="panel">
            <h3>Registros por base</h3>
            <p class="hint">As 12 bases com mais registros na posição de estoque.</p>
            <div class="hbars" id="byStation"></div>
          </section>
          <section class="panel">
            <h3>Consultar PN</h3>
            <div class="filters" style="margin-top:6px"><label>Part number<input id="pnQuery" type="search" placeholder="Digite o PN exato"></label></div>
            <div id="pnResult" style="margin-top:12px"></div>
          </section>
        </div>
        <section class="panel full">
          <div class="panel-head">
            <div><h3>Oportunidades de transferência</h3><p class="hint" id="oppCount">Linhas que não atendem, mas o PN tem estoque em alguma base.</p></div>
            <button class="btn" type="button" data-csv="tOpp">Exportar CSV</button>
          </div>
          <div id="tOpp"></div>
        </section>
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
    <div class="row">
      <button class="btn" id="cfCancel" type="button">Cancelar</button>
      <button class="btn primary" id="cfSave" type="button">Salvar carga</button>
    </div>
  </div>
</div>

<script>
(function(){
"use strict";
var SHEET = "Caminho Crítico", DAY = 86400000, DB = "preinducao-v2";
var $ = function(id){ return document.getElementById(id); };
var S = { rows: [], cargas: [], meta: null, stock: null, events: [], tab: "visao" };
var pending = null, queue = [];

/* ================= utilidades ================= */
function norm(s){ return String(s == null ? "" : s).normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s+/g," ").trim().toLowerCase(); }
function str(v){ return v == null ? "" : String(v).trim(); }
function isSerial(v){ return typeof v === "number" && v > 30000 && v < 80000; }
function toDay(v){ return isSerial(v) ? Math.floor(v) - 25569 : null; }
function today(){ var n = new Date(); return Math.floor(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) / DAY); }
function fd(d, short){
  if (d == null) return "";
  var x = new Date(d * DAY), dd = String(x.getUTCDate()).padStart(2,"0"), mm = String(x.getUTCMonth()+1).padStart(2,"0");
  return short ? dd + "/" + mm : dd + "/" + mm + "/" + x.getUTCFullYear();
}
function pct(a,b){ return b ? a/b*100 : 0; }
function fp(a,b){ return b ? pct(a,b).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1}) + "%" : "–"; }
function fi(n){ return Number(n || 0).toLocaleString("pt-BR"); }
function esc(s){ return String(s == null ? "" : s).replace(/[&<>"]/g, function(c){ return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]; }); }
function pl(n, a, b){ return fi(n) + " " + (n === 1 ? a : b); }
function upn(s){ return str(s).toUpperCase(); }
function statusOf(v){ var s = norm(v); if (s.charAt(0) === "1" || s.indexOf("disponivel") >= 0) return "d"; if (s.charAt(0) === "2" || s === "atende") return "a"; return "n"; }
var ST_CLS = { d: "disp", a: "atende", n: "nao" };
var ST_LBL = { d: "Disponível", a: "Chega a tempo", n: "Não atende" };
var ST_SHORT = { d: "Disp.", a: "Chega", n: "Não at." };
function marcoOrder(m){ var x = /^T(\d+)$/i.exec(m || ""); return x ? -parseInt(x[1],10) : 1; }
function prevTipo(raw){
  if (isSerial(raw)) return "data";
  var s = norm(raw);
  if (!s) return "sem";
  if (s.indexOf("transferir") >= 0) return "transferir";
  if (s.indexOf("sem previs") >= 0) return "sem";
  if (s.indexOf("estoque") === 0) return "estoque";
  if (s.indexOf("mro") >= 0) return "mro";
  if (s.indexOf("paliativo") >= 0) return "paliativo";
  return "outro";
}
var PREV_LBL = { estoque: "Em estoque", transferir: "Estoque a transferir", data: "Com data prevista", sem: "Sem previsão", mro: "MRO", paliativo: "Paliativo", outro: "Outros" };
function tipoPedido(order, ext){
  var o = upn(order), x = str(ext);
  if (!o) return "";
  if (/TRF$/.test(o) || /^45/.test(x)) return "PO de compra";
  if (/^MP/.test(o)) return "Produção (MP)";
  if (/^B/.test(o)) return "Requisição (RC)";
  if (/^R/.test(o)) return "Reparo (RO)";
  return "Outro";
}
function marcoAtual(days){
  if (days == null) return "";
  if (days < 0) return "Em execução";
  if (days <= 7) return "T-7";
  if (days <= 45) return "T-45";
  if (days <= 60) return "T-60";
  if (days <= 120) return "T-120";
  return "T-180";
}
function sleep(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }

/* ================= leitura: pré-indução ================= */
var COLS = {
  prog: ["programacao"], marco: ["demanda"], wo: ["wo"], acft: ["acft"], pn: ["pn"], desc: ["description","descricao"],
  qty: ["qty","quantidade"], nec: ["data necessidade"], prev: ["previsao na cia"], atend: ["atendimento"],
  resp: ["responsavel"], obs: ["observacoes"], order: ["order no. (oh_orderno)"], ext: ["ext. order (oh_ext_orderno)"],
  pacote: ["pacote"], ativ: ["atividade"], acao: ["acao confirmed"]
};
function findCol(head, names, prefix){
  var idx = -1;
  names.some(function(n){ idx = head.indexOf(n); return idx >= 0; });
  if (idx < 0 && prefix) idx = head.findIndex(function(h){ return h.indexOf(prefix) === 0; });
  return idx;
}
function parsePre(ws){
  var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  var h = -1;
  for (var i = 0; i < Math.min(15, aoa.length); i++) { var hs = (aoa[i] || []).map(norm); if (hs.indexOf("atendimento") >= 0 && hs.indexOf("pacote") >= 0) { h = i; break; } }
  if (h < 0) throw new Error("Não encontrei o cabeçalho da aba “" + SHEET + "” (colunas Atendimento e Pacote).");
  var head = aoa[h].map(norm), m = {};
  Object.keys(COLS).forEach(function(k){
    m[k] = findCol(head, COLS[k], k === "order" ? "order no" : k === "ext" ? "ext. order" : null);
  });
  ["pacote","wo","pn","atend"].forEach(function(k){ if (m[k] < 0) throw new Error("Coluna obrigatória não encontrada: " + COLS[k][0]); });
  var out = [], seen = {};
  for (var r = h + 1; r < aoa.length; r++) {
    var a = aoa[r]; if (!a) continue;
    var g = function(k){ return m[k] >= 0 ? a[m[k]] : null; };
    if (g("pn") == null && g("pacote") == null) continue;
    var prev = g("prev"), order = str(g("order"));
    var row = {
      pacote: str(g("pacote")), acft: str(g("acft")), wo: str(g("wo")), pn: str(g("pn")), desc: str(g("desc")),
      qty: typeof g("qty") === "number" ? g("qty") : str(g("qty")), nec: toDay(g("nec")), prog: toDay(g("prog")),
      marco: str(g("marco")) || "Sem marco", prevDay: toDay(prev), prevTxt: isSerial(prev) ? "" : str(prev), prevTipo: prevTipo(prev),
      st: statusOf(g("atend")), resp: str(g("resp")) || "Sem responsável", obs: str(g("obs")).slice(0, 300),
      order: order, ext: str(g("ext")), tipo: tipoPedido(order, g("ext")), ativ: str(g("ativ")).replace(/^[:\s]+/, ""), acao: str(g("acao"))
    };
    var base = row.pacote + "|" + row.wo + "|" + row.pn;
    seen[base] = (seen[base] || 0) + 1;
    row.k = seen[base] > 1 ? base + "#" + seen[base] : base;
    out.push(row);
  }
  if (!out.length) throw new Error("A aba “" + SHEET + "” está vazia.");
  return out;
}

/* ================= leitura: posição de estoque ================= */
function parseStock(ws, fileName){
  var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  var h = -1;
  for (var i = 0; i < Math.min(10, aoa.length); i++) { var hs = (aoa[i] || []).map(norm); if (hs.indexOf("partno") >= 0 && hs.indexOf("station") >= 0) { h = i; break; } }
  if (h < 0) return null;
  var head = aoa[h].map(norm);
  var cPn = head.indexOf("partno"), cSt = head.indexOf("station"), cQty = head.indexOf("qty"), cDesc = head.indexOf("description");
  var map = {}, desc = {}, byStation = {}, regs = 0, qtyTot = 0;
  for (var r = h + 1; r < aoa.length; r++) {
    var a = aoa[r]; if (!a || a[cPn] == null) continue;
    var pn = upn(a[cPn]), st = upn(a[cSt]) || "?", q = typeof a[cQty] === "number" ? a[cQty] : parseFloat(a[cQty]) || 0;
    var e = map[pn] || (map[pn] = {});
    e[st] = (e[st] || 0) + q;
    if (!desc[pn] && cDesc >= 0 && a[cDesc]) desc[pn] = str(a[cDesc]).slice(0, 60);
    byStation[st] = (byStation[st] || 0) + 1;
    regs++; qtyTot += q;
  }
  var compact = {};
  Object.keys(map).forEach(function(pn){
    compact[pn] = Object.keys(map[pn]).map(function(st){ return [st, Math.round(map[pn][st] * 1000) / 1000]; }).sort(function(x,y){ return y[1] - x[1]; });
  });
  return { arquivo: fileName, ts: Date.now(), dia: today(), regs: regs, qty: qtyTot, map: compact, desc: desc, byStation: byStation };
}

/* ================= histórico: mesclagem ================= */
var FIELDS = ["pacote","acft","wo","pn","desc","qty","nec","prog","marco","prevDay","prevTxt","prevTipo","st","resp","obs","order","ext","tipo","ativ","acao"];
function merge(oldRows, newRows, day){
  var byKey = {}, res = [], s = { lidas: newRows.length, novas: 0, mudaram: 0, voltaram: 0, sairam: 0 };
  oldRows.forEach(function(r){ byKey[r.k] = r; });
  var inNew = {};
  newRows.forEach(function(n){
    inNew[n.k] = 1;
    var o = byKey[n.k], r = { k: n.k };
    FIELDS.forEach(function(f){ r[f] = n[f]; });
    if (!o) { r.first = day; r.last = day; r.gone = null; r.chg = day; r.hist = [[day, n.st]]; s.novas++; }
    else {
      r.first = o.first; r.last = day; r.gone = null; r.hist = (o.hist || []).slice(); r.chg = o.chg;
      if (o.gone != null) s.voltaram++;
      var last = r.hist.length ? r.hist[r.hist.length-1][1] : null;
      if (last !== n.st) {
        if (r.hist.length && r.hist[r.hist.length-1][0] === day) r.hist[r.hist.length-1] = [day, n.st]; else r.hist.push([day, n.st]);
        r.chg = day; s.mudaram++;
      }
    }
    res.push(r);
  });
  oldRows.forEach(function(o){
    if (inNew[o.k]) return;
    var r = {}; Object.keys(o).forEach(function(f){ r[f] = o[f]; });
    if (r.gone == null) { r.gone = day; s.sairam++; }
    res.push(r);
  });
  return { rows: res, stats: s };
}
function isBack(r){ return !!r.order && (r.prevTipo === "sem" || (r.prevDay != null && r.nec != null && r.prevDay > r.nec)); }
function summarize(rows, day, file, stats){
  var s = { ts: Date.now(), dia: day, arquivo: file, ativas: 0, d: 0, a: 0, n: 0, back: 0 };
  rows.forEach(function(r){ if (r.gone != null) return; s.ativas++; s[r.st]++; if (isBack(r)) s.back++; });
  ["lidas","novas","mudaram","voltaram","sairam"].forEach(function(k){ s[k] = stats[k]; });
  return s;
}

/* ================= armazenamento local ================= */
var PACK = ["k"].concat(FIELDS).concat(["first","last","gone","chg","hist"]);
function pack(r){ return PACK.map(function(f){ return r[f] === undefined ? null : r[f]; }); }
function unpack(a){ var r = {}; PACK.forEach(function(f, i){ r[f] = a[i]; }); if (!r.hist) r.hist = []; return r; }
function openDb(){ return new Promise(function(res, rej){ try { var rq = indexedDB.open(DB, 1); rq.onupgradeneeded = function(){ rq.result.createObjectStore("s"); }; rq.onsuccess = function(){ res(rq.result); }; rq.onerror = function(){ rej(rq.error); }; } catch(e){ rej(e); } }); }
function dbGet(key){ return openDb().then(function(db){ return new Promise(function(res){ try { var rq = db.transaction("s").objectStore("s").get(key); rq.onsuccess = function(){ res(rq.result || null); }; rq.onerror = function(){ res(null); }; } catch(e){ res(null); } }); }).catch(function(){ return null; }); }
function dbPut(key, val){ return openDb().then(function(db){ return new Promise(function(res, rej){ var tx = db.transaction("s","readwrite"); tx.objectStore("s").put(val, key); tx.oncomplete = res; tx.onerror = function(){ rej(tx.error); }; }); }); }
function dbDel(key){ return openDb().then(function(db){ return new Promise(function(res){ var tx = db.transaction("s","readwrite"); tx.objectStore("s").delete(key); tx.oncomplete = res; tx.onerror = res; }); }).catch(function(){}); }
function savePre(){ return dbPut("pre", { meta: S.meta, rows: S.rows.map(pack), cargas: S.cargas }); }
function saveStock(){ return dbPut("estoque", S.stock); }

/* ================= carregar arquivos ================= */
function overlay(msg){ if (msg) { $("overlayMsg").textContent = msg; $("overlay").classList.add("on"); } else $("overlay").classList.remove("on"); }
function fail(msg){
  overlay(null);
  if (S.rows.length || S.stock) window.alert(msg);
  else { $("emptyMsg").className = "msg err"; $("emptyMsg").textContent = msg; }
}
function handleFiles(list){
  if (typeof XLSX === "undefined") { fail("A biblioteca que lê o Excel não carregou. Confira a conexão com a internet e recarregue a página."); return; }
  queue = Array.prototype.slice.call(list || []).filter(function(f){ return /\.xls[xm]?$/i.test(f.name); });
  if (!queue.length) { fail("Escolha arquivos do Excel (.xlsx)."); return; }
  nextFile();
}
function nextFile(){
  var file = queue.shift();
  if (!file) { overlay(null); return; }
  overlay("Lendo “" + file.name + "”…");
  var reader = new FileReader();
  reader.onerror = function(){ fail("Não foi possível ler “" + file.name + "”."); };
  reader.onload = function(e){
    setTimeout(function(){
      try {
        var buf = new Uint8Array(e.target.result);
        var wb = XLSX.read(buf, { type: "array", sheets: [SHEET], dense: true, cellDates: false, cellFormula: false, cellHTML: false, cellText: false, cellStyles: false });
        var preName = wb.SheetNames.filter(function(n){ return norm(n) === norm(SHEET); })[0];
        if (preName) {
          if (!wb.Sheets[preName]) wb = XLSX.read(buf, { type: "array", sheets: [preName], dense: true, cellDates: false, cellFormula: false, cellHTML: false, cellText: false });
          var newRows = parsePre(wb.Sheets[preName]);
          overlay(null);
          var day = today();
          pending = { file: file.name, day: day, newRows: newRows, merged: merge(S.rows, newRows, day) };
          openConfirm();
          return;
        }
        var first = wb.SheetNames[0];
        var wb2 = XLSX.read(buf, { type: "array", sheets: [first], dense: true, cellDates: false, cellFormula: false, cellHTML: false, cellText: false });
        var stock = parseStock(wb2.Sheets[first], file.name);
        if (!stock) throw new Error("“" + file.name + "” não parece ser a planilha de pré-indução (aba Caminho Crítico) nem a Posição de estoque (colunas partno e station).");
        S.stock = stock;
        overlay("Salvando posição de estoque…");
        saveStock().catch(function(){}).then(function(){ showDash(); nextFile(); });
      } catch(err) { fail(err.message || String(err)); }
    }, 40);
  };
  reader.readAsArrayBuffer(file);
}
function openConfirm(){
  var s = pending.merged.stats;
  $("cfFile").textContent = pending.file + ", carga de " + fd(pending.day);
  $("cfStats").innerHTML = [["Linhas lidas", s.lidas], ["Linhas novas", s.novas], ["Mudaram de status", s.mudaram], ["Voltaram para a planilha", s.voltaram], ["Saíram da planilha", s.sairam]]
    .map(function(x){ return "<dt>" + x[0] + "</dt><dd>" + fi(x[1]) + "</dd>"; }).join("");
  $("confirm").classList.add("on"); $("cfSave").focus();
}
function closeConfirm(){ $("confirm").classList.remove("on"); }
function saveCarga(){
  closeConfirm();
  var p = pending; pending = null;
  var carga = summarize(p.merged.rows, p.day, p.file, p.merged.stats);
  S.rows = p.merged.rows;
  S.cargas = S.cargas.filter(function(c){ return c.dia !== p.day; }).concat([carga]);
  S.meta = { arquivo: p.file, ts: carga.ts, dia: p.day };
  overlay("Salvando histórico…");
  savePre().then(function(){ showDash(); nextFile(); }).catch(function(e){ fail("A carga não foi salva no navegador (" + ((e && e.message) || "erro") + ")."); });
}

/* ================= filtros ================= */
function F(){ return { ev: $("fEvento").value, resp: $("fResp").value, ativ: $("fAtiv").value, q: norm($("fBusca").value) }; }
function match(r, f, ignoreEv){
  if (!ignoreEv && f.ev && r.pacote !== f.ev) return false;
  if (f.resp && r.resp !== f.resp) return false;
  if (f.ativ && r.ativ !== f.ativ) return false;
  if (f.q && norm(r.pn + " " + r.desc + " " + r.wo + " " + r.order + " " + r.ext).indexOf(f.q) < 0) return false;
  return true;
}
function uniq(rows, k){ var s = {}; rows.forEach(function(r){ if (r[k]) s[r[k]] = 1; }); return Object.keys(s); }
function mostCommon(arr){ var c = {}, best = "", n = 0; arr.forEach(function(x){ c[x] = (c[x]||0)+1; if (c[x] > n){ n = c[x]; best = x; } }); return best; }
function fillSelect(id, opts){
  var el = $(id), keep = el.value;
  el.innerHTML = opts.map(function(o){ return '<option value="' + esc(o[0]) + '">' + esc(o[1]) + "</option>"; }).join("");
  if (opts.some(function(o){ return o[0] === keep; })) el.value = keep;
}
function lastDay(){ return S.cargas.length ? S.cargas[S.cargas.length-1].dia : null; }
function stockOf(pn){ return S.stock ? S.stock.map[upn(pn)] || null : null; }
function stockTotal(pn){ var s = stockOf(pn); return s ? s.reduce(function(t,x){ return t + x[1]; }, 0) : 0; }
function stockTxt(pn){ var s = stockOf(pn); return s ? s.slice(0, 6).map(function(x){ return x[0] + " " + fi(x[1]); }).join(", ") + (s.length > 6 ? " e mais " + (s.length - 6) : "") : ""; }

/* ================= montagem ================= */
function showDash(){
  var hasPre = S.rows.length > 0;
  var active = S.rows.filter(function(r){ return r.gone == null; });
  var byPac = {};
  active.forEach(function(r){ (byPac[r.pacote] = byPac[r.pacote] || []).push(r); });
  S.events = Object.keys(byPac).map(function(p){
    var rs = byPac[p];
    var prog = rs.reduce(function(m,r){ return r.prog != null && (m == null || r.prog < m) ? r.prog : m; }, null);
    return { pacote: p, prog: prog, acft: mostCommon(rs.map(function(r){ return r.acft; })), ativ: mostCommon(rs.map(function(r){ return r.ativ; })) };
  }).sort(function(a,b){ return (a.prog||0) - (b.prog||0) || a.pacote.localeCompare(b.pacote); });

  fillSelect("fEvento", [["","Todas"]].concat(S.events.map(function(e){ return [e.pacote, e.acft + " (" + fd(e.prog, true) + ") " + e.pacote]; })));
  fillSelect("fResp", [["","Todos"]].concat(uniq(S.rows,"resp").sort().map(function(x){ return [x,x]; })));
  fillSelect("fAtiv", [["","Todas"]].concat(uniq(S.rows,"ativ").sort().map(function(x){ return [x,x]; })));
  fillSelect("fMarco", [["","Todos"]].concat(uniq(S.rows,"marco").sort(function(a,b){ return marcoOrder(a) - marcoOrder(b); }).map(function(x){ return [x,x]; })));

  $("srcPre").className = "src" + (hasPre ? (S.meta && S.meta.dia === today() ? " ok" : " old") : "");
  $("srcPreTxt").textContent = hasPre ? S.meta.arquivo + ", carregada em " + new Date(S.meta.ts).toLocaleString("pt-BR") : "Ainda não carregada";
  $("srcEst").className = "src" + (S.stock ? (S.stock.dia === today() ? " ok" : " old") : "");
  $("srcEstTxt").textContent = S.stock ? S.stock.arquivo + ", carregada em " + new Date(S.stock.ts).toLocaleString("pt-BR") : "Ainda não carregada (opcional)";

  var any = hasPre || !!S.stock;
  $("empty").classList.toggle("hidden", any);
  $("dash").classList.toggle("hidden", !any);
  if (any) $("emptyMsg").textContent = "";
  var gone = S.rows.length - active.length;
  $("note").textContent = hasPre ? "Histórico neste navegador: " + pl(S.cargas.length, "carga", "cargas") + ", " + pl(active.length, "linha ativa", "linhas ativas") + " e " + pl(gone, "linha que já saiu da planilha", "linhas que já saíram da planilha") + "." : "";
  render();
}

function render(){
  var hasPre = S.rows.length > 0;
  ["visao","preinducao","pedidos"].forEach(function(v){
    document.querySelector('[data-view="' + v + '"]').classList.toggle("hidden", S.tab !== v);
  });
  document.querySelector('[data-view="estoque"]').classList.toggle("hidden", S.tab !== "estoque");
  document.querySelectorAll(".tabs a").forEach(function(a){ if (a.getAttribute("data-tab") === S.tab) a.setAttribute("aria-current","page"); else a.removeAttribute("aria-current"); });
  if (!hasPre && S.tab !== "estoque") {
    var v = document.querySelector('[data-view="' + S.tab + '"]');
    if (v && !v.querySelector(".needpre")) v.insertAdjacentHTML("afterbegin", '<div class="drop needpre" tabindex="0"><h2>Carregue a planilha de pré-indução</h2><p>Esta seção usa a aba Caminho Crítico. Clique ou solte o arquivo aqui.</p></div>');
  }
  document.querySelectorAll(".needpre").forEach(function(n){ if (hasPre) n.remove(); });
  var f = F();
  if (S.tab === "visao" && hasPre) renderVisao(f);
  if (S.tab === "preinducao" && hasPre) renderPre(f);
  if (S.tab === "pedidos" && hasPre) renderPed(f);
  if (S.tab === "estoque") renderEst(f);
}
function kpis(id, items){
  $(id).innerHTML = items.map(function(x){ return '<div class="' + (x[2] || "") + '"><b>' + x[0] + "</b><span>" + x[1] + "</span></div>"; }).join("");
}
function bars(id, arr, color, clickKey){
  var max = arr.reduce(function(m,x){ return Math.max(m, x[1]); }, 0);
  $(id).innerHTML = arr.length ? arr.map(function(x){
    var inner = "<span>" + esc(x[0]) + '</span><span class="track"><span style="width:' + pct(x[1], max) + "%;background:" + (x[2] || color) + '"></span></span><b>' + fi(x[1]) + "</b>";
    return clickKey ? '<button type="button" class="hb" data-' + clickKey + '="' + esc(x[3] || x[0]) + '">' + inner + "</button>" : '<div class="hb">' + inner + "</div>";
  }).join("") : '<p class="empty">Nada com os filtros atuais.</p>';
}
function countBy(rows, key){ var c = {}; rows.forEach(function(r){ c[r[key]] = (c[r[key]]||0)+1; }); return Object.keys(c).map(function(k){ return [k, c[k]]; }); }

/* ---------- visão geral ---------- */
function renderVisao(f){
  var base = S.rows.filter(function(r){ return r.gone == null && match(r, f, false); });
  var c = { d: 0, a: 0, n: 0 }, back = 0, semPed = 0, comEst = 0;
  base.forEach(function(r){ c[r.st]++; if (isBack(r)) back++; if (r.st === "n" && !r.order) semPed++; if (r.st === "n" && stockTotal(r.pn) > 0) comEst++; });
  var t = base.length;
  kpis("kVisao", [
    [fi(t), "linhas ativas"],
    [fp(c.d, t), "material disponível", "good"],
    [fp(c.d + c.a, t), "atendido (disponível ou chega a tempo)"],
    [fi(c.n), "linhas que não atendem", "bad"],
    [fi(back), "em backorder (pedido atrasado ou sem previsão)", "warn"],
    S.stock ? [fi(comEst), "não atendem, mas têm estoque em alguma base"] : [fi(semPed), "não atendem e ainda não têm pedido"]
  ]);
  renderTimeline(f);
  renderTrend();

  var nao = base.filter(function(r){ return r.st === "n"; });
  var g = {};
  nao.forEach(function(r){
    var e = g[r.pn] || (g[r.pn] = { pn: r.pn, desc: r.desc, linhas: 0, qty: 0, acft: {}, nec: null, ped: 0 });
    e.linhas++; e.qty += typeof r.qty === "number" ? r.qty : 0; e.acft[r.acft] = 1;
    if (r.nec != null && (e.nec == null || r.nec < e.nec)) e.nec = r.nec;
    if (r.order) e.ped++;
  });
  var top = Object.keys(g).map(function(k){ var e = g[k]; e.nAcft = Object.keys(e.acft).length; e.est = stockTotal(e.pn); return e; })
    .sort(function(a,b){ return b.linhas - a.linhas || (a.nec||0) - (b.nec||0); }).slice(0, 10);
  $("topPn").innerHTML = top.length ? '<div class="scroll"><table><thead><tr><th>PN</th><th>Descrição</th><th class="num">Linhas</th><th class="num">Acft</th><th>1ª necessidade</th>' + (S.stock ? '<th class="num">Estoque</th>' : "") + "</tr></thead><tbody>" +
    top.map(function(e){ return '<tr class="click" data-pn="' + esc(e.pn) + '"><td>' + esc(e.pn) + '</td><td class="desc">' + esc(e.desc) + '</td><td class="num">' + fi(e.linhas) + '</td><td class="num">' + e.nAcft + "</td><td>" + fd(e.nec) + "</td>" + (S.stock ? '<td class="num">' + fi(e.est) + "</td>" : "") + "</tr>"; }).join("") +
    "</tbody></table></div>" : '<p class="empty">Nenhuma linha que não atende com os filtros atuais.</p>';

  var byAcft = {};
  S.rows.forEach(function(r){ if (r.gone != null || r.st !== "n" || !match(r, f, true)) return; var k = r.pacote; byAcft[k] = (byAcft[k] || 0) + 1; });
  var evMap = {}; S.events.forEach(function(e){ evMap[e.pacote] = e; });
  bars("topAcft", Object.keys(byAcft).map(function(k){ var e = evMap[k] || {}; return [(e.acft || "?") + " (" + fd(e.prog, true) + ")", byAcft[k], null, k]; }).sort(function(a,b){ return b[1]-a[1]; }).slice(0, 10), "var(--nao)", "ev");
  bars("byResp", countBy(nao, "resp").sort(function(a,b){ return b[1]-a[1]; }), "var(--nao)");
  var pc = {}; base.forEach(function(r){ pc[r.prevTipo] = (pc[r.prevTipo] || 0) + 1; });
  var colors = { estoque: "var(--disp)", transferir: "var(--navy)", data: "var(--atende)", sem: "var(--nao)", mro: "var(--muted)", paliativo: "var(--muted)", outro: "var(--muted)" };
  bars("byPrev", ["estoque","transferir","data","sem","mro","paliativo","outro"].filter(function(k){ return pc[k]; }).map(function(k){ return [PREV_LBL[k], pc[k], colors[k]]; }), "var(--navy)");
}

function evAgg(f){
  var agg = {};
  S.rows.forEach(function(r){ if (r.gone != null || !match(r, f, true)) return; var a = agg[r.pacote] || (agg[r.pacote] = { d:0, a:0, n:0, t:0 }); a[r.st]++; a.t++; });
  return agg;
}
function renderTimeline(f){
  var agg = evAgg(f), t0 = today(), MAX = 180;
  var marks = [[0,"Hoje"],[7,""],[45,"T-45"],[60,"T-60"],[120,"T-120"],[180,"T-180"]];
  var x = function(d){ return Math.max(0, Math.min(MAX, d)) / MAX * 100; };
  var evs = S.events.filter(function(e){ return agg[e.pacote]; });
  var alerts = [];
  var rows = evs.map(function(e){
    var a = agg[e.pacote], days = e.prog != null ? e.prog - t0 : null, at = Math.round(pct(a.d + a.a, a.t));
    var cls = a.n === 0 ? "ok" : (days != null && days <= 60 ? "bad" : "mid");
    if (a.n > 0 && days != null && days <= 60) alerts.push({ e: e, days: days, n: a.n, at: at });
    var sel = f.ev === e.pacote ? " sel" : "";
    var when = days == null ? "sem data" : days < 0 ? "em execução" : days === 0 ? "hoje" : "em " + pl(days, "dia", "dias");
    return '<button type="button" class="tl-row' + sel + '" data-ev="' + esc(e.pacote) + '" title="' + esc(e.acft + ", " + e.ativ + ", " + fd(e.prog) + ", " + a.t + " linhas, " + a.n + " não atendem") + '">' +
      '<span class="tl-name">' + esc(e.acft) + "<small>" + fd(e.prog, true) + ", " + when + "</small></span>" +
      '<span class="tl-track"><span class="tl-pill ' + cls + (days != null && days < 0 ? " past" : "") + '" style="left:' + x(days == null ? MAX : days) + '%">' + at + "%</span></span>" +
      '<span class="tl-pct" style="color:' + (a.n ? "var(--nao)" : "var(--disp)") + '">' + (a.n ? fi(a.n) : "✓") + "</span></button>";
  }).join("");
  var scale = marks.filter(function(m){ return m[1]; }).map(function(m){ return '<span style="left:' + x(m[0]) + '%"><b>' + m[1] + "</b>" + (m[0] ? fd(t0 + m[0], true) : fd(t0, true)) + "</span>"; }).join("");
  var guides = '<span class="zone" style="width:' + x(60) + '%"></span>' + marks.slice(1).map(function(m){ return '<i style="left:' + x(m[0]) + '%"></i>'; }).join("");
  $("timeline").innerHTML = evs.length ? '<div class="tl-axis"><span>Aeronave</span><div class="tl-scale">' + scale + '</div><span style="text-align:right">Pend.</span></div>' +
    '<div class="tl-body"><div class="tl-guides">' + guides + "</div>" + rows + "</div>" : '<p class="empty">Nenhuma aeronave com os filtros atuais.</p>';
  alerts.sort(function(a,b){ return a.days - b.days; });
  $("alerts").innerHTML = alerts.length ? alerts.map(function(x){
    return '<button type="button" class="alert" data-ev="' + esc(x.e.pacote) + '"><b class="reg">' + esc(x.e.acft) + '</b><span><span class="marco">' + marcoAtual(x.days) + "</span>" + fd(x.e.prog) + '<br><span class="why">' + x.at + "% atendido, " + esc(x.e.ativ) + '</span></span><span class="n">' + fi(x.n) + "</span></button>";
  }).join("") : '<p class="empty">Nenhuma aeronave fora da meta. Todas as induções dos próximos 60 dias estão 100% atendidas.</p>';
}
function renderTrend(){
  var cs = S.cargas.slice(-30), svg = $("trend");
  if (cs.length < 2) {
    svg.innerHTML = '<text x="200" y="78" text-anchor="middle">O gráfico aparece a partir da segunda carga.</text>';
    $("trendNote").textContent = cs.length ? "1 carga salva, em " + fd(cs[0].dia) + "." : "";
    return;
  }
  var W = 400, H = 150, L = 6, R = 6, T = 10, B = 18;
  var xx = function(i){ return L + i * (W - L - R) / (cs.length - 1); };
  var pa = cs.map(function(c){ return pct(c.d + c.a, c.ativas); });
  var mn = Math.min.apply(null, pa), mx = Math.max.apply(null, pa); if (mx - mn < 2) { mn -= 1; mx += 1; }
  var nn = cs.map(function(c){ return c.n; }), nmax = Math.max.apply(null, nn) || 1;
  var y1 = function(v){ return T + (1 - (v - mn) / (mx - mn)) * (H - T - B); };
  var y2 = function(v){ return T + (1 - v / nmax) * (H - T - B); };
  var path = function(vals, y){ return vals.map(function(v,i){ return (i ? "L" : "M") + xx(i).toFixed(1) + " " + y(v).toFixed(1); }).join(" "); };
  svg.innerHTML = '<line x1="0" x2="400" y1="' + (H - B) + '" y2="' + (H - B) + '"/><path class="l1" d="' + path(pa, y1) + '"/><path class="l2" d="' + path(nn, y2) + '"/>' +
    '<text x="' + L + '" y="' + (H - 4) + '">' + fd(cs[0].dia, true) + '</text><text x="' + (W - R) + '" y="' + (H - 4) + '" text-anchor="end">' + fd(cs[cs.length-1].dia, true) + "</text>";
  var a = cs[cs.length-2], b = cs[cs.length-1], d = b.n - a.n;
  $("trendNote").textContent = "Desde a carga anterior (" + fd(a.dia, true) + "): " + (d === 0 ? "mesma quantidade de linhas que não atendem" : (d > 0 ? "+" : "") + fi(d) + " linhas que não atendem") + ", " + pl(b.novas || 0, "linha nova", "linhas novas") + ", " + pl(b.mudaram || 0, "mudança de status", "mudanças de status") + ".";
}

/* ---------- tabelas ---------- */
var TB = {};
function table(id, cols, rows, opts){
  opts = opts || {};
  var t = TB[id] || (TB[id] = { key: opts.key || cols[0].k, dir: opts.dir || 1, page: 0 });
  t.cols = cols;
  var col = cols.filter(function(c){ return c.k === t.key; })[0] || cols[0];
  var val = col.val || function(r){ return r[col.k]; };
  rows = rows.slice().sort(function(a,b){
    var x = val(a), y = val(b);
    if (x == null || x === "") return 1; if (y == null || y === "") return -1;
    if (typeof x === "number" && typeof y === "number") return (x - y) * t.dir;
    return String(x).localeCompare(String(y), "pt-BR", { numeric: true }) * t.dir;
  });
  t.rows = rows;
  var size = opts.page || 100, pages = Math.max(1, Math.ceil(rows.length / size));
  if (t.page >= pages) t.page = pages - 1;
  var head = cols.map(function(c){ return '<th scope="col" tabindex="0" data-k="' + c.k + '"' + (c.num ? ' class="num"' : "") + ' aria-sort="' + (c.k === t.key ? (t.dir > 0 ? "ascending" : "descending") : "none") + '">' + c.label + "</th>"; }).join("");
  var body = rows.slice(t.page * size, t.page * size + size).map(function(r){
    return "<tr>" + cols.map(function(c){ return "<td" + (c.cls ? ' class="' + (typeof c.cls === "function" ? c.cls(r) : c.cls) + '"' : c.num ? ' class="num"' : "") + ">" + (c.html ? c.html(r) : esc(c.txt ? c.txt(r) : r[c.k])) + "</td>"; }).join("") + "</tr>";
  }).join("") || '<tr><td colspan="' + cols.length + '">Nenhuma linha com os filtros atuais.</td></tr>';
  $(id).innerHTML = '<div class="scroll"><table class="wide"><thead><tr>' + head + "</tr></thead><tbody>" + body + "</tbody></table></div>" +
    (pages > 1 ? '<div class="pager"><button class="btn small" type="button" data-pg="-1"' + (t.page === 0 ? " disabled" : "") + '>Anterior</button><span>Página ' + (t.page + 1) + " de " + pages + '</span><button class="btn small" type="button" data-pg="1"' + (t.page >= pages - 1 ? " disabled" : "") + ">Próxima</button></div>" : "");
}
function wireTable(id){
  $(id).addEventListener("click", function(e){
    var th = e.target.closest("th[data-k]"), pg = e.target.closest("[data-pg]"), t = TB[id];
    if (th && t) { var k = th.getAttribute("data-k"); if (t.key === k) t.dir *= -1; else { t.key = k; t.dir = 1; } t.page = 0; render(); }
    if (pg && t) { t.page += parseInt(pg.getAttribute("data-pg"), 10); render(); }
  });
  $(id).addEventListener("keydown", function(e){ if (e.key === "Enter" && e.target.matches("th[data-k]")) e.target.click(); });
}
function exportCsv(id){
  var t = TB[id]; if (!t) return;
  var cell = function(v){ v = v == null ? "" : String(v); return /[;"\n]/.test(v) ? '"' + v.replace(/"/g,'""') + '"' : v; };
  var lines = [t.cols.map(function(c){ return c.label; }).join(";")];
  t.rows.forEach(function(r){ lines.push(t.cols.map(function(c){ return cell(c.csv ? c.csv(r) : c.txt ? c.txt(r) : r[c.k]); }).join(";")); });
  download("\ufeff" + lines.join("\r\n"), "preinducao-" + id + "-" + new Date().toISOString().slice(0,10) + ".csv", "text/csv;charset=utf-8");
}
function download(text, name, type){
  var a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([text], { type: type })); a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(function(){ URL.revokeObjectURL(a.href); }, 1500);
}
function stTag(r){ return r.gone != null ? '<span class="tag gone">Saiu em ' + fd(r.gone, true) + "</span>" : '<span class="tag ' + ST_CLS[r.st] + '">' + ST_LBL[r.st] + "</span>"; }
function prevText(r){ return r.prevDay != null ? fd(r.prevDay) : r.prevTxt; }
function histText(r){ var h = r.hist.map(function(x){ return fd(x[0], true) + " " + ST_SHORT[x[1]]; }); if (r.gone != null) h.push(fd(r.gone, true) + " saiu"); return h.join(" → "); }
function qtyTxt(r){ return typeof r.qty === "number" ? r.qty.toLocaleString("pt-BR") : r.qty; }
function atraso(r){ return r.prevDay != null && r.nec != null ? r.prevDay - r.nec : null; }

/* ---------- pré-indução ---------- */
function renderPre(f){
  var base = S.rows.filter(function(r){ return r.gone == null && match(r, f, false); });
  var c = { d: 0, a: 0, n: 0 }; base.forEach(function(r){ c[r.st]++; });
  var d = lastDay(), mud = 0, nov = 0;
  base.forEach(function(r){ if (r.first === d) nov++; else if (r.chg === d && r.hist.length > 1) mud++; });
  kpis("kPre", [
    [fi(base.length), "linhas ativas"], [fi(c.d), "disponíveis", "good"], [fi(c.a), "chegam a tempo", "warn"],
    [fi(c.n), "não atendem", "bad"], [fi(nov), "entraram na última carga"], [fi(mud), "mudaram de status na última carga"]
  ]);
  var st = $("fStatus").value, marco = $("fMarco").value, md = $("fMud").value;
  var rows = S.rows.filter(function(r){
    if (!match(r, f, false)) return false;
    if (st === "gone") { if (r.gone == null) return false; }
    else { if (r.gone != null) return false; if (st !== "ativas" && r.st !== st) return false; }
    if (marco && r.marco !== marco) return false;
    if (md === "mudou" && !(r.chg === d && r.first !== d && r.hist.length > 1)) return false;
    if (md === "nova" && r.first !== d) return false;
    return true;
  });
  $("preCount").textContent = pl(rows.length, "linha", "linhas");
  table("tPre", [
    { k: "acft", label: "Acft" }, { k: "pacote", label: "Pacote" },
    { k: "nec", label: "Necessidade", txt: function(r){ return fd(r.nec); } },
    { k: "pn", label: "PN" }, { k: "desc", label: "Descrição", cls: "desc" },
    { k: "qty", label: "Qtd", num: true, txt: qtyTxt },
    { k: "st", label: "Status", html: stTag, csv: function(r){ return r.gone != null ? "Saiu da planilha" : ST_LBL[r.st]; } },
    { k: "resp", label: "Responsável" },
    { k: "prevDay", label: "Previsão", txt: prevText, val: function(r){ return r.prevDay != null ? r.prevDay : r.prevTxt; } },
    { k: "order", label: "Pedido" }, { k: "marco", label: "Marco" },
    { k: "first", label: "Desde", txt: function(r){ return fd(r.first, true); } },
    { k: "hist", label: "Histórico de status", cls: "hist", txt: histText, val: function(r){ return r.hist.length; } },
    { k: "obs", label: "Observações", cls: "obs" }
  ], rows, { key: "nec" });
}

/* ---------- pedidos ---------- */
function renderPed(f){
  var t0 = today();
  var base = S.rows.filter(function(r){ return r.gone == null && match(r, f, false); });
  var ped = base.filter(function(r){ return r.order; });
  var tipos = { "PO de compra": 0, "Requisição (RC)": 0, "Reparo (RO)": 0, "Produção (MP)": 0, "Outro": 0 }, atr = 0, sem = 0;
  ped.forEach(function(r){ tipos[r.tipo || "Outro"]++; var x = atraso(r); if (x != null && x > 0) atr++; if (r.prevTipo === "sem") sem++; });
  kpis("kPed", [
    [fi(ped.length), "linhas com pedido"], [fi(tipos["PO de compra"]), "com PO de compra"], [fi(tipos["Requisição (RC)"]), "com requisição (RC), aguardando PO"],
    [fi(tipos["Reparo (RO)"] + tipos["Produção (MP)"]), "em reparo (RO) ou produção (MP)"], [fi(atr), "com previsão depois da necessidade", "bad"], [fi(sem), "com pedido, sem previsão", "warn"]
  ]);
  var weeks = [];
  for (var w = 0; w < 8; w++) weeks.push({ ini: t0 + w * 7, n: 0 });
  ped.forEach(function(r){ if (r.prevDay == null || r.prevDay < t0) return; var i = Math.floor((r.prevDay - t0) / 7); if (i < 8) weeks[i].n++; });
  var wmax = weeks.reduce(function(m,x){ return Math.max(m, x.n); }, 0) || 1;
  $("weeks").innerHTML = weeks.map(function(x){ return '<div title="' + fd(x.ini) + " a " + fd(x.ini + 6) + '"><b>' + (x.n || "") + '</b><span class="bar" style="height:' + Math.max(2, pct(x.n, wmax) * 0.82) + '%"></span><em>' + fd(x.ini, true) + "</em></div>"; }).join("");

  var g = {};
  ped.filter(isBack).forEach(function(r){
    var e = g[r.order] || (g[r.order] = { order: r.order, ext: r.ext, tipo: r.tipo, linhas: 0, acft: {}, maior: null, sem: r.prevTipo === "sem" });
    e.linhas++; e.acft[r.acft] = 1; var x = atraso(r); if (x != null && (e.maior == null || x > e.maior)) e.maior = x;
  });
  var top = Object.keys(g).map(function(k){ var e = g[k]; e.nAcft = Object.keys(e.acft).length; return e; }).sort(function(a,b){ return b.linhas - a.linhas || (b.maior||0) - (a.maior||0); }).slice(0, 10);
  $("topPed").innerHTML = top.length ? '<div class="scroll"><table><thead><tr><th>Pedido</th><th>Ext. order</th><th>Tipo</th><th class="num">Linhas</th><th class="num">Acft</th><th class="num">Atraso</th></tr></thead><tbody>' +
    top.map(function(e){ return '<tr class="click" data-q="' + esc(e.order) + '"><td>' + esc(e.order) + "</td><td>" + esc(e.ext) + "</td><td>" + esc(e.tipo) + '</td><td class="num">' + fi(e.linhas) + '</td><td class="num">' + e.nAcft + '</td><td class="num">' + (e.maior != null ? pl(e.maior, "dia", "dias") : "sem previsão") + "</td></tr>"; }).join("") +
    "</tbody></table></div>" : '<p class="empty">Nenhum pedido em backorder com os filtros atuais.</p>';

  var sit = $("fSit").value, tipo = $("fTipo").value;
  var rows = ped.filter(function(r){
    if (tipo && (r.tipo || "Outro") !== tipo) return false;
    var x = atraso(r);
    if (sit === "back") return isBack(r);
    if (sit === "atrasado") return x != null && x > 0;
    if (sit === "sem") return r.prevTipo === "sem";
    if (sit === "prazo") return !isBack(r);
    return true;
  });
  $("pedCount").textContent = pl(rows.length, "linha", "linhas");
  table("tPed", [
    { k: "order", label: "Pedido" }, { k: "ext", label: "Ext. order" }, { k: "tipo", label: "Tipo" },
    { k: "acft", label: "Acft" }, { k: "pn", label: "PN" }, { k: "desc", label: "Descrição", cls: "desc" },
    { k: "qty", label: "Qtd", num: true, txt: qtyTxt },
    { k: "nec", label: "Necessidade", txt: function(r){ return fd(r.nec); } },
    { k: "prevDay", label: "Previsão", txt: prevText, val: function(r){ return r.prevDay != null ? r.prevDay : r.prevTxt; } },
    { k: "atraso", label: "Atraso", num: true, val: atraso, txt: function(r){ var x = atraso(r); return x == null ? "" : x > 0 ? pl(x, "dia", "dias") : "no prazo"; }, cls: function(r){ var x = atraso(r); return "num" + (x != null && x > 0 ? " late" : ""); } },
    { k: "st", label: "Status", html: stTag, csv: function(r){ return ST_LBL[r.st]; } },
    { k: "resp", label: "Responsável" }, { k: "obs", label: "Observações", cls: "obs" }
  ], rows, { key: "atraso", dir: -1 });

  var semPed = base.filter(function(r){ return r.st === "n" && !r.order; });
  table("tSemPed", [
    { k: "acft", label: "Acft" }, { k: "nec", label: "Necessidade", txt: function(r){ return fd(r.nec); } },
    { k: "pn", label: "PN" }, { k: "desc", label: "Descrição", cls: "desc" }, { k: "qty", label: "Qtd", num: true, txt: qtyTxt },
    { k: "prevDay", label: "Previsão", txt: prevText, val: function(r){ return r.prevDay != null ? r.prevDay : r.prevTxt; } },
    { k: "resp", label: "Responsável" }, { k: "obs", label: "Observações", cls: "obs" }
  ], semPed, { key: "nec", page: 50 });
}

/* ---------- estoque ---------- */
function renderEst(f){
  $("estEmpty").classList.toggle("hidden", !!S.stock);
  $("estDash").classList.toggle("hidden", !S.stock);
  if (!S.stock) return;
  var s = S.stock, pns = Object.keys(s.map).length, bases = Object.keys(s.byStation).length;
  var nao = S.rows.filter(function(r){ return r.gone == null && r.st === "n" && match(r, f, false); });
  var opp = nao.filter(function(r){ return stockTotal(r.pn) > 0; });
  var transf = S.rows.filter(function(r){ return r.gone == null && r.prevTipo === "transferir" && match(r, f, false); }).length;
  kpis("kEst", [
    [fi(s.regs), "registros na posição"], [fi(pns), "PNs diferentes"], [fi(bases), "bases"],
    [fd(s.dia), "data da posição"], [fi(opp.length), "não atendem com estoque em alguma base", "warn"], [fi(transf), "linhas aguardando transferência"]
  ]);
  bars("byStation", Object.keys(s.byStation).map(function(k){ return [k, s.byStation[k]]; }).sort(function(a,b){ return b[1]-a[1]; }).slice(0, 12), "var(--navy)");
  renderPnQuery();
  $("oppCount").textContent = S.rows.length ? pl(opp.length, "linha que não atende", "linhas que não atendem") + ", mas o PN tem estoque em alguma base." : "Carregue também a planilha de pré-indução para cruzar o estoque com as linhas que não atendem.";
  table("tOpp", [
    { k: "acft", label: "Acft" }, { k: "nec", label: "Necessidade", txt: function(r){ return fd(r.nec); } },
    { k: "pn", label: "PN" }, { k: "desc", label: "Descrição", cls: "desc" }, { k: "qty", label: "Qtd necessária", num: true, txt: qtyTxt },
    { k: "estTot", label: "Estoque total", num: true, val: function(r){ return stockTotal(r.pn); }, txt: function(r){ return fi(stockTotal(r.pn)); } },
    { k: "estBases", label: "Onde está", cls: "stockcell", txt: function(r){ return stockTxt(r.pn); } },
    { k: "resp", label: "Responsável" }, { k: "order", label: "Pedido" }
  ], opp, { key: "nec" });
}
function renderPnQuery(){
  var q = upn($("pnQuery").value);
  if (!q) { $("pnResult").innerHTML = '<p class="empty">Digite um PN para ver em quais bases ele está.</p>'; return; }
  var st = stockOf(q);
  var dem = S.rows.filter(function(r){ return r.gone == null && upn(r.pn) === q; });
  var dn = dem.filter(function(r){ return r.st === "n"; }).length;
  var html = st ? "<p><b>" + esc(q) + "</b> " + esc(S.stock.desc[q] || "") + "</p>" + '<div class="hbars" id="pnBars"></div>' : '<p class="empty">O PN ' + esc(q) + " não aparece na posição de estoque carregada.</p>";
  if (dem.length) html += '<p class="hint" style="margin-top:10px">Na pré-indução: ' + pl(dem.length, "linha ativa", "linhas ativas") + ", " + pl(dn, "não atende", "não atendem") + ".</p>";
  $("pnResult").innerHTML = html;
  if (st) bars("pnBars", st.map(function(x){ return [x[0], x[1]]; }), "var(--disp)");
}

/* ================= histórico: cópia e restauração ================= */
function exportHist(){
  var data = { versao: 2, gerado: new Date().toISOString(), meta: S.meta, rows: S.rows.map(pack), cargas: S.cargas, stock: S.stock };
  download(JSON.stringify(data), "historico-preinducao-" + new Date().toISOString().slice(0,10) + ".json", "application/json");
}
function importHist(file){
  var rd = new FileReader();
  rd.onload = function(){
    try {
      var d = JSON.parse(rd.result);
      if (!d || d.versao !== 2 || !Array.isArray(d.rows)) throw new Error("Este arquivo não é uma cópia do histórico deste painel.");
      if (!window.confirm("Substituir o histórico deste navegador pela cópia de " + (d.gerado ? new Date(d.gerado).toLocaleString("pt-BR") : "arquivo") + "?")) return;
      S.meta = d.meta; S.rows = d.rows.map(unpack); S.cargas = d.cargas || []; S.stock = d.stock || null;
      overlay("Restaurando…");
      Promise.all([savePre(), S.stock ? saveStock() : dbDel("estoque")]).then(function(){ overlay(null); showDash(); });
    } catch(e){ window.alert(e.message || String(e)); }
  };
  rd.readAsText(file);
}
function clearAll(){
  if (!window.confirm("Apagar todo o histórico e a posição de estoque guardados neste navegador? Baixe uma cópia antes, se ainda precisar.")) return;
  Promise.all([dbDel("pre"), dbDel("estoque")]).then(function(){ S.rows = []; S.cargas = []; S.meta = null; S.stock = null; showDash(); });
}

/* ================= eventos ================= */
function setTab(t){ S.tab = ["visao","preinducao","pedidos","estoque"].indexOf(t) >= 0 ? t : "visao"; render(); }
window.addEventListener("hashchange", function(){ setTab(location.hash.slice(1)); });
$("loadBtn").addEventListener("click", function(){ $("fileInput").click(); });
$("fileInput").addEventListener("change", function(e){ handleFiles(e.target.files); e.target.value = ""; });
document.addEventListener("click", function(e){
  if (e.target.closest(".drop")) { $("fileInput").click(); return; }
  var ev = e.target.closest("[data-ev]");
  if (ev) { var p = ev.getAttribute("data-ev"); $("fEvento").value = $("fEvento").value === p ? "" : p; resetPages(); render(); return; }
  var pn = e.target.closest("tr[data-pn]");
  if (pn) { $("fBusca").value = pn.getAttribute("data-pn"); location.hash = "preinducao"; resetPages(); setTab("preinducao"); return; }
  var q = e.target.closest("tr[data-q]");
  if (q) { $("fBusca").value = q.getAttribute("data-q"); $("fSit").value = ""; resetPages(); render(); return; }
  var csv = e.target.closest("[data-csv]");
  if (csv) { exportCsv(csv.getAttribute("data-csv")); return; }
  if (!e.target.closest(".menu")) closeMenu();
});
document.addEventListener("keydown", function(e){
  if ((e.key === "Enter" || e.key === " ") && e.target.classList && e.target.classList.contains("drop")) { e.preventDefault(); $("fileInput").click(); }
  if (e.key === "Escape") { if ($("confirm").classList.contains("on")) { pending = null; closeConfirm(); nextFile(); } closeMenu(); }
});
["dragenter","dragover"].forEach(function(t){ document.addEventListener(t, function(e){ e.preventDefault(); document.querySelectorAll(".drop").forEach(function(d){ d.classList.add("over"); }); }); });
["dragleave","drop"].forEach(function(t){ document.addEventListener(t, function(e){ e.preventDefault(); document.querySelectorAll(".drop").forEach(function(d){ d.classList.remove("over"); }); }); });
document.addEventListener("drop", function(e){ if (e.dataTransfer && e.dataTransfer.files.length) handleFiles(e.dataTransfer.files); });
$("cfCancel").addEventListener("click", function(){ pending = null; closeConfirm(); nextFile(); });
$("cfSave").addEventListener("click", saveCarga);
function resetPages(){ Object.keys(TB).forEach(function(k){ TB[k].page = 0; }); }
["fEvento","fResp","fAtiv","fStatus","fMarco","fMud","fSit","fTipo"].forEach(function(id){ $(id).addEventListener("change", function(){ resetPages(); render(); }); });
var tmr; $("fBusca").addEventListener("input", function(){ clearTimeout(tmr); tmr = setTimeout(function(){ resetPages(); render(); }, 220); });
var tmr2; $("pnQuery").addEventListener("input", function(){ clearTimeout(tmr2); tmr2 = setTimeout(renderPnQuery, 200); });
$("resetBtn").addEventListener("click", function(){ ["fEvento","fResp","fAtiv"].forEach(function(id){ $(id).value = ""; }); $("fBusca").value = ""; resetPages(); render(); });
["tPre","tPed","tSemPed","tOpp"].forEach(wireTable);
function closeMenu(){ $("menuList").classList.add("hidden"); $("menuBtn").setAttribute("aria-expanded","false"); }
$("menuBtn").addEventListener("click", function(e){ e.stopPropagation(); var open = $("menuList").classList.toggle("hidden") === false; $("menuBtn").setAttribute("aria-expanded", String(open)); });
$("exportBtn").addEventListener("click", function(){ closeMenu(); exportHist(); });
$("importBtn").addEventListener("click", function(){ closeMenu(); $("importInput").click(); });
$("importInput").addEventListener("change", function(e){ if (e.target.files[0]) importHist(e.target.files[0]); e.target.value = ""; });
$("clearBtn").addEventListener("click", function(){ closeMenu(); clearAll(); });
$("themeBtn").addEventListener("click", function(){
  var cur = document.documentElement.getAttribute("data-theme");
  var dark = cur ? cur === "dark" : window.matchMedia("(prefers-color-scheme: dark)").matches;
  var next = dark ? "light" : "dark"; document.documentElement.setAttribute("data-theme", next);
  try { localStorage.setItem("preinducao-tema", next); } catch(e){}
});
try { var th = localStorage.getItem("preinducao-tema"); if (th) document.documentElement.setAttribute("data-theme", th); } catch(e){}

/* ================= início ================= */
S.tab = ["visao","preinducao","pedidos","estoque"].indexOf(location.hash.slice(1)) >= 0 ? location.hash.slice(1) : "visao";
Promise.all([dbGet("pre"), dbGet("estoque")]).then(function(v){
  if (v[0] && v[0].rows) { S.meta = v[0].meta; S.rows = v[0].rows.map(unpack); S.cargas = v[0].cargas || []; }
  if (v[1]) S.stock = v[1];
  showDash();
});
window.__pi = { S: S, merge: merge };
})();

</script>
</body>
</html>
