<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Pré-indução: painel gerencial</title>
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

/* ===== painel gerencial ===== */
.kcards{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px;margin-top:16px}
.kcard{background:var(--surface);border:1px solid var(--line);border-top:4px solid var(--line);border-radius:8px;padding:12px 14px 12px;min-width:0}
.kcard.good{border-top-color:var(--disp)} .kcard.warn{border-top-color:var(--atende)} .kcard.bad{border-top-color:var(--nao)}
.kcard h4{margin:0;font-size:13.5px;font-weight:600;line-height:1.25}
.kcard .v{font-family:var(--cond);font-size:36px;font-weight:700;line-height:1.05;margin-top:6px}
.kcard.good .v{color:var(--disp)} .kcard.warn .v{color:var(--atende)} .kcard.bad .v{color:var(--nao)}
.kcard .meta{font-size:12.5px;color:var(--muted);margin-top:2px}
.kcard .delta{font-size:12.5px;margin-top:6px;font-weight:600}
.kcard .delta.up{color:var(--disp)} .kcard .delta.down{color:var(--nao)} .kcard .delta.eq{color:var(--muted)}
.kcard .ctx{font-size:12.5px;color:var(--muted);margin-top:6px;line-height:1.35}
.actions-list{display:flex;flex-direction:column}
.act{display:grid;grid-template-columns:6px minmax(0,1fr) auto;gap:12px;padding:10px 0;border-top:1px solid var(--line);align-items:start}
.act:first-child{border-top:0}
.act .sev{align-self:stretch;border-radius:3px;background:var(--atende)}
.act.alta .sev{background:var(--nao)}
.act b{display:block;font-size:14.5px;line-height:1.3}
.act span.d{display:block;color:var(--muted);font-size:13px;margin-top:2px;line-height:1.35}
.act .area{font-size:12px;font-weight:600;border:1px solid var(--line);border-radius:10px;padding:1px 8px;white-space:nowrap;color:var(--ink)}
.act-filter{display:flex;gap:6px;flex-wrap:wrap;margin:-2px 0 10px}
.chip{border:1px solid var(--line);background:var(--surface);border-radius:14px;padding:3px 10px;font-size:12.5px;cursor:pointer}
.chip[aria-pressed="true"]{background:var(--ink);color:var(--bg);border-color:var(--ink)}
.chart{width:100%;display:block}
.chart text{fill:var(--muted);font-family:var(--font);font-size:11px}
.chart .lbl{fill:var(--ink);font-weight:600}
.chart .meta-line{stroke:var(--navy);stroke-dasharray:4 3;stroke-width:1.5}
.chart .grid{stroke:var(--line)}
.reg{font-size:12px;font-weight:600}
.reg.Regular{color:var(--nao)} .reg.Irregular{color:var(--atende)}
.who{font-size:12.5px;color:var(--muted)}
.upload-hint{font-size:13px;color:var(--muted)}
@media (max-width:1100px){ .kcards{grid-template-columns:repeat(3,minmax(0,1fr))} }
@media (max-width:640px){ .kcards{grid-template-columns:repeat(2,minmax(0,1fr))} .kcard .v{font-size:30px} }
.badge{display:inline-block;font-size:12px;font-weight:600;padding:3px 9px;border-radius:10px;border:1px solid var(--line);color:var(--muted)}
.badge.cloud{border-color:var(--disp);color:var(--disp)}
.badge:empty{display:none}
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <div>
      <h1>Pré-indução: painel gerencial</h1>
      <div class="sub" id="subInfo">Nenhuma planilha carregada ainda</div>
    </div>
    <div class="actions">
      <span class="badge" id="storeBadge"></span>
      <button class="btn" id="themeBtn" type="button">Tema</button>
      <button class="btn primary" id="loadBtn" type="button">Atualizar com a planilha da semana</button>
      <input type="file" id="fileInput" accept=".xlsx,.xlsm" class="hidden">
    </div>
  </header>
  <div class="tabs" style="height:0;border-bottom-width:2px"></div>
  <div class="banner hidden" id="newBanner">Outra pessoa atualizou o painel. <button class="btn small" id="reloadBtn" type="button">Ver a versão nova</button></div>
  <div class="banner hidden" id="localBanner">Este painel está aberto fora do Claude: os dados ficam só neste navegador. Para a gerência ver a mesma versão, use o link publicado no Claude.</div>

  <section id="empty" class="drop" tabindex="0" aria-label="Área para soltar a planilha">
    <h2>Solte aqui a planilha de pré-indução da semana</h2>
    <p>O painel lê as abas Caminho Crítico, Base_evolutivo, Regularidade e 1328 AMOS e calcula os indicadores gerenciais. A planilha é grande: a leitura pode levar até 1 minuto.</p>
    <p class="msg" id="emptyMsg"></p>
  </section>

  <main id="dash" class="hidden">
    <section class="kcards" id="kcards" aria-label="Indicadores"></section>

    <div class="grid2">
      <section class="panel">
        <div class="panel-head">
          <div><h3>Pontos de ação</h3><p class="hint">Gerados a partir da planilha: o que precisa de decisão ou cobrança agora.</p></div>
        </div>
        <div class="act-filter" id="actFilter"></div>
        <div class="actions-list" id="actions"></div>
      </section>
      <section class="panel">
        <h3>Próximas induções</h3>
        <p class="hint">Posição de cada aeronave até a entrada no hangar. Faixa azul: zona da meta (T-60).</p>
        <div class="tl-legend">
          <span><i style="background:var(--disp)"></i>100% atendido</span>
          <span><i style="background:var(--atende)"></i>Pendente, fora da meta</span>
          <span><i style="background:var(--nao)"></i>Pendente dentro de T-60</span>
        </div>
        <div class="tl" id="timeline"></div>
      </section>
    </div>

    <div class="grid2">
      <section class="panel">
        <h3>Disponibilidade na indução, mês a mês</h3>
        <p class="hint">Média do % de material disponível na última atualização antes de cada indução (Base_evolutivo). Linha tracejada: meta de 98%.</p>
        <svg class="chart" id="chMes" viewBox="0 0 600 220" role="img" aria-label="Disponibilidade na indução por mês"></svg>
      </section>
      <section class="panel">
        <h3>Aderência ao T-60, por trimestre</h3>
        <p class="hint">% das aeronaves que chegaram ao T-60 com 100% da demanda atendida.</p>
        <svg class="chart" id="chT60" viewBox="0 0 600 220" role="img" aria-label="Aderência ao T-60 por trimestre"></svg>
      </section>
    </div>

    <div class="grid2">
      <section class="panel">
        <h3>Gargalos por área</h3>
        <p class="hint">Linhas que não atendem e há quantos dias, em média, estão pendentes.</p>
        <div id="resp"></div>
      </section>
      <section class="panel">
        <h3>Pedidos que mais travam induções</h3>
        <p class="hint">Pedidos atrasados ou sem previsão, pelo número de linhas afetadas.</p>
        <div id="orders"></div>
      </section>
    </div>

    <div class="grid2">
      <section class="panel">
        <h3>Itens críticos e perfil de consumo</h3>
        <p class="hint">PNs que mais faltam, com a classificação da aba Regularidade. Item regular que falta com frequência indica estoque mínimo a revisar.</p>
        <div id="pns"></div>
      </section>
      <section class="panel">
        <h3>Fornecedores com itens vencidos</h3>
        <p class="hint">Itens abertos de pedidos da pré-indução com data confirmada ou target vencida (1328 AMOS).</p>
        <div class="hbars" id="vend"></div>
      </section>
    </div>

    <section class="panel full">
      <h3>Últimas induções</h3>
      <p class="hint">Disponibilidade de material registrada na véspera de cada entrada no hangar.</p>
      <div id="inducoes"></div>
    </section>

    <p class="note" id="note"></p>
  </main>
</div>

<div class="overlay" id="overlay" role="status" aria-live="polite"><div id="overlayMsg">Lendo planilha…</div></div>

<script>
(function(){
"use strict";
var DAY = 86400000;
var $ = function(id){ return document.getElementById(id); };
var S = { snap: null, resumos: [], pend: {}, snapId: null, actArea: "" };
var store = null;

/* ================= utilidades ================= */
function norm(s){ return String(s == null ? "" : s).normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s+/g," ").trim().toLowerCase(); }
function str(v){ return v == null ? "" : String(v).trim(); }
function upn(v){ return str(v).toUpperCase(); }
function isSerial(v){ return typeof v === "number" && v > 30000 && v < 80000; }
function toDay(v){ return isSerial(v) ? Math.floor(v) - 25569 : null; }
function today(){ var n = new Date(); return Math.floor(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) / DAY); }
function fd(d, short){ if (d == null) return ""; var x = new Date(d * DAY), dd = String(x.getUTCDate()).padStart(2,"0"), mm = String(x.getUTCMonth()+1).padStart(2,"0"); return short ? dd + "/" + mm : dd + "/" + mm + "/" + x.getUTCFullYear(); }
var MES = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
function monthKey(d){ var x = new Date(d * DAY); return x.getUTCFullYear() + "-" + String(x.getUTCMonth()+1).padStart(2,"0"); }
function monthLbl(k){ var p = k.split("-"); return MES[parseInt(p[1],10)-1] + "/" + p[0].slice(2); }
function quarterKey(d){ var x = new Date(d * DAY); return x.getUTCFullYear() + "T" + (Math.floor(x.getUTCMonth()/3)+1); }
function pct(a,b){ return b ? a/b*100 : 0; }
function f1(v){ return v == null ? "–" : v.toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1}); }
function fi(n){ return Number(n || 0).toLocaleString("pt-BR"); }
function esc(s){ return String(s == null ? "" : s).replace(/[&<>"]/g, function(c){ return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]; }); }
function pl(n, a, b){ return fi(n) + " " + (n === 1 ? a : b); }
function statusOf(v){ var s = norm(v); if (s.charAt(0) === "1" || s.indexOf("disponivel") >= 0) return "d"; if (s.charAt(0) === "2" || s === "atende") return "a"; return "n"; }
function prevTipo(raw){
  if (isSerial(raw)) return "data";
  var s = norm(raw);
  if (!s || s.indexOf("sem previs") >= 0) return "sem";
  if (s.indexOf("transferir") >= 0) return "transferir";
  if (s.indexOf("estoque") === 0) return "estoque";
  return "outro";
}
function tipoPedido(order, ext){
  var o = upn(order), x = str(ext);
  if (!o) return "";
  if (/TRF$/.test(o) || /^45/.test(x)) return "PO de compra";
  if (/^MP/.test(o)) return "Produção (MP)";
  if (/^B/.test(o)) return "Requisição (RC)";
  if (/^R/.test(o)) return "Reparo (RO)";
  return "Outro";
}
function marcoAtual(days){ if (days == null) return ""; if (days < 0) return "em execução"; if (days <= 7) return "T-7"; if (days <= 45) return "T-45"; if (days <= 60) return "T-60"; if (days <= 120) return "T-120"; return "T-180"; }
function header(aoa, must, maxRows){
  for (var i = 0; i < Math.min(maxRows || 15, aoa.length); i++) {
    var hs = (aoa[i] || []).map(norm);
    if (must.every(function(m){ return hs.indexOf(m) >= 0; })) return { row: i, head: hs };
  }
  return null;
}
function col(head, names){ for (var i = 0; i < names.length; i++) { var j = head.indexOf(names[i]); if (j >= 0) return j; } return -1; }
function findSheet(wb, name){ return wb.SheetNames.filter(function(n){ return norm(n) === norm(name); })[0]; }
function aoaOf(wb, name){ var n = findSheet(wb, name); return n && wb.Sheets[n] ? XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, blankrows: false }) : null; }

/* ================= leitura das abas ================= */
function readCC(aoa){
  var h = header(aoa, ["atendimento","pacote"]);
  if (!h) throw new Error("Não encontrei o cabeçalho da aba Caminho Crítico (colunas Atendimento e Pacote).");
  var H = h.head, c = {
    prog: col(H,["programacao"]), marco: col(H,["demanda"]), wo: col(H,["wo"]), acft: col(H,["acft"]), pn: col(H,["pn"]), desc: col(H,["description"]),
    nec: col(H,["data necessidade"]), prev: col(H,["previsao na cia"]), atend: col(H,["atendimento"]), resp: col(H,["responsavel"]),
    order: col(H,["order no. (oh_orderno)"]), ext: col(H,["ext. order (oh_ext_orderno)"]), pacote: col(H,["pacote"]), ativ: col(H,["atividade"])
  };
  var out = [], seen = {};
  for (var r = h.row + 1; r < aoa.length; r++) {
    var a = aoa[r]; if (!a || (a[c.pn] == null && a[c.pacote] == null)) continue;
    var g = function(k){ return c[k] >= 0 ? a[c[k]] : null; };
    var x = { pacote: str(g("pacote")), acft: str(g("acft")), wo: str(g("wo")), pn: str(g("pn")), desc: str(g("desc")),
      nec: toDay(g("nec")), prog: toDay(g("prog")), marco: str(g("marco")) || "?", prevDay: toDay(g("prev")), prevTipo: prevTipo(g("prev")),
      st: statusOf(g("atend")), resp: str(g("resp")) || "Sem responsável", order: str(g("order")), ext: str(g("ext")), ativ: str(g("ativ")).replace(/^[:\s]+/, "") };
    x.tipo = tipoPedido(x.order, x.ext);
    var b = x.pacote + "|" + x.wo + "|" + x.pn; seen[b] = (seen[b] || 0) + 1; x.k = seen[b] > 1 ? b + "#" + seen[b] : b;
    out.push(x);
  }
  if (!out.length) throw new Error("A aba Caminho Crítico está vazia.");
  return out;
}
function readBE(aoa){
  if (!aoa) return [];
  var h = header(aoa, ["data da atualizacao","wp"]); if (!h) return [];
  var H = h.head, c = { d: col(H,["data da atualizacao"]), acft: col(H,["acft"]), ini: col(H,["inicio"]), wp: col(H,["wp"]), disp: col(H,["% disponivel"]), tot: col(H,["qtde eventos"]), t180: col(H,["t180"]), at: col(H,["atendimento"]) };
  var out = [];
  for (var r = h.row + 1; r < aoa.length; r++) {
    var a = aoa[r]; if (!a) continue;
    var d = toDay(a[c.d]), ini = toDay(a[c.ini]); if (d == null || ini == null || !a[c.wp]) continue;
    var disp = typeof a[c.disp] === "number" ? a[c.disp] : null; if (disp != null && disp > 1.5) disp = disp / 100;
    var tot = typeof a[c.tot] === "number" ? a[c.tot] : null, at = c.at >= 0 && typeof a[c.at] === "number" ? a[c.at] : null;
    out.push({ d: d, acft: str(a[c.acft]), ini: ini, wp: str(a[c.wp]), disp: disp, tot: tot, at: at, t180: typeof a[c.t180] === "number" ? a[c.t180] : null });
  }
  return out;
}
function readReg(aoa){
  var m = {}; if (!aoa) return m;
  var h = header(aoa, ["p/n","regularidade"]); if (!h) return m;
  var cp = h.head.indexOf("p/n"), cr = h.head.indexOf("regularidade"), cc = col(h.head, ["consumo mes"]);
  for (var r = h.row + 1; r < aoa.length; r++) { var a = aoa[r]; if (a && a[cp] != null) m[upn(a[cp])] = [str(a[cr]), typeof a[cc] === "number" ? a[cc] : null]; }
  return m;
}
function read1328(aoa){
  if (!aoa) return [];
  var h = header(aoa, ["order no.","detail state"]); if (!h) return [];
  var H = h.head, c = { o: col(H,["order no."]), st: col(H,["detail state"]), name: col(H,["name"]), bo: col(H,["backorder"]), tg: col(H,["target date"]), cf: col(H,["confirmed date:","confirmed date"]) };
  var out = [];
  for (var r = h.row + 1; r < aoa.length; r++) {
    var a = aoa[r]; if (!a || !a[c.o]) continue;
    out.push({ o: upn(a[c.o]), st: upn(a[c.st]), name: str(a[c.name]), bo: typeof a[c.bo] === "number" ? a[c.bo] : 0, due: toDay(a[c.cf]) != null ? toDay(a[c.cf]) : toDay(a[c.tg]) });
  }
  return out;
}

/* ================= cálculo do retrato gerencial ================= */
function isBack(r){ return !!r.order && (r.prevTipo === "sem" || (r.prevDay != null && r.nec != null && r.prevDay > r.nec)); }
function build(rows, be, reg, ord, oldPend, file){
  var t0 = today();
  /* eventos */
  var ev = {};
  rows.forEach(function(r){
    var e = ev[r.pacote] || (ev[r.pacote] = { p: r.pacote, acftC: {}, ativC: {}, prog: null, t: 0, d: 0, a: 0, n: 0, fora: 0, transf: 0, resp: {} });
    e.t++; e[r.st]++; e.acftC[r.acft] = (e.acftC[r.acft] || 0) + 1; e.ativC[r.ativ] = (e.ativC[r.ativ] || 0) + 1;
    if (r.prog != null && (e.prog == null || r.prog < e.prog)) e.prog = r.prog;
    if (r.marco && r.marco.toUpperCase() !== "T180") e.fora++;
    if (r.prevTipo === "transferir") e.transf++;
    if (r.st === "n") e.resp[r.resp] = (e.resp[r.resp] || 0) + 1;
  });
  var top = function(o){ var b = "", n = -1; Object.keys(o).forEach(function(k){ if (o[k] > n) { n = o[k]; b = k; } }); return b; };
  var events = Object.keys(ev).map(function(k){
    var e = ev[k]; e.acft = top(e.acftC); e.ativ = top(e.ativC); delete e.acftC; delete e.ativC;
    e.days = e.prog != null ? e.prog - t0 : null; return e;
  }).filter(function(e){ return e.days == null || e.days >= -60; }).sort(function(a,b){ return (a.prog||0) - (b.prog||0); });

  /* pendências e tempo */
  var pend = {}, idades = [], respAge = {}, resolvidas = 0, tRes = [];
  rows.forEach(function(r){
    if (r.st !== "n") return;
    var first = oldPend[r.k] != null ? oldPend[r.k] : t0;
    pend[r.k] = first; idades.push(t0 - first);
    var ra = respAge[r.resp] || (respAge[r.resp] = { n: 0, soma: 0 }); ra.n++; ra.soma += t0 - first;
  });
  var hadOld = Object.keys(oldPend).length > 0;
  Object.keys(oldPend).forEach(function(k){ if (pend[k] == null) { resolvidas++; tRes.push(t0 - oldPend[k]); } });
  var avg = function(a){ return a.length ? a.reduce(function(s,x){ return s + x; }, 0) / a.length : null; };

  /* KPIs atuais */
  var in60 = events.filter(function(e){ return e.days != null && e.days >= 0 && e.days <= 60; });
  var ok60 = in60.filter(function(e){ return e.n === 0; }).length;
  var nao60 = events.filter(function(e){ return e.days != null && e.days <= 60; }).reduce(function(s,e){ return s + e.n; }, 0);
  var comPed = rows.filter(function(r){ return r.order; }), back = comPed.filter(isBack);
  var fora = rows.filter(function(r){ return r.marco && r.marco.toUpperCase() !== "T180"; }).length;

  /* histórico Base_evolutivo */
  var byWp = {};
  be.forEach(function(x){ (byWp[x.wp] = byWp[x.wp] || []).push(x); });
  var ind = [], t60h = [];
  Object.keys(byWp).forEach(function(wp){
    var xs = byWp[wp].sort(function(a,b){ return a.d - b.d; }), ini = xs[0].ini;
    if (ini > t0) return;
    var before = xs.filter(function(x){ return x.d <= ini && x.disp != null; });
    if (before.length) { var l = before[before.length-1]; ind.push({ wp: wp, acft: l.acft, ini: ini, disp: l.disp, d: l.d }); }
    var near = xs.filter(function(x){ var dd = ini - x.d; return dd >= 50 && dd <= 70; }).sort(function(a,b){ return Math.abs(ini - a.d - 60) - Math.abs(ini - b.d - 60); })[0];
    if (near) { var at = near.at != null && near.tot ? near.at / near.tot : near.disp; if (at != null) t60h.push({ wp: wp, ini: ini, ok: at >= 0.999 }); }
  });
  ind.sort(function(a,b){ return a.ini - b.ini; });
  var win = function(from, to){ return ind.filter(function(x){ return x.ini > t0 - from && x.ini <= t0 - to; }); };
  var w1 = win(90, 0), w0 = win(180, 90);
  var mdisp = function(a){ return a.length ? avg(a.map(function(x){ return x.disp * 100; })) : null; };
  var hm = {};
  ind.forEach(function(x){ var k = monthKey(x.ini); var m = hm[k] || (hm[k] = { n: 0, s: 0, n98: 0 }); m.n++; m.s += x.disp * 100; if (x.disp >= 0.98) m.n98++; });
  var histMes = Object.keys(hm).sort().slice(-18).map(function(k){ return { m: k, n: hm[k].n, v: hm[k].s / hm[k].n, n98: hm[k].n98 }; });
  var hq = {};
  t60h.forEach(function(x){ var k = quarterKey(x.ini); var q = hq[k] || (hq[k] = { n: 0, ok: 0 }); q.n++; if (x.ok) q.ok++; });
  var histT60 = Object.keys(hq).sort().slice(-8).map(function(k){ return { q: k, n: hq[k].n, ok: hq[k].ok }; });

  /* gargalos por área */
  var resp = Object.keys(respAge).map(function(k){ return { r: k, n: respAge[k].n, idade: hadOld ? respAge[k].soma / respAge[k].n : null }; }).sort(function(a,b){ return b.n - a.n; });

  /* pedidos */
  var og = {};
  back.forEach(function(r){
    var o = og[r.order] || (og[r.order] = { o: r.order, ext: r.ext, tipo: r.tipo, linhas: 0, acft: {}, sem: r.prevTipo === "sem", atraso: null, nec: null });
    o.linhas++; o.acft[r.acft] = 1;
    if (r.prevDay != null && r.nec != null) { var x = r.prevDay - r.nec; if (o.atraso == null || x > o.atraso) o.atraso = x; }
    if (r.nec != null && (o.nec == null || r.nec < o.nec)) o.nec = r.nec;
  });
  var orders = Object.keys(og).map(function(k){ var o = og[k]; o.nAcft = Object.keys(o.acft).length; o.acft = Object.keys(o.acft).slice(0, 6); return o; })
    .sort(function(a,b){ return b.linhas - a.linhas || (b.atraso||0) - (a.atraso||0); }).slice(0, 10);

  /* itens críticos x regularidade */
  var pg = {};
  rows.forEach(function(r){
    if (r.st !== "n") return;
    var p = pg[r.pn] || (pg[r.pn] = { pn: r.pn, desc: r.desc, linhas: 0, acft: {}, nec: null, semPed: 0 });
    p.linhas++; p.acft[r.acft] = 1; if (!r.order) p.semPed++;
    if (r.nec != null && (p.nec == null || r.nec < p.nec)) p.nec = r.nec;
  });
  var pnsAll = Object.keys(pg).map(function(k){ var p = pg[k], rg = reg[upn(p.pn)]; p.nAcft = Object.keys(p.acft).length; delete p.acft; p.reg = rg ? rg[0] : ""; p.cons = rg ? rg[1] : null; return p; })
    .sort(function(a,b){ return b.linhas - a.linhas || (a.nec||0) - (b.nec||0); });
  var pns = pnsAll.slice(0, 10);

  /* fornecedores (1328) ligados à pré-indução */
  var ordSet = {}; rows.forEach(function(r){ if (r.order) ordSet[upn(r.order)] = 1; });
  var vg = {}, abertos = 0, vencidos = 0;
  ord.forEach(function(x){
    if (!ordSet[x.o]) return;
    if (["CLOSED","RECEIVED","CANCELLED"].indexOf(x.st) >= 0) return;
    abertos++;
    if (x.due != null && x.due < t0) {
      vencidos++;
      var nm = x.name || "Sem fornecedor"; if (norm(nm) === "req. spot") nm = "Requisição sem fornecedor (REQ. SPOT)";
      vg[nm] = (vg[nm] || 0) + 1;
    }
  });
  var vend = Object.keys(vg).map(function(k){ return { name: k, n: vg[k] }; }).sort(function(a,b){ return b.n - a.n; }).slice(0, 8);

  /* pontos de ação */
  var acts = [];
  in60.concat(events.filter(function(e){ return e.days != null && e.days < 0; })).forEach(function(e){
    if (!e.n) return;
    var rs = Object.keys(e.resp).sort(function(a,b){ return e.resp[b] - e.resp[a]; }).slice(0, 3).map(function(k){ return k.replace(/^\d+-/, "") + " " + e.resp[k]; }).join(", ");
    acts.push({ sev: e.days <= 30 || e.n >= 50 ? "alta" : "media", ord: e.days, area: "Pré-indução",
      t: e.acft + (e.days < 0 ? " em execução" : " entra em " + fd(e.prog, true) + " (" + marcoAtual(e.days) + ")") + " com " + pl(e.n, "linha sem atendimento", "linhas sem atendimento"),
      d: "Atendido: " + Math.round(pct(e.d + e.a, e.t)) + "%. Pendências por área: " + rs + "." });
  });
  events.forEach(function(e){
    if (e.days != null && e.days >= 0 && e.days <= 10 && e.transf > 0)
      acts.push({ sev: "alta", ord: e.days - 0.5, area: "Manutenção", t: "Pedir reserva e transferência de " + pl(e.transf, "linha", "linhas") + " para " + e.acft, d: "Indução em " + fd(e.prog, true) + ". Pela regra do manual, a transferência deve ser solicitada ao AOGDESK entre T-7 e T-3." });
  });
  orders.slice(0, 4).forEach(function(o){
    if (o.linhas < 5) return;
    acts.push({ sev: o.nec != null && o.nec - t0 <= 60 ? "alta" : "media", ord: 100 + (o.nec != null ? o.nec - t0 : 999), area: "Compras",
      t: o.tipo + " " + o.o + (o.sem ? " sem previsão" : " atrasado " + pl(o.atraso, "dia", "dias")) + " trava " + pl(o.linhas, "linha", "linhas"),
      d: "Afeta " + pl(o.nAcft, "aeronave", "aeronaves") + " (" + o.acft.join(", ") + (o.nAcft > o.acft.length ? "…" : "") + "). Primeira necessidade em " + fd(o.nec) + "." });
  });
  pnsAll.filter(function(p){ return norm(p.reg) === "regular" && p.nAcft >= 2; }).slice(0, 3).forEach(function(p){
    acts.push({ sev: "media", ord: 300, area: "Planejamento", t: "Revisar estoque mínimo de " + p.pn + " (" + p.desc + ")", d: "Consumo regular e falta em " + pl(p.nAcft, "aeronave", "aeronaves") + " (" + pl(p.linhas, "linha", "linhas") + ")." });
  });
  var semPedPerto = rows.filter(function(r){ return r.st === "n" && !r.order && r.nec != null && r.nec - t0 <= 45; });
  if (semPedPerto.length) acts.push({ sev: "alta", ord: 50, area: "Planejamento", t: pl(semPedPerto.length, "linha sem pedido", "linhas sem pedido") + " com necessidade nos próximos 45 dias", d: "Não atendem e ainda não têm PO, requisição, reparo ou ordem de produção." });
  events.forEach(function(e){
    if (e.days != null && e.days >= 0 && e.days <= 120 && e.t >= 100 && e.fora / e.t > 0.3)
      acts.push({ sev: "media", ord: 400 + e.days, area: "Planej. Manutenção", t: Math.round(pct(e.fora, e.t)) + "% da demanda de " + e.acft + " entrou depois do T-180", d: "Alinhar com o Planejamento de Manutenção a liberação antecipada do pacote " + e.p + "." });
  });
  if (vend.length && vend[0].n >= 10) acts.push({ sev: "media", ord: 500, area: "Compras", t: vend[0].name + " com " + pl(vend[0].n, "item vencido", "itens vencidos"), d: "Itens de pedidos que abastecem a pré-indução, com data confirmada ou target já ultrapassada." });
  acts.sort(function(a,b){ return (a.sev === b.sev ? 0 : a.sev === "alta" ? -1 : 1) || a.ord - b.ord; });

  var kpi = {
    dispInd: { v: mdisp(w1), n: w1.length, n98: w1.filter(function(x){ return x.disp >= 0.98; }).length, prev: mdisp(w0) },
    t60: { ok: ok60, n: in60.length },
    nao60: nao60,
    back: { n: back.length, base: comPed.length },
    fora: { n: fora, base: rows.length },
    idade: hadOld ? avg(idades) : null,
    res: { n: hadOld ? resolvidas : null, t: avg(tRes) },
    venc: { n: vencidos, base: abertos }
  };
  var snap = { v: 1, ts: Date.now(), dia: t0, arquivo: file, linhas: rows.length, kpi: kpi, events: events, actions: acts.slice(0, 30),
    histMes: histMes, histT60: histT60, inducoes: ind.slice(-12).reverse(), resp: resp, orders: orders, pns: pns, vend: vend, beN: be.length };
  return { snap: snap, pend: pend };
}
function resumoOf(s){
  var k = s.kpi;
  return { ts: s.ts, dia: s.dia, arquivo: s.arquivo, dispInd: k.dispInd.v, t60: k.t60.n ? pct(k.t60.ok, k.t60.n) : null, nao60: k.nao60,
    back: k.back.base ? pct(k.back.n, k.back.base) : null, fora: k.fora.base ? pct(k.fora.n, k.fora.base) : null, idade: k.idade };
}

/* ================= armazenamento ================= */
function sleep(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }
function retry(fn){ var n = 0; return (function go(){ return fn().catch(function(e){ n++; if (n < 4 && e && (e.code === "unavailable" || e.code === "resource_exhausted")) return sleep(700 * n).then(go); throw e; }); })(); }
function cloudStore(db){
  return { kind: "cloud",
    load: function(){
      return retry(function(){ return db.doc("estado/atual").get(); }).then(function(s){
        if (!s.exists) return { snap: null, resumos: [], pend: {}, id: null };
        var id = s.data().snapId;
        return Promise.all([
          retry(function(){ return db.doc("snaps/" + id).get(); }),
          retry(function(){ return db.doc("pend/atual").get(); }),
          retry(function(){ return db.collection("resumos").orderBy("ts","desc").limit(60).get(); })
        ]).then(function(v){
          return { snap: v[0].exists ? v[0].data() : null, pend: v[1].exists ? (v[1].data().map || {}) : {}, resumos: v[2].docs.map(function(d){ return d.data(); }).reverse(), id: id };
        });
      });
    },
    save: function(snap, pend){
      var id = "s" + snap.ts;
      return retry(function(){ return db.doc("snaps/" + id).set(snap); })
        .then(function(){ return retry(function(){ return db.doc("pend/atual").set({ map: pend, ts: snap.ts }); }); })
        .then(function(){ return retry(function(){ return db.doc("resumos/" + id).set(resumoOf(snap)); }); })
        .then(function(){ return retry(function(){ return db.doc("estado/atual").set({ snapId: id, ts: snap.ts, arquivo: snap.arquivo }); }); })
        .then(function(){
          return db.collection("snaps").orderBy("ts","desc").limit(40).get().then(function(q){
            return q.docs.slice(12).reduce(function(p, d){ return p.then(function(){ return db.doc("snaps/" + d.id).delete(); }); }, Promise.resolve());
          }).catch(function(){});
        }).then(function(){ return id; });
    },
    watch: function(cb){ return db.doc("estado/atual").onSnapshot(function(s){ cb(s.exists ? s.data().snapId : null); }, function(){}); }
  };
}
function localStore(){
  function open(){ return new Promise(function(res, rej){ try { var rq = indexedDB.open("preinducao-gerencial", 1); rq.onupgradeneeded = function(){ rq.result.createObjectStore("s"); }; rq.onsuccess = function(){ res(rq.result); }; rq.onerror = function(){ rej(rq.error); }; } catch(e){ rej(e); } }); }
  function get(){ return open().then(function(db){ return new Promise(function(res){ try { var rq = db.transaction("s").objectStore("s").get("estado"); rq.onsuccess = function(){ res(rq.result || null); }; rq.onerror = function(){ res(null); }; } catch(e){ res(null); } }); }).catch(function(){ return null; }); }
  return { kind: "local",
    load: function(){ return get().then(function(v){ return v || { snap: null, resumos: [], pend: {}, id: null }; }); },
    save: function(snap, pend){
      return get().then(function(v){
        var resumos = ((v && v.resumos) || []).concat([resumoOf(snap)]).slice(-60), id = "s" + snap.ts;
        return open().then(function(db){ return new Promise(function(res, rej){ var tx = db.transaction("s","readwrite"); tx.objectStore("s").put({ snap: snap, pend: pend, resumos: resumos, id: id }, "estado"); tx.oncomplete = function(){ res(id); }; tx.onerror = function(){ rej(tx.error); }; }); });
      });
    },
    watch: function(){}
  };
}
function initStore(){
  var c = window.claude;
  if (!c || typeof c.use !== "function") return Promise.resolve(localStore());
  return c.use("db").then(function(db){ return db ? cloudStore(db) : localStore(); }).catch(function(){ return localStore(); });
}

/* ================= carregar planilha ================= */
function overlay(msg){ if (msg) { $("overlayMsg").textContent = msg; $("overlay").classList.add("on"); } else $("overlay").classList.remove("on"); }
function fail(msg){ overlay(null); if (S.snap) window.alert(msg); else { $("emptyMsg").className = "msg err"; $("emptyMsg").textContent = msg; } }
function handleFile(file){
  if (!file) return;
  if (!store) { fail("O painel ainda está carregando. Tente de novo em alguns segundos."); return; }
  if (typeof XLSX === "undefined") { fail("A biblioteca que lê o Excel não carregou. Confira a conexão e recarregue a página."); return; }
  overlay("Lendo “" + file.name + "”. Pode levar até 1 minuto…");
  var rd = new FileReader();
  rd.onerror = function(){ fail("Não foi possível ler o arquivo."); };
  rd.onload = function(e){
    setTimeout(function(){
      var res;
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", sheets: ["Caminho Crítico","Base_evolutivo","Regularidade","1328 AMOS"], sheetRows: 60000, dense: true, cellDates: false, cellFormula: false, cellHTML: false, cellText: false, cellStyles: false });
        if (!findSheet(wb, "Caminho Crítico")) throw new Error("Este arquivo não tem a aba Caminho Crítico. Use a planilha de pré-indução da semana.");
        overlay("Calculando indicadores…");
        var rows = readCC(aoaOf(wb, "Caminho Crítico"));
        var be = readBE(aoaOf(wb, "Base_evolutivo")), reg = readReg(aoaOf(wb, "Regularidade")), ord = read1328(aoaOf(wb, "1328 AMOS"));
        wb = null;
        res = build(rows, be, reg, ord, S.pend || {}, file.name);
      } catch(err) { fail(err.message || String(err)); return; }
      overlay("Salvando…");
      store.save(res.snap, res.pend).then(function(id){
        S.resumos = S.resumos.filter(function(x){ return x.ts !== res.snap.ts; }).concat([resumoOf(res.snap)]);
        S.snap = res.snap; S.pend = res.pend; S.snapId = id; overlay(null); render();
      }).catch(function(e){
        var c = e && e.code;
        fail(c === "invalid_argument" || c === "not_granted" ? "Você pode ver o painel, mas não tem permissão para atualizá-lo. Peça ao dono acesso de colaborador ou editor." : "A atualização não foi salva (" + ((e && (e.message || e.code)) || "erro") + "). Tente de novo.");
      });
    }, 60);
  };
  rd.readAsArrayBuffer(file);
}

/* ================= desenho ================= */
function deltaHtml(cur, prev, higherBetter, unit, digits){
  if (cur == null || prev == null) return '<div class="delta eq">Sem comparação ainda</div>';
  var d = cur - prev, eps = digits ? 0.05 : 0.5;
  if (Math.abs(d) < eps) return '<div class="delta eq">Estável em relação à atualização anterior</div>';
  var good = higherBetter ? d > 0 : d < 0;
  var val = digits ? f1(Math.abs(d)) : fi(Math.round(Math.abs(d)));
  return '<div class="delta ' + (good ? "up" : "down") + '">' + (d > 0 ? "▲ " : "▼ ") + val + (unit || "") + " vs. atualização anterior</div>";
}
function cls(v, good, warn, higherBetter){ if (v == null) return ""; return higherBetter ? (v >= good ? "good" : v >= warn ? "warn" : "bad") : (v <= good ? "good" : v <= warn ? "warn" : "bad"); }
function renderKpis(){
  var k = S.snap.kpi, prev = S.resumos.length > 1 ? S.resumos[S.resumos.length-2] : null, cur = resumoOf(S.snap);
  var t60 = k.t60.n ? pct(k.t60.ok, k.t60.n) : null, back = k.back.base ? pct(k.back.n, k.back.base) : null, fora = k.fora.base ? pct(k.fora.n, k.fora.base) : null;
  var cards = [
    { h: "Disponibilidade na indução", v: k.dispInd.v != null ? f1(k.dispInd.v) + "%" : "–", c: cls(k.dispInd.v, 98, 95, true), meta: "Meta: 98% ou mais",
      ctx: k.dispInd.n ? "Média das " + pl(k.dispInd.n, "indução", "induções") + " dos últimos 90 dias; " + k.dispInd.n98 + " atingiram 98%." : "Sem induções nos últimos 90 dias na Base_evolutivo.",
      dl: k.dispInd.prev != null && k.dispInd.v != null ? deltaHtml(k.dispInd.v, k.dispInd.prev, true, " p.p.", 1).replace("vs. atualização anterior", "vs. 90 dias anteriores") : '<div class="delta eq">Sem comparação ainda</div>' },
    { h: "Aderência ao T-60", v: t60 != null ? Math.round(t60) + "%" : "–", c: cls(t60, 100, 80, true), meta: "Meta: 100%",
      ctx: k.t60.n ? k.t60.ok + " de " + pl(k.t60.n, "aeronave", "aeronaves") + " a até 60 dias da indução estão 100% atendidas." : "Nenhuma indução nos próximos 60 dias.", dl: deltaHtml(cur.t60, prev && prev.t60, true, " p.p.", 1) },
    { h: "Linhas sem atendimento até T-60", v: fi(k.nao60), c: cls(k.nao60, 0, 50, false), meta: "Meta: zero",
      ctx: "Inclui aeronaves em execução e as que entram nos próximos 60 dias.", dl: deltaHtml(cur.nao60, prev && prev.nao60, false, "", 0) },
    { h: "Pedidos em backorder", v: back != null ? f1(back) + "%" : "–", c: cls(back, 5, 15, false), meta: "Referência: até 5%",
      ctx: fi(k.back.n) + " de " + fi(k.back.base) + " linhas com pedido estão atrasadas ou sem previsão.", dl: deltaHtml(cur.back, prev && prev.back, false, " p.p.", 1) },
    { h: "Demanda fora do marco", v: fora != null ? f1(fora) + "%" : "–", c: cls(fora, 15, 25, false), meta: "Referência: até 15%",
      ctx: fi(k.fora.n) + " linhas entraram depois do T-180. Depende do Planejamento de Manutenção.", dl: deltaHtml(cur.fora, prev && prev.fora, false, " p.p.", 1) }
  ];
  $("kcards").innerHTML = cards.map(function(c){ return '<div class="kcard ' + c.c + '"><h4>' + c.h + '</h4><div class="v">' + c.v + '</div><div class="meta">' + c.meta + "</div>" + c.dl + '<div class="ctx">' + c.ctx + "</div></div>"; }).join("");
}
function renderActions(){
  var acts = S.snap.actions, areas = {};
  acts.forEach(function(a){ areas[a.area] = (areas[a.area] || 0) + 1; });
  if (S.actArea && !areas[S.actArea]) S.actArea = "";
  $("actFilter").innerHTML = '<button type="button" class="chip" data-area="" aria-pressed="' + !S.actArea + '">Todas (' + acts.length + ")</button>" +
    Object.keys(areas).map(function(k){ return '<button type="button" class="chip" data-area="' + esc(k) + '" aria-pressed="' + (S.actArea === k) + '">' + esc(k) + " (" + areas[k] + ")</button>"; }).join("");
  var list = acts.filter(function(a){ return !S.actArea || a.area === S.actArea; });
  $("actions").innerHTML = list.length ? list.map(function(a){ return '<div class="act ' + a.sev + '"><span class="sev" title="' + (a.sev === "alta" ? "Prioridade alta" : "Prioridade média") + '"></span><div><b>' + esc(a.t) + '</b><span class="d">' + esc(a.d) + '</span></div><span class="area">' + esc(a.area) + "</span></div>"; }).join("")
    : '<p class="empty">Nenhum ponto de ação: as induções dos próximos 60 dias estão atendidas e não há pedidos críticos.</p>';
}
function renderTimeline(){
  var t0 = today(), MAX = 180, x = function(d){ return Math.max(0, Math.min(MAX, d)) / MAX * 100; };
  var evs = S.snap.events.map(function(e){ var c = {}; Object.keys(e).forEach(function(k){ c[k] = e[k]; }); c.days = e.prog != null ? e.prog - t0 : null; return c; })
    .filter(function(e){ return e.days == null || e.days >= -45; });
  var marks = [[0,"Hoje"],[60,"T-60"],[120,"T-120"],[180,"T-180"]];
  var scale = marks.map(function(m){ return '<span style="left:' + x(m[0]) + '%"><b>' + m[1] + "</b>" + fd(t0 + m[0], true) + "</span>"; }).join("");
  var guides = '<span class="zone" style="width:' + x(60) + '%"></span>' + [7,45,60,120,180].map(function(d){ return '<i style="left:' + x(d) + '%"></i>'; }).join("");
  $("timeline").innerHTML = evs.length ? '<div class="tl-axis"><span>Aeronave</span><div class="tl-scale">' + scale + '</div><span style="text-align:right">Pend.</span></div><div class="tl-body"><div class="tl-guides">' + guides + "</div>" +
    evs.map(function(e){
      var at = Math.round(pct(e.d + e.a, e.t)), c = e.n === 0 ? "ok" : (e.days != null && e.days <= 60 ? "bad" : "mid");
      var when = e.days == null ? "sem data" : e.days < 0 ? "em execução" : e.days === 0 ? "hoje" : "em " + pl(e.days, "dia", "dias");
      return '<div class="tl-row" style="cursor:default" title="' + esc(e.acft + ", " + e.ativ + ", " + e.t + " linhas, " + e.n + " sem atendimento") + '"><span class="tl-name">' + esc(e.acft) + "<small>" + fd(e.prog, true) + ", " + when + '</small></span><span class="tl-track"><span class="tl-pill ' + c + (e.days != null && e.days < 0 ? " past" : "") + '" style="left:' + x(e.days == null ? MAX : e.days) + '%">' + at + '%</span></span><span class="tl-pct" style="color:' + (e.n ? "var(--nao)" : "var(--disp)") + '">' + (e.n ? fi(e.n) : "✓") + "</span></div>";
    }).join("") + "</div>" : '<p class="empty">Nenhuma indução programada.</p>';
}
function barChart(id, items, opts){
  var W = 600, H = 220, L = 34, R = 8, T = 18, B = 34, svg = $(id);
  if (!items.length) { svg.innerHTML = '<text x="300" y="110" text-anchor="middle">' + opts.empty + "</text>"; return; }
  var mn = opts.min, mx = opts.max, n = items.length, bw = (W - L - R) / n;
  var y = function(v){ return T + (1 - (Math.max(mn, Math.min(mx, v)) - mn) / (mx - mn)) * (H - T - B); };
  var out = "";
  opts.ticks.forEach(function(t){ out += '<line class="grid" x1="' + L + '" x2="' + (W - R) + '" y1="' + y(t) + '" y2="' + y(t) + '"/><text x="' + (L - 6) + '" y="' + (y(t) + 4) + '" text-anchor="end">' + t + "%</text>"; });
  items.forEach(function(it, i){
    var x0 = L + i * bw + bw * 0.18, w = bw * 0.64, yv = y(it.v), color = it.v >= opts.meta ? "var(--disp)" : it.v >= opts.warn ? "var(--atende)" : "var(--nao)";
    out += '<rect x="' + x0.toFixed(1) + '" y="' + yv.toFixed(1) + '" width="' + w.toFixed(1) + '" height="' + (H - B - yv).toFixed(1) + '" rx="2" fill="' + color + '"><title>' + esc(it.title) + "</title></rect>";
    if (n <= 18) out += '<text class="lbl" x="' + (x0 + w/2).toFixed(1) + '" y="' + (yv - 4).toFixed(1) + '" text-anchor="middle">' + Math.round(it.v) + "</text>";
    out += '<text x="' + (x0 + w/2).toFixed(1) + '" y="' + (H - B + 15) + '" text-anchor="middle">' + it.lbl + "</text>";
    out += '<text x="' + (x0 + w/2).toFixed(1) + '" y="' + (H - B + 28) + '" text-anchor="middle">' + it.sub + "</text>";
  });
  out += '<line class="meta-line" x1="' + L + '" x2="' + (W - R) + '" y1="' + y(opts.meta) + '" y2="' + y(opts.meta) + '"/>';
  svg.innerHTML = out;
}
function renderCharts(){
  barChart("chMes", S.snap.histMes.map(function(m, i){ return { v: m.v, lbl: MES[parseInt(m.m.split("-")[1],10)-1], sub: (i === 0 || m.m.slice(5) === "01") ? m.m.slice(0,4) : "", title: monthLbl(m.m) + ": " + f1(m.v) + "% em " + pl(m.n, "indução", "induções") + ", " + m.n98 + " com 98% ou mais" }; }),
    { min: 80, max: 100, ticks: [80, 85, 90, 95, 100], meta: 98, warn: 95, empty: "A aba Base_evolutivo não foi encontrada ou não tem induções passadas." });
  barChart("chT60", S.snap.histT60.map(function(q){ var v = pct(q.ok, q.n); return { v: v, lbl: q.q.replace("T", " T"), sub: q.ok + " de " + q.n, title: q.q + ": " + q.ok + " de " + pl(q.n, "aeronave", "aeronaves") + " 100% atendidas no T-60" }; }),
    { min: 0, max: 100, ticks: [0, 25, 50, 75, 100], meta: 100, warn: 80, empty: "Sem histórico de T-60 na Base_evolutivo." });
}
function tableHtml(head, rows){ return '<div class="scroll"><table><thead><tr>' + head.map(function(h){ return "<th" + (h[1] ? ' class="num"' : "") + ">" + h[0] + "</th>"; }).join("") + "</tr></thead><tbody>" + rows.join("") + "</tbody></table></div>"; }
function renderTables(){
  var s = S.snap, max = s.resp.reduce(function(m,x){ return Math.max(m, x.n); }, 0);
  $("resp").innerHTML = s.resp.length ? '<div class="hbars">' + s.resp.map(function(r){ return '<div class="hb"><span>' + esc(r.r) + '</span><span class="track"><span style="width:' + pct(r.n, max) + '%;background:var(--nao)"></span></span><b>' + fi(r.n) + '</b></div>' + (r.idade != null ? '<div class="who" style="margin:-5px 0 4px 0">Pendentes há ' + pl(Math.round(r.idade), "dia", "dias") + ' em média</div>' : ""); }).join("") + "</div>" + (s.resp.length && s.resp[0].idade == null ? '<p class="who" style="margin-top:10px">O tempo médio de pendência por área aparece a partir da segunda atualização.</p>' : "") : '<p class="empty">Nenhuma linha sem atendimento.</p>';
  $("orders").innerHTML = s.orders.length ? tableHtml([["Pedido"],["Tipo"],["Linhas",1],["Acft",1],["Situação"]], s.orders.map(function(o){ return "<tr><td>" + esc(o.o) + (o.ext ? '<div class="who">' + esc(o.ext) + "</div>" : "") + "</td><td>" + esc(o.tipo) + '</td><td class="num">' + fi(o.linhas) + '</td><td class="num">' + o.nAcft + "</td><td" + (o.sem ? "" : ' class="late"') + ">" + (o.sem ? "Sem previsão" : "Atraso de " + pl(o.atraso, "dia", "dias")) + "</td></tr>"; }))
    : '<p class="empty">Nenhum pedido em backorder.</p>';
  $("pns").innerHTML = s.pns.length ? tableHtml([["PN"],["Descrição"],["Linhas",1],["Acft",1],["Perfil"],["Cons./mês",1]], s.pns.map(function(p){ return "<tr><td>" + esc(p.pn) + '</td><td class="desc">' + esc(p.desc) + '</td><td class="num">' + fi(p.linhas) + '</td><td class="num">' + p.nAcft + '</td><td><span class="reg ' + esc(p.reg) + '">' + esc(p.reg || "Sem classificação") + '</span></td><td class="num">' + (p.cons != null ? fi(Math.round(p.cons)) : "") + "</td></tr>"; }))
    : '<p class="empty">Nenhum item crítico.</p>';
  var vmax = s.vend.reduce(function(m,x){ return Math.max(m, x.n); }, 0);
  $("vend").innerHTML = s.vend.length ? s.vend.map(function(v){ return '<div class="hb" style="grid-template-columns:minmax(120px,230px) minmax(0,1fr) 50px"><span title="' + esc(v.name) + '">' + esc(v.name) + '</span><span class="track"><span style="width:' + pct(v.n, vmax) + '%;background:var(--atende)"></span></span><b>' + fi(v.n) + "</b></div>"; }).join("") +
    '<p class="who" style="margin-top:10px">' + fi(s.kpi.venc.n) + " de " + fi(s.kpi.venc.base) + " itens abertos desses pedidos estão vencidos.</p>"
    : '<p class="empty">' + (s.kpi.venc.base ? "Nenhum item vencido." : "A aba 1328 AMOS não foi encontrada na planilha.") + "</p>";
  $("inducoes").innerHTML = s.inducoes.length ? tableHtml([["Aeronave"],["Pacote"],["Entrada"],["Disponível na véspera",1],["Situação"]], s.inducoes.map(function(x){ var v = x.disp * 100; return "<tr><td>" + esc(x.acft) + "</td><td>" + esc(x.wp) + "</td><td>" + fd(x.ini) + '</td><td class="num">' + f1(v) + '%</td><td><span class="tag ' + (v >= 99.95 ? "disp" : v >= 98 ? "atende" : "nao") + '">' + (v >= 99.95 ? "100%" : v >= 98 ? "Dentro da meta" : "Abaixo da meta") + "</span></td></tr>"; }))
    : '<p class="empty">Sem induções registradas na Base_evolutivo.</p>';
}
function render(){
  var has = !!S.snap;
  $("empty").classList.toggle("hidden", has); $("dash").classList.toggle("hidden", !has);
  if (!has) return;
  var s = S.snap;
  $("subInfo").textContent = "Atualizado em " + new Date(s.ts).toLocaleString("pt-BR") + " com " + s.arquivo;
  renderKpis(); renderActions(); renderTimeline(); renderCharts(); renderTables();
  $("note").textContent = "Base: " + fi(s.linhas) + " linhas do Caminho Crítico e " + fi(s.beN) + " registros da Base_evolutivo. " + pl(S.resumos.length, "atualização salva", "atualizações salvas") + " no histórico do painel.";
}

/* ================= início e eventos ================= */
function loadAll(){
  overlay("Carregando painel…");
  return store.load().then(function(d){ S.snap = d.snap; S.pend = d.pend || {}; S.resumos = d.resumos || []; S.snapId = d.id; $("newBanner").classList.add("hidden"); overlay(null); render(); })
    .catch(function(e){ overlay(null); fail("Não foi possível carregar os dados salvos (" + ((e && (e.message || e.code)) || "erro") + "). Recarregue a página."); });
}
initStore().then(function(s){
  store = s;
  var b = $("storeBadge");
  if (s.kind === "cloud") { b.textContent = "Dados compartilhados com a equipe"; b.classList.add("cloud"); }
  else { b.textContent = "Dados só neste navegador"; $("localBanner").classList.remove("hidden"); }
  loadAll().then(function(){ var first = true; s.watch(function(id){ if (first) { first = false; return; } if (id && id !== S.snapId) $("newBanner").classList.remove("hidden"); }); });
});
$("loadBtn").addEventListener("click", function(){ $("fileInput").click(); });
$("empty").addEventListener("click", function(){ $("fileInput").click(); });
$("empty").addEventListener("keydown", function(e){ if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("fileInput").click(); } });
$("fileInput").addEventListener("change", function(e){ handleFile(e.target.files[0]); e.target.value = ""; });
["dragenter","dragover"].forEach(function(t){ document.addEventListener(t, function(e){ e.preventDefault(); $("empty").classList.add("over"); }); });
["dragleave","drop"].forEach(function(t){ document.addEventListener(t, function(e){ e.preventDefault(); $("empty").classList.remove("over"); }); });
document.addEventListener("drop", function(e){ var f = e.dataTransfer && e.dataTransfer.files[0]; if (f) handleFile(f); });
$("reloadBtn").addEventListener("click", loadAll);
$("actFilter").addEventListener("click", function(e){ var c = e.target.closest("[data-area]"); if (c) { S.actArea = c.getAttribute("data-area"); renderActions(); } });
$("themeBtn").addEventListener("click", function(){
  var cur = document.documentElement.getAttribute("data-theme");
  var dark = cur ? cur === "dark" : window.matchMedia("(prefers-color-scheme: dark)").matches;
  var next = dark ? "light" : "dark"; document.documentElement.setAttribute("data-theme", next);
  try { localStorage.setItem("preinducao-tema", next); } catch(e){}
});
try { var th = localStorage.getItem("preinducao-tema"); if (th) document.documentElement.setAttribute("data-theme", th); } catch(e){}
window.__pg = { S: S, build: build };
})();

</script>
</body>
</html>
