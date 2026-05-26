// ============================================================
//  ADEGA POSTIN — GESTÃO FINANCEIRA v3.0
//  Extensões > Apps Script > Cole e salve > Recarregue a planilha
//  Depois: menu 🍺 Postin > Configurar Planilha
// ============================================================

const NOME_ADEGA         = "Adega Postin";
const SHEET_LANCAMENTOS  = "📒 Lançamentos";
const SHEET_CATEGORIAS   = "🏷️ Categorias";
const SHEET_CONTAS_PAGAR = "📅 Contas a Pagar";
const SHEET_RELATORIOS   = "📊 Relatórios";
const SHEET_DASHBOARD    = "🏠 Dashboard";
const SHEET_ANUAL        = "📆 Resumo Anual";

const COR_RECEITA  = "#e6f4ea";
const COR_DESPESA  = "#fce8e6";
const COR_HEADER   = "#1a3a5c";
const COR_HEADER_T = "#ffffff";
const COR_ACCENT   = "#c8a951";

// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🍺 Postin Financeiro")
    .addItem("🚀 Configurar Planilha",            "configurarPlanilha")
    .addSeparator()
    .addItem("☀️ Fechamento do Dia (Receita)",    "abrirFechamentoDia")
    .addItem("➕ Lançar Despesa",                 "abrirLancamentoDespesa")
    .addItem("📅 Nova Conta a Pagar",             "abrirContaPagar")
    .addSeparator()
    .addItem("🔄 Atualizar Dashboard",            "atualizarDashboard")
    .addItem("🔄 Atualizar Relatórios",           "atualizarRelatorios")
    .addItem("🔄 Atualizar Resumo Anual",         "atualizarResumoAnual")
    .addSeparator()
    .addItem("📲 Gerar Resumo Semanal (WhatsApp)","gerarResumoWhatsApp")
    .addItem("⚠️ Verificar Vencimentos",          "alertarVencimentos")
    .addToUi();
}

// ============================================================
//  CONFIGURAÇÃO INICIAL
// ============================================================
function configurarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert("Configurar Planilha",
    "Isso vai criar/recriar todas as abas.\nDados existentes serão apagados.\n\nContinuar?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  _criarAbaCategorias(ss);
  _criarAbaLancamentos(ss);
  _criarAbaContasPagar(ss);
  _criarAbaRelatorios(ss);
  _criarAbaDashboard(ss);
  _criarAbaAnual(ss);

  const ordem = [SHEET_DASHBOARD, SHEET_ANUAL, SHEET_LANCAMENTOS, SHEET_CATEGORIAS, SHEET_CONTAS_PAGAR, SHEET_RELATORIOS];
  ordem.forEach((nome, i) => {
    const sh = ss.getSheetByName(nome);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });

  ss.getSheetByName(SHEET_DASHBOARD).activate();
  ui.alert("✅ Adega Postin configurada!\n\nUse o menu 🍺 Postin Financeiro para começar.");
}

// ============================================================
//  ABA: CATEGORIAS
// ============================================================
function _criarAbaCategorias(ss) {
  let sh = ss.getSheetByName(SHEET_CATEGORIAS);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_CATEGORIAS);

  sh.getRange(1,1,1,3).setValues([["Categoria","Tipo","Descrição"]]);
  _estilizarHeader(sh, 1, 3);

  const cats = [
    ["Venda Balcão",             "Receita", "Vendas presenciais no balcão"],
    ["Delivery",                 "Receita", "Vendas por entrega / motoboy"],
    ["Venda Atacado",            "Receita", "Vendas em quantidade para eventos/bares"],
    ["Outras Receitas",          "Receita", "Receitas eventuais diversas"],
    ["Mercadoria",               "Despesa", "Compra de bebidas e reposição de estoque"],
    ["Aluguel",                  "Despesa", "Aluguel do ponto comercial"],
    ["Energia Elétrica",         "Despesa", "Conta de luz"],
    ["Água",                     "Despesa", "Conta de água"],
    ["Salários",                 "Despesa", "Pagamento mensal da equipe"],
    ["Prêmio de Funcionário",    "Despesa", "Bonificações e prêmios por desempenho"],
    ["Lanche de Funcionário",    "Despesa", "Alimentação da equipe durante o expediente"],
    ["Simples Nacional",         "Despesa", "Imposto mensal — DAS Simples Nacional"],
    ["INSS",                     "Despesa", "Contribuição previdenciária patronal"],
    ["FGTS",                     "Despesa", "Fundo de Garantia por Tempo de Serviço"],
    ["Material de Limpeza",      "Despesa", "Produtos de higiene e limpeza"],
    ["Material de Escritório",   "Despesa", "Papelaria, canetas, papel, impressões"],
    ["Frete / Entrega",          "Despesa", "Custo de frete de mercadoria recebida"],
    ["Manutenção e Reparos",     "Despesa", "Freezer, geladeiras, equipamentos"],
    ["Taxas Bancárias",          "Despesa", "Tarifas da maquininha, IOF, tarifas bancárias"],
    ["Marketing",                "Despesa", "Redes sociais, apps de delivery, promoções"],
    ["Telefone / Internet",      "Despesa", "Plano de dados, internet fixa"],
    ["Combustível / Transporte", "Despesa", "Combustível, transporte operacional"],
    ["Embalagens",               "Despesa", "Sacolas, caixas, isopor, etiquetas"],
    ["Quebra e Avaria",          "Despesa", "Produto quebrado, vencido ou devolvido"],
    ["Outras Despesas",          "Despesa", "Despesas eventuais não categorizadas"],
  ];

  sh.getRange(2,1,cats.length,3).setValues(cats);
  cats.forEach((c,i) => {
    sh.getRange(i+2,1,1,3).setBackground(c[1]==="Receita" ? COR_RECEITA : COR_DESPESA);
  });

  const tipoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Receita","Despesa"],true).build();
  sh.getRange("B2:B500").setDataValidation(tipoRule);
  [220,100,300].forEach((w,i) => sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1);
  sh.getRange(cats.length+3,1)
    .setValue("💡 Adicione ou renomeie categorias nesta aba. O sistema se atualiza automaticamente.")
    .setFontColor("#888").setFontSize(10).setFontStyle("italic");
}

// ============================================================
//  ABA: LANÇAMENTOS
// ============================================================
function _criarAbaLancamentos(ss) {
  let sh = ss.getSheetByName(SHEET_LANCAMENTOS);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_LANCAMENTOS);

  const headers = ["Data","Tipo","Categoria","Descrição","Débito (R$)","Crédito (R$)","PIX (R$)","Dinheiro (R$)","Delivery (R$)","Total (R$)","Observação"];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  _estilizarHeader(sh, 1, headers.length);
  [100,80,170,220,105,105,105,115,115,115,180].forEach((w,i) => sh.setColumnWidth(i+1,w));

  const tipoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Receita","Despesa"],true).build();
  sh.getRange("B2:B10000").setDataValidation(tipoRule);

  const shCat = ss.getSheetByName(SHEET_CATEGORIAS);
  if (shCat) {
    const catRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(shCat.getRange("A2:A200"),true).build();
    sh.getRange("C2:C10000").setDataValidation(catRule);
  }
  ["E","F","G","H","I","J"].forEach(col => {
    sh.getRange(col+"2:"+col+"10000").setNumberFormat("R$ #,##0.00");
  });
  sh.getRange("A2:A10000").setNumberFormat("dd/mm/yyyy");
  sh.setFrozenRows(1);
}

// ============================================================
//  ABA: CONTAS A PAGAR
// ============================================================
function _criarAbaContasPagar(ss) {
  let sh = ss.getSheetByName(SHEET_CONTAS_PAGAR);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_CONTAS_PAGAR);

  const headers = ["Descrição","Fornecedor/Credor","Valor (R$)","Vencimento","Status","Data Pgto","Forma Pgto","Observação"];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  _estilizarHeader(sh, 1, headers.length);
  [230,190,120,110,100,110,130,190].forEach((w,i) => sh.setColumnWidth(i+1,w));

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pendente","Pago","Vencido","Cancelado"],true).build();
  sh.getRange("E2:E10000").setDataValidation(statusRule);

  const pgtoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["PIX","Boleto","Transferência","Cartão Débito","Cartão Crédito","Dinheiro","Outro"],true).build();
  sh.getRange("G2:G10000").setDataValidation(pgtoRule);

  sh.getRange("C2:C10000").setNumberFormat("R$ #,##0.00");
  sh.getRange("D2:D10000").setNumberFormat("dd/mm/yyyy");
  sh.getRange("F2:F10000").setNumberFormat("dd/mm/yyyy");
  sh.setFrozenRows(1);

  const range = sh.getRange("A2:H10000");
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$E2="Pago"')
      .setBackground("#e6f4ea").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$E2="Vencido"')
      .setBackground("#fce8e6").setFontColor("#a50000").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$E2="Pendente"')
      .setBackground("#fff9c4").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$E2="Cancelado"')
      .setBackground("#f1f3f4").setFontColor("#aaa").setRanges([range]).build(),
  ]);
}

function _criarAbaRelatorios(ss) {
  let sh = ss.getSheetByName(SHEET_RELATORIOS);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_RELATORIOS);
  sh.getRange("A1")
    .setValue("📊 Relatórios — use o menu 🍺 Postin > Atualizar Relatórios para gerar.")
    .setFontColor("#888").setFontSize(12).setFontStyle("italic");
}

function _criarAbaAnual(ss) {
  let sh = ss.getSheetByName(SHEET_ANUAL);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_ANUAL);
  sh.getRange("A1")
    .setValue("📆 Resumo Anual — use o menu 🍺 Postin > Atualizar Resumo Anual para gerar.")
    .setFontColor("#888").setFontSize(12).setFontStyle("italic");
}

function _criarAbaDashboard(ss) {
  let sh = ss.getSheetByName(SHEET_DASHBOARD);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_DASHBOARD);
  atualizarDashboard();
}

// ============================================================
//  FORMULÁRIO — FECHAMENTO DO DIA
// ============================================================
function abrirFechamentoDia() {
  const html = HtmlService.createHtmlOutput(_htmlFechamento())
    .setWidth(460).setHeight(590).setTitle("☀️ Fechamento do Dia");
  SpreadsheetApp.getUi().showModalDialog(html, "☀️ Fechamento do Dia — Adega Postin");
}

function _htmlFechamento() {
  const ontem = _dataOffset(-1);
  return '<!DOCTYPE html><html><head>' +
'<meta name="viewport" content="width=device-width,initial-scale=1">' +
'<style>' +
'@import url(\'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600&display=swap\');' +
'*{box-sizing:border-box;margin:0;padding:0;font-family:\'DM Sans\',Arial,sans-serif;}' +
'body{background:#f0f4f0;padding:18px;}' +
'.header{background:linear-gradient(135deg,#1a3a5c,#2d6a4f);border-radius:12px;padding:16px 18px;margin-bottom:18px;color:#fff;}' +
'.header h2{font-size:17px;font-weight:600;margin-bottom:2px;}' +
'.header p{font-size:11px;opacity:.75;}' +
'label{display:block;font-size:11px;color:#555;margin:12px 0 4px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;}' +
'input{width:100%;padding:11px 13px;border:1.5px solid #dde3dd;border-radius:8px;font-size:15px;background:#fff;transition:border .2s;}' +
'input:focus{outline:none;border-color:#2d6a4f;}' +
'.grid2{display:grid;grid-template-columns:1fr 1fr;gap:10px;}' +
'.total-box{background:linear-gradient(135deg,#1a3a5c,#c8a951);border-radius:10px;padding:14px 18px;text-align:center;margin-top:16px;color:#fff;}' +
'.total-box small{font-size:11px;opacity:.8;display:block;margin-bottom:4px;}' +
'.total-box span{font-size:26px;font-weight:700;}' +
'.btn{width:100%;margin-top:14px;padding:13px;background:linear-gradient(135deg,#1a3a5c,#2d6a4f);color:#fff;border:none;border-radius:9px;font-size:15px;font-weight:600;cursor:pointer;}' +
'.btn:hover{opacity:.9;}' +
'.msg{text-align:center;font-size:12px;margin-top:8px;min-height:18px;}' +
'.msg.ok{color:#2d6a4f;font-weight:600;}.msg.err{color:#c5221f;}' +
'</style></head><body>' +
'<div class="header"><h2>☀️ Fechamento do Dia</h2><p>Lançamento de receita por forma de pagamento</p></div>' +
'<label>📅 Data do fechamento</label>' +
'<input type="date" id="data" value="' + ontem + '" oninput="calcTotal()">' +
'<div class="grid2">' +
'<div><label>💳 Débito (R$)</label><input type="number" id="debito" value="0" step="0.01" min="0" oninput="calcTotal()"></div>' +
'<div><label>💳 Crédito (R$)</label><input type="number" id="credito" value="0" step="0.01" min="0" oninput="calcTotal()"></div>' +
'</div>' +
'<div class="grid2">' +
'<div><label>📱 PIX (R$)</label><input type="number" id="pix" value="0" step="0.01" min="0" oninput="calcTotal()"></div>' +
'<div><label>💵 Dinheiro (R$)</label><input type="number" id="dinheiro" value="0" step="0.01" min="0" oninput="calcTotal()"></div>' +
'</div>' +
'<label>🛵 Delivery (R$)</label>' +
'<input type="number" id="delivery" value="0" step="0.01" min="0" oninput="calcTotal()">' +
'<label>Observação</label>' +
'<input type="text" id="obs" placeholder="Ex: movimento fraco, feriado...">' +
'<div class="total-box"><small>TOTAL DO DIA</small><span id="total">R$ 0,00</span></div>' +
'<button class="btn" onclick="salvar()">💾 Salvar Fechamento</button>' +
'<div class="msg" id="msg"></div>' +
'<script>' +
'function calcTotal(){' +
'var v=["debito","credito","pix","dinheiro","delivery"].reduce(function(s,id){return s+(parseFloat(document.getElementById(id).value)||0);},0);' +
'document.getElementById("total").textContent="R$ "+v.toLocaleString("pt-BR",{minimumFractionDigits:2});return v;}' +
'function salvar(){' +
'var data=document.getElementById("data").value;' +
'var debito=parseFloat(document.getElementById("debito").value)||0;' +
'var credito=parseFloat(document.getElementById("credito").value)||0;' +
'var pix=parseFloat(document.getElementById("pix").value)||0;' +
'var dinheiro=parseFloat(document.getElementById("dinheiro").value)||0;' +
'var delivery=parseFloat(document.getElementById("delivery").value)||0;' +
'var obs=document.getElementById("obs").value.trim();' +
'var total=debito+credito+pix+dinheiro+delivery;' +
'var msg=document.getElementById("msg");' +
'if(!data){msg.className="msg err";msg.textContent="Informe a data.";return;}' +
'if(total<=0){msg.className="msg err";msg.textContent="Informe ao menos um valor.";return;}' +
'msg.className="msg";msg.textContent="Salvando...";' +
'google.script.run' +
'.withSuccessHandler(function(){msg.className="msg ok";msg.textContent="✅ Fechamento salvo! Total: R$ "+total.toLocaleString("pt-BR",{minimumFractionDigits:2});' +
'["debito","credito","pix","dinheiro","delivery"].forEach(function(id){document.getElementById(id).value="0";});' +
'document.getElementById("obs").value="";calcTotal();})' +
'.withFailureHandler(function(e){msg.className="msg err";msg.textContent="Erro: "+e.message;})' +
'.salvarFechamentoDia({data:data,debito:debito,credito:credito,pix:pix,dinheiro:dinheiro,delivery:delivery,obs:obs});}' +
'calcTotal();' +
'<\/script></body></html>';
}

function salvarFechamentoDia(d) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(SHEET_LANCAMENTOS);
  if (!sh) throw new Error("Aba Lançamentos não encontrada.");
  const p   = d.data.split("-");
  const data= new Date(+p[0],+p[1]-1,+p[2]);
  const total= d.debito+d.credito+d.pix+d.dinheiro+d.delivery;
  const row = [data,"Receita","Venda Balcão","Fechamento do Dia",
               d.debito,d.credito,d.pix,d.dinheiro,d.delivery,total,d.obs];
  sh.appendRow(row);
  const ul = sh.getLastRow();
  sh.getRange(ul,1,1,row.length).setBackground(COR_RECEITA);
  sh.getRange(ul,1).setNumberFormat("dd/mm/yyyy");
  ["E","F","G","H","I","J"].forEach(function(col,i){sh.getRange(ul,5+i).setNumberFormat("R$ #,##0.00");});
  return {sucesso:true, total:total};
}

// ============================================================
//  FORMULÁRIO — LANÇAR DESPESA
// ============================================================
function abrirLancamentoDespesa() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shCat = ss.getSheetByName(SHEET_CATEGORIAS);
  let opts = "<option value=''>— Selecione —</option>";
  if (shCat) {
    shCat.getRange("A2:B200").getValues()
      .filter(function(r){return r[0]&&r[1]==="Despesa";})
      .forEach(function(r){opts += "<option value='"+r[0]+"'>"+r[0]+"</option>";});
  }
  const html = HtmlService.createHtmlOutput(_htmlDespesa(opts))
    .setWidth(440).setHeight(510).setTitle("➕ Lançar Despesa");
  SpreadsheetApp.getUi().showModalDialog(html, "➕ Lançar Despesa — Adega Postin");
}

function _htmlDespesa(opts) {
  const hoje = _hoje();
  return '<!DOCTYPE html><html><head>' +
'<meta name="viewport" content="width=device-width,initial-scale=1">' +
'<style>' +
'@import url(\'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600&display=swap\');' +
'*{box-sizing:border-box;margin:0;padding:0;font-family:\'DM Sans\',Arial,sans-serif;}' +
'body{background:#fdf3e7;padding:18px;}' +
'.header{background:linear-gradient(135deg,#7f3f00,#c8a951);border-radius:12px;padding:16px 18px;margin-bottom:18px;color:#fff;}' +
'.header h2{font-size:17px;font-weight:600;margin-bottom:2px;}.header p{font-size:11px;opacity:.75;}' +
'label{display:block;font-size:11px;color:#555;margin:12px 0 4px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;}' +
'input,select{width:100%;padding:11px 13px;border:1.5px solid #e4ceaa;border-radius:8px;font-size:14px;background:#fff;transition:border .2s;}' +
'input:focus,select:focus{outline:none;border-color:#c8a951;}' +
'.grid2{display:grid;grid-template-columns:1fr 1fr;gap:10px;}' +
'.btn{width:100%;margin-top:16px;padding:13px;background:linear-gradient(135deg,#7f3f00,#c8a951);color:#fff;border:none;border-radius:9px;font-size:15px;font-weight:600;cursor:pointer;}' +
'.btn:hover{opacity:.9;}' +
'.msg{text-align:center;font-size:12px;margin-top:8px;min-height:18px;}' +
'.msg.ok{color:#2d6a4f;font-weight:600;}.msg.err{color:#c5221f;}' +
'</style></head><body>' +
'<div class="header"><h2>➕ Lançar Despesa</h2><p>Registre uma despesa da Adega Postin</p></div>' +
'<label>📅 Data</label><input type="date" id="data" value="' + hoje + '">' +
'<label>Categoria *</label><select id="cat">' + opts + '</select>' +
'<label>Descrição *</label><input type="text" id="desc" placeholder="Ex: Compra de cerveja — Distribuidora X">' +
'<div class="grid2">' +
'<div><label>Valor (R$) *</label><input type="number" id="valor" placeholder="0,00" step="0.01" min="0"></div>' +
'<div><label>Forma de Pgto</label><select id="pgto">' +
'<option>PIX</option><option>Boleto</option><option>Transferência</option>' +
'<option>Cartão Débito</option><option>Cartão Crédito</option><option>Dinheiro</option><option>Outro</option>' +
'</select></div></div>' +
'<label>Observação</label><input type="text" id="obs" placeholder="Opcional">' +
'<button class="btn" onclick="salvar()">💾 Salvar Despesa</button>' +
'<div class="msg" id="msg"></div>' +
'<script>' +
'function salvar(){' +
'var data=document.getElementById("data").value;' +
'var cat=document.getElementById("cat").value;' +
'var desc=document.getElementById("desc").value.trim();' +
'var val=parseFloat(document.getElementById("valor").value);' +
'var pgto=document.getElementById("pgto").value;' +
'var obs=document.getElementById("obs").value.trim();' +
'var msg=document.getElementById("msg");' +
'if(!data||!cat||!desc||isNaN(val)||val<=0){msg.className="msg err";msg.textContent="Preencha todos os campos obrigatórios.";return;}' +
'msg.className="msg";msg.textContent="Salvando...";' +
'google.script.run' +
'.withSuccessHandler(function(){msg.className="msg ok";msg.textContent="✅ Despesa de R$ "+val.toLocaleString("pt-BR",{minimumFractionDigits:2})+" salva!";' +
'document.getElementById("desc").value="";document.getElementById("valor").value="";document.getElementById("obs").value="";})' +
'.withFailureHandler(function(e){msg.className="msg err";msg.textContent="Erro: "+e.message;})' +
'.salvarDespesa({data:data,cat:cat,desc:desc,val:val,pgto:pgto,obs:obs});}' +
'<\/script></body></html>';
}

function salvarDespesa(d) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(SHEET_LANCAMENTOS);
  if (!sh) throw new Error("Aba Lançamentos não encontrada.");
  const p   = d.data.split("-");
  const data= new Date(+p[0],+p[1]-1,+p[2]);
  const row = [data,"Despesa",d.cat,d.desc,0,0,0,0,0,d.val,d.obs];
  sh.appendRow(row);
  const ul = sh.getLastRow();
  sh.getRange(ul,1,1,row.length).setBackground(COR_DESPESA);
  sh.getRange(ul,1).setNumberFormat("dd/mm/yyyy");
  ["E","F","G","H","I","J"].forEach(function(col,i){sh.getRange(ul,5+i).setNumberFormat("R$ #,##0.00");});
  return {sucesso:true};
}

// ============================================================
//  FORMULÁRIO — CONTA A PAGAR
// ============================================================
function abrirContaPagar() {
  const html = HtmlService.createHtmlOutput(_htmlContaPagar())
    .setWidth(440).setHeight(500).setTitle("📅 Nova Conta a Pagar");
  SpreadsheetApp.getUi().showModalDialog(html, "📅 Nova Conta a Pagar — Adega Postin");
}

function _htmlContaPagar() {
  return '<!DOCTYPE html><html><head>' +
'<meta name="viewport" content="width=device-width,initial-scale=1">' +
'<style>' +
'@import url(\'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600&display=swap\');' +
'*{box-sizing:border-box;margin:0;padding:0;font-family:\'DM Sans\',Arial,sans-serif;}' +
'body{background:#eef2ff;padding:18px;}' +
'.header{background:linear-gradient(135deg,#1a3a5c,#3a5fa0);border-radius:12px;padding:16px 18px;margin-bottom:18px;color:#fff;}' +
'.header h2{font-size:17px;font-weight:600;margin-bottom:2px;}.header p{font-size:11px;opacity:.75;}' +
'label{display:block;font-size:11px;color:#555;margin:12px 0 4px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;}' +
'input,select{width:100%;padding:11px 13px;border:1.5px solid #c5cfe8;border-radius:8px;font-size:14px;background:#fff;transition:border .2s;}' +
'input:focus,select:focus{outline:none;border-color:#3a5fa0;}' +
'.grid2{display:grid;grid-template-columns:1fr 1fr;gap:10px;}' +
'.btn{width:100%;margin-top:16px;padding:13px;background:linear-gradient(135deg,#1a3a5c,#3a5fa0);color:#fff;border:none;border-radius:9px;font-size:15px;font-weight:600;cursor:pointer;}' +
'.btn:hover{opacity:.9;}' +
'.msg{text-align:center;font-size:12px;margin-top:8px;min-height:18px;}' +
'.msg.ok{color:#2d6a4f;font-weight:600;}.msg.err{color:#c5221f;}' +
'</style></head><body>' +
'<div class="header"><h2>📅 Nova Conta a Pagar</h2><p>Registre uma conta com vencimento futuro</p></div>' +
'<label>Descrição *</label><input type="text" id="desc" placeholder="Ex: Mercadoria — Distribuidora X">' +
'<label>Fornecedor / Credor</label><input type="text" id="forn" placeholder="Ex: Cervejaria ABC">' +
'<div class="grid2">' +
'<div><label>Valor (R$) *</label><input type="number" id="valor" placeholder="0,00" step="0.01" min="0"></div>' +
'<div><label>Vencimento *</label><input type="date" id="venc"></div>' +
'</div>' +
'<label>Forma de Pagamento</label>' +
'<select id="pgto"><option>PIX</option><option>Boleto</option><option>Transferência</option>' +
'<option>Cartão Débito</option><option>Cartão Crédito</option><option>Dinheiro</option><option>Outro</option></select>' +
'<label>Observação</label><input type="text" id="obs" placeholder="Ex: nº do boleto, prazo negociado...">' +
'<button class="btn" onclick="salvar()">💾 Salvar Conta</button>' +
'<div class="msg" id="msg"></div>' +
'<script>' +
'function salvar(){' +
'var desc=document.getElementById("desc").value.trim();' +
'var forn=document.getElementById("forn").value.trim();' +
'var val=parseFloat(document.getElementById("valor").value);' +
'var venc=document.getElementById("venc").value;' +
'var pgto=document.getElementById("pgto").value;' +
'var obs=document.getElementById("obs").value.trim();' +
'var msg=document.getElementById("msg");' +
'if(!desc||isNaN(val)||val<=0||!venc){msg.className="msg err";msg.textContent="Preencha os campos obrigatórios.";return;}' +
'msg.className="msg";msg.textContent="Salvando...";' +
'google.script.run' +
'.withSuccessHandler(function(){msg.className="msg ok";msg.textContent="✅ Conta salva com sucesso!";' +
'document.getElementById("desc").value="";document.getElementById("forn").value="";' +
'document.getElementById("valor").value="";document.getElementById("obs").value="";})' +
'.withFailureHandler(function(e){msg.className="msg err";msg.textContent="Erro: "+e.message;})' +
'.salvarContaPagar({desc:desc,forn:forn,val:val,venc:venc,pgto:pgto,obs:obs});}' +
'<\/script></body></html>';
}

function salvarContaPagar(d) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(SHEET_CONTAS_PAGAR);
  if (!sh) throw new Error("Aba Contas a Pagar não encontrada.");
  const p    = d.venc.split("-");
  const venc = new Date(+p[0],+p[1]-1,+p[2]);
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const status = venc < hoje ? "Vencido" : "Pendente";
  sh.appendRow([d.desc,d.forn,d.val,venc,status,"",d.pgto,d.obs]);
  const ul = sh.getLastRow();
  sh.getRange(ul,3).setNumberFormat("R$ #,##0.00");
  sh.getRange(ul,4).setNumberFormat("dd/mm/yyyy");
  sh.getRange(ul,6).setNumberFormat("dd/mm/yyyy");
  return {sucesso:true};
}

// ============================================================
//  DASHBOARD
// ============================================================
function atualizarDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_DASHBOARD);
  if (!sh) sh = ss.insertSheet(SHEET_DASHBOARD);
  sh.clearContents().clearFormats();
  sh.getCharts().forEach(function(c){sh.removeChart(c);});

  const shLanc = ss.getSheetByName(SHEET_LANCAMENTOS);
  const shCP   = ss.getSheetByName(SHEET_CONTAS_PAGAR);
  var dados = [];
  if (shLanc && shLanc.getLastRow()>1)
    dados = shLanc.getRange(2,1,shLanc.getLastRow()-1,10).getValues().filter(function(r){return r[0]&&r[9];});

  const hoje     = new Date(); hoje.setHours(0,0,0,0);
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();
  const meses    = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

  const dadosMes = dados.filter(function(r){var d=new Date(r[0]);return d.getMonth()===mesAtual&&d.getFullYear()===anoAtual;});

  const diaSem   = hoje.getDay();
  const ateSeg   = diaSem===0?6:diaSem-1;
  const iniSem   = new Date(hoje); iniSem.setDate(hoje.getDate()-ateSeg);
  const fimSem   = new Date(iniSem); fimSem.setDate(iniSem.getDate()+6);
  const dadosSem = dados.filter(function(r){var d=new Date(r[0]);d.setHours(0,0,0,0);return d>=iniSem&&d<=fimSem;});

  const recMes   = dadosMes.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[9]);},0);
  const despMes  = dadosMes.filter(function(r){return r[1]==="Despesa";}).reduce(function(s,r){return s+(+r[9]);},0);
  const recSem   = dadosSem.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[9]);},0);
  const despSem  = dadosSem.filter(function(r){return r[1]==="Despesa";}).reduce(function(s,r){return s+(+r[9]);},0);
  const recTotal = dados.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[9]);},0);
  const despTotal= dados.filter(function(r){return r[1]==="Despesa";}).reduce(function(s,r){return s+(+r[9]);},0);

  var contasPend=0,contasVenc=0,totalPend=0;
  if (shCP&&shCP.getLastRow()>1) {
    shCP.getRange(2,1,shCP.getLastRow()-1,5).getValues().filter(function(r){return r[0];}).forEach(function(r){
      if(r[4]==="Pendente"){contasPend++;totalPend+=+r[2];}
      if(r[4]==="Vencido"){contasVenc++;totalPend+=+r[2];}
    });
  }

  // Título
  sh.getRange("A1").setValue("🍺 "+NOME_ADEGA)
    .setFontSize(22).setFontWeight("bold").setFontColor(COR_HEADER);
  sh.getRange("A2").setValue("Dashboard Financeiro  •  "+meses[mesAtual]+" "+anoAtual+"  •  Atualizado: "+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm"))
    .setFontColor(COR_ACCENT).setFontSize(10).setFontWeight("bold");
  sh.setRowHeight(1,44); sh.setRowHeight(2,22);

  // Cards
  const saldoMes  = recMes-despMes;
  const saldoGeral= recTotal-despTotal;
  const cards = [
    ["B4","D5","💰 Receitas\n"+meses[mesAtual],          recMes,   COR_RECEITA,"#137333"],
    ["E4","G5","💸 Despesas\n"+meses[mesAtual],          despMes,  COR_DESPESA,"#a50000"],
    ["H4","J5","📊 Saldo\n"+meses[mesAtual],             saldoMes, saldoMes>=0?COR_RECEITA:COR_DESPESA,saldoMes>=0?"#137333":"#a50000"],
    ["B6","D7","📅 Receitas\nSemana Atual",              recSem,   "#e8f0fe","#1a73e8"],
    ["E6","G7","📅 Despesas\nSemana Atual",              despSem,  "#fce8e6","#a50000"],
    ["H6","J7","💹 Saldo Geral\nTodo o período",         saldoGeral,saldoGeral>=0?COR_RECEITA:COR_DESPESA,saldoGeral>=0?"#137333":"#a50000"],
    ["B8","D9","⚠️ Total a Pagar\n(Pend. + Venc.)",     totalPend,contasVenc>0?"#fce8e6":"#fff9c4",contasVenc>0?"#a50000":"#7f4f00"],
    ["E8","G9","🔴 Contas Vencidas",                     contasVenc+" conta(s)","#fce8e6","#a50000"],
    ["H8","J9","🟡 Contas Pendentes",                    contasPend+" conta(s)","#fff9c4","#7f4f00"],
  ];

  cards.forEach(function(c){
    var ini=c[0],fim=c[1],titulo=c[2],val=c[3],bg=c[4],fg=c[5];
    var rI=parseInt(ini.slice(1)), cI=_colIdx(ini[0]), cF=_colIdx(fim[0]);
    sh.getRange(rI,cI,1,cF-cI+1).merge().setValue(titulo)
      .setBackground(bg).setFontColor(fg).setFontWeight("bold")
      .setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    var valStr=typeof val==="number"?"R$ "+val.toLocaleString("pt-BR",{minimumFractionDigits:2}):val;
    sh.getRange(rI+1,cI,1,cF-cI+1).merge().setValue(valStr)
      .setBackground(bg).setFontColor(fg).setFontSize(14).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sh.setRowHeight(rI,26); sh.setRowHeight(rI+1,36);
  });

  // Últimos 7 fechamentos
  var lin=11;
  sh.getRange(lin,2,1,7).merge().setValue("📝 Últimos Fechamentos do Dia")
    .setFontWeight("bold").setFontSize(12).setFontColor(COR_HEADER).setBackground("#f8f9fa");
  sh.setRowHeight(lin,30); lin++;
  ["Data","Débito","Crédito","PIX","Dinheiro","Delivery","Total"].forEach(function(h,i){
    sh.getRange(lin,2+i).setValue(h).setFontWeight("bold")
      .setBackground(COR_HEADER).setFontColor(COR_HEADER_T).setFontSize(10).setHorizontalAlignment("center");
  });
  sh.setRowHeight(lin,26); lin++;
  var fechamentos=dados.filter(function(r){return r[1]==="Receita"&&r[3]==="Fechamento do Dia";}).slice(-7).reverse();
  if(fechamentos.length===0){
    sh.getRange(lin,2,1,7).merge().setValue("Nenhum fechamento registrado ainda.")
      .setFontColor("#aaa").setFontStyle("italic").setFontSize(10); lin++;
  } else {
    fechamentos.forEach(function(r){
      sh.getRange(lin,2).setValue(r[0]).setNumberFormat("dd/mm/yyyy");
      [4,5,6,7,8,9].forEach(function(ci,i){sh.getRange(lin,3+i).setValue(r[ci]).setNumberFormat("R$ #,##0.00");});
      sh.getRange(lin,2,1,7).setBackground(COR_RECEITA);
      sh.setRowHeight(lin,24); lin++;
    });
  }

  // Dados auxiliares para gráficos
  var mesesGraf=[];
  for(var i=5;i>=0;i--){
    var d=new Date(anoAtual,mesAtual-i,1);
    mesesGraf.push({mes:d.getMonth(),ano:d.getFullYear(),label:meses[d.getMonth()].substring(0,3)+"/"+String(d.getFullYear()).slice(2)});
  }
  sh.getRange(1,13).setValue("Mês").setFontWeight("bold");
  sh.getRange(1,14).setValue("Receitas").setFontWeight("bold");
  sh.getRange(1,15).setValue("Despesas").setFontWeight("bold");
  mesesGraf.forEach(function(m,i){
    var rec =dados.filter(function(r){return r[1]==="Receita"&&new Date(r[0]).getMonth()===m.mes&&new Date(r[0]).getFullYear()===m.ano;}).reduce(function(s,r){return s+(+r[9]);},0);
    var desp=dados.filter(function(r){return r[1]==="Despesa"&&new Date(r[0]).getMonth()===m.mes&&new Date(r[0]).getFullYear()===m.ano;}).reduce(function(s,r){return s+(+r[9]);},0);
    sh.getRange(2+i,13).setValue(m.label);
    sh.getRange(2+i,14).setValue(rec);
    sh.getRange(2+i,15).setValue(desp);
  });

  var despCat={};
  dadosMes.filter(function(r){return r[1]==="Despesa";}).forEach(function(r){despCat[r[2]]=(despCat[r[2]]||0)+(+r[9]);});
  var catRows=Object.entries(despCat).sort(function(a,b){return b[1]-a[1];});
  sh.getRange(1,17).setValue("Categoria").setFontWeight("bold");
  sh.getRange(1,18).setValue("Valor").setFontWeight("bold");
  catRows.forEach(function(c,i){sh.getRange(2+i,17).setValue(c[0]);sh.getRange(2+i,18).setValue(c[1]);});

  if(mesesGraf.length>0){
    var barChart=sh.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(1,13,7,3))
      .setPosition(lin+1,2,0,0).setNumHeaders(1)
      .setOption("title","📊 Receitas × Despesas — Últimos 6 Meses")
      .setOption("titleTextStyle",{color:COR_HEADER,fontSize:12,bold:true})
      .setOption("width",500).setOption("height",260)
      .setOption("colors",["#2d6a4f","#c5221f"])
      .setOption("legend",{position:"bottom"})
      .setOption("vAxis",{format:"R$ #,##0"})
      .setOption("chartArea",{width:"80%",height:"65%"})
      .build();
    sh.insertChart(barChart);
  }
  if(catRows.length>0){
    var pieChart=sh.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sh.getRange(1,17,catRows.length+1,2))
      .setPosition(lin+1,7,0,0).setNumHeaders(1)
      .setOption("title","🏷️ Despesas por Categoria — "+meses[mesAtual])
      .setOption("titleTextStyle",{color:COR_HEADER,fontSize:12,bold:true})
      .setOption("width",430).setOption("height",260)
      .setOption("pieSliceText","percentage")
      .setOption("legend",{position:"right"})
      .build();
    sh.insertChart(pieChart);
  }

  sh.hideColumns(13,6);
  sh.setColumnWidth(1,18);
  [2,3,4,5,6,7,8,9,10].forEach(function(i){sh.setColumnWidth(i,110);});
}

// ============================================================
//  RESUMO ANUAL — comparativo mês a mês
// ============================================================
function atualizarResumoAnual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_ANUAL);
  if (!sh) sh = ss.insertSheet(SHEET_ANUAL);
  sh.clearContents().clearFormats();
  sh.getCharts().forEach(function(c){sh.removeChart(c);});

  const shLanc = ss.getSheetByName(SHEET_LANCAMENTOS);
  var dados = [];
  if (shLanc && shLanc.getLastRow()>1)
    dados = shLanc.getRange(2,1,shLanc.getLastRow()-1,10).getValues().filter(function(r){return r[0]&&r[9];});

  const hoje     = new Date(); hoje.setHours(0,0,0,0);
  const anoAtual = hoje.getFullYear();
  const mesesAbr = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
  const mesesFull= ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

  var anosSet = {};
  dados.forEach(function(r){anosSet[new Date(r[0]).getFullYear()]=true;});
  anosSet[anoAtual]=true;
  var anos = Object.keys(anosSet).map(Number).sort();

  sh.getRange("A1").setValue("🍺 "+NOME_ADEGA+" — Resumo Anual Comparativo")
    .setFontSize(18).setFontWeight("bold").setFontColor(COR_HEADER);
  sh.getRange("A2").setValue("Atualizado: "+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm"))
    .setFontColor("#888").setFontSize(10);
  sh.setRowHeight(1,44);

  var lin=4;

  anos.forEach(function(ano){
    sh.getRange(lin,1,1,9).merge()
      .setValue("📅 "+ano)
      .setBackground(COR_HEADER).setFontColor(COR_HEADER_T)
      .setFontWeight("bold").setFontSize(13).setHorizontalAlignment("left");
    sh.setRowHeight(lin,34); lin++;

    ["Mês","Receita (R$)","Despesa (R$)","Saldo (R$)","Débito","Crédito","PIX","Dinheiro","Delivery"].forEach(function(h,i){
      sh.getRange(lin,i+1).setValue(h).setFontWeight("bold").setBackground("#e8eaed").setFontSize(10).setHorizontalAlignment("center");
    });
    sh.setRowHeight(lin,28); lin++;

    var totalRecAno=0, totalDespAno=0;
    mesesFull.forEach(function(mes,m){
      var dadosMes=dados.filter(function(r){var d=new Date(r[0]);return d.getMonth()===m&&d.getFullYear()===ano;});
      var rec =dadosMes.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[9]);},0);
      var desp=dadosMes.filter(function(r){return r[1]==="Despesa";}).reduce(function(s,r){return s+(+r[9]);},0);
      var saldo=rec-desp;
      var deb =dadosMes.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[4]);},0);
      var cred=dadosMes.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[5]);},0);
      var pix =dadosMes.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[6]);},0);
      var din =dadosMes.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[7]);},0);
      var del =dadosMes.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[8]);},0);
      totalRecAno+=rec; totalDespAno+=desp;

      var bg=dadosMes.length===0?"#f8f9fa":saldo>=0?COR_RECEITA:COR_DESPESA;
      var fgSaldo=saldo>=0?"#137333":"#a50000";

      sh.getRange(lin,1).setValue(mes).setBackground(bg);
      sh.getRange(lin,2).setValue(rec).setNumberFormat("R$ #,##0.00").setBackground(bg);
      sh.getRange(lin,3).setValue(desp).setNumberFormat("R$ #,##0.00").setBackground(bg);
      sh.getRange(lin,4).setValue(saldo).setNumberFormat("R$ #,##0.00").setBackground(bg).setFontColor(fgSaldo).setFontWeight("bold");
      sh.getRange(lin,5).setValue(deb).setNumberFormat("R$ #,##0.00").setBackground(bg);
      sh.getRange(lin,6).setValue(cred).setNumberFormat("R$ #,##0.00").setBackground(bg);
      sh.getRange(lin,7).setValue(pix).setNumberFormat("R$ #,##0.00").setBackground(bg);
      sh.getRange(lin,8).setValue(din).setNumberFormat("R$ #,##0.00").setBackground(bg);
      sh.getRange(lin,9).setValue(del).setNumberFormat("R$ #,##0.00").setBackground(bg);
      sh.setRowHeight(lin,22); lin++;
    });

    var saldoAno=totalRecAno-totalDespAno;
    var bgT=saldoAno>=0?"#b7e1cd":"#f4c7c3";
    var fgT=saldoAno>=0?"#0d652b":"#a50000";
    sh.getRange(lin,1).setValue("TOTAL "+ano).setFontWeight("bold").setBackground(bgT).setFontColor(fgT);
    sh.getRange(lin,2).setValue(totalRecAno).setNumberFormat("R$ #,##0.00").setFontWeight("bold").setBackground(bgT).setFontColor(fgT);
    sh.getRange(lin,3).setValue(totalDespAno).setNumberFormat("R$ #,##0.00").setFontWeight("bold").setBackground(bgT).setFontColor(fgT);
    sh.getRange(lin,4).setValue(saldoAno).setNumberFormat("R$ #,##0.00").setFontWeight("bold").setBackground(bgT).setFontColor(fgT);
    [5,6,7,8,9].forEach(function(c){sh.getRange(lin,c).setBackground(bgT);});
    sh.setRowHeight(lin,30); lin+=2;
  });

  // Tabela comparativa para gráfico de linhas
  var grafLin=lin;
  var grafHeaders=["Mês"].concat(anos);
  sh.getRange(lin,1,1,grafHeaders.length).setValues([grafHeaders])
    .setFontWeight("bold").setBackground("#e8eaed");
  lin++;
  mesesAbr.forEach(function(mes,m){
    var row=[mes];
    anos.forEach(function(ano){
      var rec=dados.filter(function(r){var d=new Date(r[0]);return d.getMonth()===m&&d.getFullYear()===ano&&r[1]==="Receita";})
        .reduce(function(s,r){return s+(+r[9]);},0);
      row.push(rec);
    });
    sh.getRange(lin,1,1,row.length).setValues([row]);
    sh.setRowHeight(lin,20); lin++;
  });

  if(anos.length>0){
    var lineChart=sh.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sh.getRange(grafLin,1,13,anos.length+1))
      .setPosition(grafLin,anos.length+3,0,0).setNumHeaders(1)
      .setOption("title","📈 Receita Mensal — Comparativo Anual")
      .setOption("titleTextStyle",{color:COR_HEADER,fontSize:13,bold:true})
      .setOption("width",620).setOption("height",360)
      .setOption("legend",{position:"bottom"})
      .setOption("vAxis",{format:"R$ #,##0",title:"Receita"})
      .setOption("hAxis",{title:"Mês"})
      .setOption("curveType","function")
      .setOption("chartArea",{width:"75%",height:"65%"})
      .build();
    sh.insertChart(lineChart);
  }

  [130,135,135,135,115,115,115,115,115].forEach(function(w,i){sh.setColumnWidth(i+1,w);});
  sh.setFrozenRows(1);
  SpreadsheetApp.getUi().alert("✅ Resumo Anual atualizado!");
}

// ============================================================
//  RELATÓRIOS DETALHADOS
// ============================================================
function atualizarRelatorios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_RELATORIOS);
  if (!sh) sh = ss.insertSheet(SHEET_RELATORIOS);
  sh.clearContents().clearFormats();

  const shLanc = ss.getSheetByName(SHEET_LANCAMENTOS);
  const shCP   = ss.getSheetByName(SHEET_CONTAS_PAGAR);
  var dados = [];
  if (shLanc && shLanc.getLastRow()>1)
    dados = shLanc.getRange(2,1,shLanc.getLastRow()-1,10).getValues().filter(function(r){return r[0]&&r[9];});

  const hoje     = new Date(); hoje.setHours(0,0,0,0);
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();
  const meses    = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

  var lin=1;
  sh.getRange(lin,1).setValue("📊 RELATÓRIOS — "+NOME_ADEGA)
    .setFontSize(16).setFontWeight("bold").setFontColor(COR_HEADER);
  sh.getRange(lin+1,1).setValue("Atualizado: "+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm"))
    .setFontColor("#666").setFontSize(10);
  lin+=3;

  var totalRec =dados.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[9]);},0);
  var totalDesp=dados.filter(function(r){return r[1]==="Despesa";}).reduce(function(s,r){return s+(+r[9]);},0);

  lin=_blocoTitulo(sh,lin,"1. RESUMO GERAL (TODO O PERÍODO)");
  [["Total de Receitas",totalRec,COR_RECEITA,"#137333"],
   ["Total de Despesas",totalDesp,COR_DESPESA,"#a50000"],
   ["Resultado Líquido",totalRec-totalDesp,(totalRec-totalDesp)>=0?COR_RECEITA:COR_DESPESA,(totalRec-totalDesp)>=0?"#137333":"#a50000"]]
  .forEach(function(arr){
    sh.getRange(lin,1).setValue(arr[0]).setFontWeight("bold").setBackground(arr[2]).setFontColor(arr[3]);
    sh.getRange(lin,2).setValue(arr[1]).setNumberFormat("R$ #,##0.00").setFontWeight("bold").setBackground(arr[2]).setFontColor(arr[3]);
    lin++;
  });
  lin++;

  var dadosMes=dados.filter(function(r){var d=new Date(r[0]);return d.getMonth()===mesAtual&&d.getFullYear()===anoAtual;});
  var recMes  =dadosMes.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[9]);},0);

  lin=_blocoTitulo(sh,lin,"2. RECEITAS POR FORMA DE PAGAMENTO — "+meses[mesAtual]);
  ["Forma","Total (R$)","% do Total"].forEach(function(h,i){
    sh.getRange(lin,i+1).setValue(h).setFontWeight("bold").setBackground("#e8eaed").setFontSize(10);
  });
  lin++;
  [["💳 Débito",4],["💳 Crédito",5],["📱 PIX",6],["💵 Dinheiro",7],["🛵 Delivery",8]]
  .forEach(function(f){
    var val=dadosMes.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[f[1]]);},0);
    sh.getRange(lin,1).setValue(f[0]).setBackground(COR_RECEITA);
    sh.getRange(lin,2).setValue(val).setNumberFormat("R$ #,##0.00").setBackground(COR_RECEITA);
    sh.getRange(lin,3).setValue(recMes>0?val/recMes:0).setNumberFormat("0.0%").setBackground(COR_RECEITA);
    lin++;
  });
  lin++;

  var despMes=dadosMes.filter(function(r){return r[1]==="Despesa";}).reduce(function(s,r){return s+(+r[9]);},0);
  var despCat={};
  dadosMes.filter(function(r){return r[1]==="Despesa";}).forEach(function(r){despCat[r[2]]=(despCat[r[2]]||0)+(+r[9]);});

  lin=_blocoTitulo(sh,lin,"3. DESPESAS POR CATEGORIA — "+meses[mesAtual]);
  ["Categoria","Total (R$)","% do Total"].forEach(function(h,i){
    sh.getRange(lin,i+1).setValue(h).setFontWeight("bold").setBackground("#e8eaed").setFontSize(10);
  });
  lin++;
  Object.entries(despCat).sort(function(a,b){return b[1]-a[1];}).forEach(function(c){
    sh.getRange(lin,1).setValue(c[0]).setBackground(COR_DESPESA);
    sh.getRange(lin,2).setValue(c[1]).setNumberFormat("R$ #,##0.00").setBackground(COR_DESPESA);
    sh.getRange(lin,3).setValue(despMes>0?c[1]/despMes:0).setNumberFormat("0.0%").setBackground(COR_DESPESA);
    lin++;
  });
  lin++;

  lin=_blocoTitulo(sh,lin,"4. FECHAMENTO DIÁRIO — "+meses[mesAtual]);
  ["Data","Débito","Crédito","PIX","Dinheiro","Delivery","Total Dia"].forEach(function(h,i){
    sh.getRange(lin,i+1).setValue(h).setFontWeight("bold").setBackground("#e8eaed").setFontSize(10);
  });
  lin++;
  dadosMes.filter(function(r){return r[1]==="Receita"&&r[3]==="Fechamento do Dia";})
    .sort(function(a,b){return new Date(b[0])-new Date(a[0]);})
    .forEach(function(r){
      sh.getRange(lin,1).setValue(r[0]).setNumberFormat("dd/mm/yyyy").setBackground(COR_RECEITA);
      [4,5,6,7,8,9].forEach(function(ci,i){sh.getRange(lin,2+i).setValue(r[ci]).setNumberFormat("R$ #,##0.00").setBackground(COR_RECEITA);});
      lin++;
    });
  lin++;

  lin=_blocoTitulo(sh,lin,"5. CONTAS A PAGAR — Pendentes e Vencidas");
  ["Descrição","Fornecedor","Valor (R$)","Vencimento","Status"].forEach(function(h,i){
    sh.getRange(lin,i+1).setValue(h).setFontWeight("bold").setBackground("#e8eaed").setFontSize(10);
  });
  lin++;
  if(shCP&&shCP.getLastRow()>1){
    shCP.getRange(2,1,shCP.getLastRow()-1,5).getValues()
      .filter(function(r){return r[0]&&(r[4]==="Pendente"||r[4]==="Vencido");})
      .sort(function(a,b){return new Date(a[3])-new Date(b[3]);})
      .forEach(function(r){
        var bg=r[4]==="Vencido"?"#fce8e6":"#fff9c4";
        sh.getRange(lin,1).setValue(r[0]).setBackground(bg);
        sh.getRange(lin,2).setValue(r[1]).setBackground(bg);
        sh.getRange(lin,3).setValue(r[2]).setNumberFormat("R$ #,##0.00").setBackground(bg);
        sh.getRange(lin,4).setValue(r[3]).setNumberFormat("dd/mm/yyyy").setBackground(bg);
        sh.getRange(lin,5).setValue(r[4]).setFontWeight("bold")
          .setFontColor(r[4]==="Vencido"?"#a50000":"#7f4f00").setBackground(bg);
        lin++;
      });
  } else {
    sh.getRange(lin,1).setValue("✅ Nenhuma conta pendente.").setFontColor("#999"); lin++;
  }

  [200,190,130,120,110,110,140].forEach(function(w,i){sh.setColumnWidth(i+1,w);});
  SpreadsheetApp.getUi().alert("✅ Relatórios atualizados!");
}

// ============================================================
//  RESUMO SEMANAL — WHATSAPP
// ============================================================
function gerarResumoWhatsApp() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shLanc= ss.getSheetByName(SHEET_LANCAMENTOS);
  const shCP  = ss.getSheetByName(SHEET_CONTAS_PAGAR);
  var dados=[];
  if(shLanc&&shLanc.getLastRow()>1)
    dados=shLanc.getRange(2,1,shLanc.getLastRow()-1,10).getValues().filter(function(r){return r[0]&&r[9];});

  // Semana anterior completa (seg a dom)
  var hoje=new Date(); hoje.setHours(0,0,0,0);
  var diaSem=hoje.getDay();
  var ateSeg=diaSem===0?6:diaSem-1;
  var iniSemAtual=new Date(hoje); iniSemAtual.setDate(hoje.getDate()-ateSeg);
  var iniSemAnt=new Date(iniSemAtual); iniSemAnt.setDate(iniSemAtual.getDate()-7);
  var fimSemAnt=new Date(iniSemAtual); fimSemAnt.setDate(iniSemAtual.getDate()-1);

  var dadosSem=dados.filter(function(r){var d=new Date(r[0]);d.setHours(0,0,0,0);return d>=iniSemAnt&&d<=fimSemAnt;});

  var iniRetras=new Date(iniSemAnt); iniRetras.setDate(iniSemAnt.getDate()-7);
  var fimRetras=new Date(iniSemAnt); fimRetras.setDate(iniSemAnt.getDate()-1);
  var dadosRetras=dados.filter(function(r){var d=new Date(r[0]);d.setHours(0,0,0,0);return d>=iniRetras&&d<=fimRetras;});

  var recSem   =dadosSem.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[9]);},0);
  var despSem  =dadosSem.filter(function(r){return r[1]==="Despesa";}).reduce(function(s,r){return s+(+r[9]);},0);
  var recRetras=dadosRetras.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[9]);},0);
  var variacao =recRetras>0?((recSem-recRetras)/recRetras*100):0;
  var varSinal =variacao>=0?"📈 +":"📉 ";
  var varEmoji =variacao>=0?"✅":"⚠️";

  var mesSem=iniSemAnt.getMonth(), anoSem=iniSemAnt.getFullYear();
  var recMes=dados.filter(function(r){return r[1]==="Receita"&&new Date(r[0]).getMonth()===mesSem&&new Date(r[0]).getFullYear()===anoSem;})
    .reduce(function(s,r){return s+(+r[9]);},0);

  var diasLabel=["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];
  var linhasDia="";
  for(var i=0;i<7;i++){
    var d=new Date(iniSemAnt); d.setDate(iniSemAnt.getDate()+i);
    var recDia=dadosSem.filter(function(r){var dr=new Date(r[0]);dr.setHours(0,0,0,0);return dr.getTime()===d.getTime()&&r[1]==="Receita";})
      .reduce(function(s,r){return s+(+r[9]);},0);
    linhasDia+=diasLabel[d.getDay()]+" "+Utilities.formatDate(d,Session.getScriptTimeZone(),"dd/MM")+" → R$ "+_fmt(recDia)+"\n";
  }

  var deb =dadosSem.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[4]);},0);
  var cred=dadosSem.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[5]);},0);
  var pix =dadosSem.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[6]);},0);
  var din =dadosSem.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[7]);},0);
  var del =dadosSem.filter(function(r){return r[1]==="Receita";}).reduce(function(s,r){return s+(+r[8]);},0);

  var despCat={};
  dadosSem.filter(function(r){return r[1]==="Despesa";}).forEach(function(r){despCat[r[2]]=(despCat[r[2]]||0)+(+r[9]);});
  var linhasDespCat="";
  Object.entries(despCat).sort(function(a,b){return b[1]-a[1];}).forEach(function(c){
    linhasDespCat+="• "+c[0]+": R$ "+_fmt(c[1])+"\n";
  });
  if(!linhasDespCat) linhasDespCat="• Nenhuma despesa lançada\n";

  var linhasContas="";
  if(shCP&&shCP.getLastRow()>1){
    shCP.getRange(2,1,shCP.getLastRow()-1,5).getValues()
      .filter(function(r){return r[0]&&(r[4]==="Pendente"||r[4]==="Vencido");})
      .sort(function(a,b){return new Date(a[3])-new Date(b[3]);})
      .slice(0,6)
      .forEach(function(r){
        var emoji=r[4]==="Vencido"?"🔴":"🟡";
        linhasContas+=emoji+" "+r[0]+" — R$ "+_fmt(+r[2])+" (vence "+Utilities.formatDate(new Date(r[3]),Session.getScriptTimeZone(),"dd/MM")+")\n";
      });
  }
  if(!linhasContas) linhasContas="✅ Nenhuma conta pendente\n";

  var iniStr=Utilities.formatDate(iniSemAnt,Session.getScriptTimeZone(),"dd/MM");
  var fimStr=Utilities.formatDate(fimSemAnt,Session.getScriptTimeZone(),"dd/MM");
  var mesesN=["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];

  var texto="🍺 *"+NOME_ADEGA+" — Fechamento Semanal*\n"+
"📅 "+iniStr+" a "+fimStr+"\n"+
"━━━━━━━━━━━━━━━━━━━━\n\n"+
"*💰 RECEITAS POR DIA:*\n"+linhasDia+"\n"+
"*💳 POR FORMA DE PAGAMENTO:*\n"+
"• Débito: R$ "+_fmt(deb)+"\n"+
"• Crédito: R$ "+_fmt(cred)+"\n"+
"• PIX: R$ "+_fmt(pix)+"\n"+
"• Dinheiro: R$ "+_fmt(din)+"\n"+
"• Delivery: R$ "+_fmt(del)+"\n\n"+
"💰 *Total da Semana:* R$ "+_fmt(recSem)+"\n"+
"📊 *Semana Anterior:* R$ "+_fmt(recRetras)+"\n"+
varSinal+Math.abs(variacao).toFixed(1)+"% "+varEmoji+"\n"+
"📈 *Acumulado "+mesesN[mesSem]+":* R$ "+_fmt(recMes)+"\n\n"+
"━━━━━━━━━━━━━━━━━━━━\n"+
"*💸 DESPESAS DA SEMANA:*\n"+linhasDespCat+
"💸 *Total Despesas:* R$ "+_fmt(despSem)+"\n\n"+
"━━━━━━━━━━━━━━━━━━━━\n"+
"*📅 CONTAS A PAGAR:*\n"+linhasContas+"\n"+
"━━━━━━━━━━━━━━━━━━━━\n"+
"💹 *Resultado Líquido:* R$ "+_fmt(recSem-despSem)+"\n"+
"━━━━━━━━━━━━━━━━━━━━\n"+
"_Gerado automaticamente — "+NOME_ADEGA+"_ 🤖";

  var textoEscapado=texto.replace(/</g,"&lt;").replace(/>/g,"&gt;");
  var htmlStr='<!DOCTYPE html><html><head>'+
'<meta name="viewport" content="width=device-width,initial-scale=1">'+
'<style>'+
'@import url(\'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600&display=swap\');'+
'*{box-sizing:border-box;font-family:\'DM Sans\',Arial,sans-serif;}'+
'body{margin:0;padding:16px;background:#f0f4f0;}'+
'.header{background:linear-gradient(135deg,#1a3a5c,#25D366);border-radius:12px;padding:14px 16px;margin-bottom:14px;color:#fff;}'+
'.header h2{margin:0 0 2px;font-size:16px;font-weight:600;}.header p{margin:0;font-size:11px;opacity:.8;}'+
'p{margin:0 0 10px;font-size:12px;color:#555;}'+
'textarea{width:100%;height:400px;padding:12px;border:1.5px solid #ccc;border-radius:8px;'+
'font-size:12px;font-family:monospace;background:#fff;resize:none;line-height:1.7;}'+
'.btn{width:100%;margin-top:10px;padding:13px;background:#25D366;color:#fff;border:none;'+
'border-radius:9px;font-size:15px;font-weight:700;cursor:pointer;}'+
'.btn:hover{background:#1da851;}'+
'.msg{text-align:center;font-size:13px;margin-top:6px;color:#137333;min-height:16px;font-weight:600;}'+
'</style></head><body>'+
'<div class="header"><h2>📲 Resumo Semanal — WhatsApp</h2><p>Semana de '+iniStr+' a '+fimStr+'</p></div>'+
'<p>Copie o texto abaixo e cole no grupo do WhatsApp:</p>'+
'<textarea id="txt">'+textoEscapado+'</textarea>'+
'<button class="btn" onclick="copiar()">📋 Copiar Texto</button>'+
'<div class="msg" id="msg"></div>'+
'<script>function copiar(){var t=document.getElementById("txt");t.select();t.setSelectionRange(0,99999);'+
'document.execCommand("copy");document.getElementById("msg").textContent="✅ Copiado! Cole no WhatsApp.";}<\/script>'+
'</body></html>';

  var html=HtmlService.createHtmlOutput(htmlStr)
    .setWidth(500).setHeight(620).setTitle("📲 Resumo WhatsApp");
  SpreadsheetApp.getUi().showModalDialog(html,"📲 Resumo Semanal — WhatsApp");
}

// ============================================================
//  ALERTAR VENCIMENTOS
// ============================================================
function alertarVencimentos() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const shCP=ss.getSheetByName(SHEET_CONTAS_PAGAR);
  if(!shCP||shCP.getLastRow()<=1){SpreadsheetApp.getUi().alert("Nenhuma conta cadastrada.");return;}
  var hoje=new Date(); hoje.setHours(0,0,0,0);
  var em7=new Date(hoje); em7.setDate(em7.getDate()+7);
  var contas=shCP.getRange(2,1,shCP.getLastRow()-1,5).getValues().filter(function(r){return r[0];});
  var vencidas=contas.filter(function(r){return new Date(r[3])<hoje&&r[4]!=="Pago"&&r[4]!=="Cancelado";});
  var proximas=contas.filter(function(r){var d=new Date(r[3]);return d>=hoje&&d<=em7&&r[4]!=="Pago"&&r[4]!=="Cancelado";});
  var msg="";
  if(vencidas.length){
    msg+="🔴 VENCIDAS ("+vencidas.length+"):\n";
    vencidas.forEach(function(r){msg+="  • "+r[0]+" — R$ "+_fmt(+r[2])+" — venceu "+Utilities.formatDate(new Date(r[3]),Session.getScriptTimeZone(),"dd/MM/yyyy")+"\n";});
    msg+="\n";
  }
  if(proximas.length){
    msg+="🟡 VENCEM EM ATÉ 7 DIAS ("+proximas.length+"):\n";
    proximas.forEach(function(r){msg+="  • "+r[0]+" — R$ "+_fmt(+r[2])+" — vence "+Utilities.formatDate(new Date(r[3]),Session.getScriptTimeZone(),"dd/MM/yyyy")+"\n";});
  }
  if(!msg) msg="✅ Nenhuma conta vencida ou vencendo nos próximos 7 dias.";
  SpreadsheetApp.getUi().alert("⚠️ Vencimentos",msg,SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
//  WEB APP — doGet e doPost (app mobile)
// ============================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Adega Postin — Lançamentos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var payload=JSON.parse(e.postData.contents);
    var resultado;
    switch(payload.acao){
      case 'lancar_fechamento':  resultado=salvarFechamentoDia(payload.dados); break;
      case 'lancar_despesa':     resultado=salvarDespesa(payload.dados);       break;
      case 'lancar_conta_pagar': resultado=salvarContaPagar(payload.dados);    break;
      case 'buscar_categorias':  resultado=_getCategorias();                   break;
      case 'buscar_contas_pagar': resultado=_getContasPagar();                  break;
      case 'pagar_conta':        resultado=_pagarConta(payload.dados);          break;
      case 'resumo_semana':      resultado=_getResumoSemana();                  break;
      case 'resumo_hoje':        resultado=_getResumoHoje();                   break;
      default: resultado={sucesso:false,mensagem:'Ação desconhecida.'};
    }
    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({sucesso:false,mensagem:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function _getCategorias(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(SHEET_CATEGORIAS);
  if(!sh) return {sucesso:false,mensagem:"Aba Categorias não encontrada."};
  var cats=sh.getRange("A2:B200").getValues().filter(function(r){return r[0]&&r[1]==="Despesa";}).map(function(r){return r[0];});
  return {sucesso:true,categorias:cats};
}

function _getResumoHoje(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(SHEET_LANCAMENTOS);
  if(!sh||sh.getLastRow()<=1) return {sucesso:true,receita:0,despesa:0,saldo:0};
  var hoje=new Date(); hoje.setHours(0,0,0,0);
  var dados=sh.getRange(2,1,sh.getLastRow()-1,10).getValues().filter(function(r){return r[0];});
  var rec =dados.filter(function(r){var d=new Date(r[0]);d.setHours(0,0,0,0);return d.getTime()===hoje.getTime()&&r[1]==="Receita";}).reduce(function(s,r){return s+(+r[9]);},0);
  var desp=dados.filter(function(r){var d=new Date(r[0]);d.setHours(0,0,0,0);return d.getTime()===hoje.getTime()&&r[1]==="Despesa";}).reduce(function(s,r){return s+(+r[9]);},0);
  return {sucesso:true,receita:rec,despesa:desp,saldo:rec-desp};
}

// ============================================================
//  HELPERS
// ============================================================
function _estilizarHeader(sh,linha,numCols){
  sh.getRange(linha,1,1,numCols)
    .setBackground(COR_HEADER).setFontColor(COR_HEADER_T)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(linha,34);
}
function _blocoTitulo(sh,linha,titulo){
  sh.getRange(linha,1,1,7).merge().setValue(titulo)
    .setBackground(COR_HEADER).setFontColor(COR_HEADER_T)
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");
  sh.setRowHeight(linha,30);
  return linha+1;
}
function _colIdx(letra){return {A:1,B:2,C:3,D:4,E:5,F:6,G:7,H:8,I:9,J:10}[letra]||1;}
function _hoje(){
  var d=new Date();
  return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
}
function _dataOffset(dias){
  var d=new Date(); d.setDate(d.getDate()+dias);
  return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
}
function _fmt(v){return (v||0).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2});}

// ============================================================
//  FUNÇÕES EXTRAS — usadas pelo app mobile (doPost)
// ============================================================

// Retorna contas pendentes/vencidas para o app
function _getContasPagar() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName(SHEET_CONTAS_PAGAR);
  if (!sh || sh.getLastRow() <= 1) return {sucesso:true, contas:[]};
  var dados = sh.getRange(2,1,sh.getLastRow()-1,5).getValues().filter(function(r){return r[0];});
  var contas = [];
  dados.forEach(function(r, i) {
    if (r[4] === 'Pendente' || r[4] === 'Vencido') {
      contas.push({
        linha:      i + 2,
        desc:       r[0],
        forn:       r[1],
        valor:      r[2],
        vencimento: Utilities.formatDate(new Date(r[3]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        status:     r[4]
      });
    }
  });
  contas.sort(function(a,b){return new Date(a.vencimento)-new Date(b.vencimento);});
  return {sucesso:true, contas:contas};
}

// Marca conta como paga
function _pagarConta(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CONTAS_PAGAR);
  if (!sh) return {sucesso:false, mensagem:'Aba não encontrada.'};
  sh.getRange(d.linha, 5).setValue('Pago');
  sh.getRange(d.linha, 6).setValue(new Date()).setNumberFormat('dd/mm/yyyy');
  return {sucesso:true};
}

// Resumo da semana anterior para WhatsApp (gerado no servidor)
function _getResumoSemana() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var shLanc= ss.getSheetByName(SHEET_LANCAMENTOS);
  var shCP  = ss.getSheetByName(SHEET_CONTAS_PAGAR);
  var dados = [];
  if (shLanc && shLanc.getLastRow() > 1)
    dados = shLanc.getRange(2,1,shLanc.getLastRow()-1,10).getValues().filter(function(r){return r[0]&&r[9];});

  var hoje = new Date(); hoje.setHours(0,0,0,0);
  var diaSem = hoje.getDay();
  var ateSeg = diaSem === 0 ? 6 : diaSem - 1;
  var iniSemAtual = new Date(hoje); iniSemAtual.setDate(hoje.getDate() - ateSeg);
  var iniSemAnt   = new Date(iniSemAtual); iniSemAnt.setDate(iniSemAtual.getDate() - 7);
  var fimSemAnt   = new Date(iniSemAtual); fimSemAnt.setDate(iniSemAtual.getDate() - 1);

  var dadosSem = dados.filter(function(r){var d=new Date(r[0]);d.setHours(0,0,0,0);return d>=iniSemAnt&&d<=fimSemAnt;});
  var iniRetras= new Date(iniSemAnt); iniRetras.setDate(iniSemAnt.getDate()-7);
  var fimRetras= new Date(iniSemAnt); fimRetras.setDate(iniSemAnt.getDate()-1);
  var dadosRetras=dados.filter(function(r){var d=new Date(r[0]);d.setHours(0,0,0,0);return d>=iniRetras&&d<=fimRetras;});

  var recSem   =dadosSem.filter(function(r){return r[1]==='Receita';}).reduce(function(s,r){return s+(+r[9]);},0);
  var despSem  =dadosSem.filter(function(r){return r[1]==='Despesa';}).reduce(function(s,r){return s+(+r[9]);},0);
  var recRetras=dadosRetras.filter(function(r){return r[1]==='Receita';}).reduce(function(s,r){return s+(+r[9]);},0);
  var variacao =recRetras>0?((recSem-recRetras)/recRetras*100):0;
  var varSinal =variacao>=0?'📈 +':'📉 ';
  var varEmoji =variacao>=0?'✅':'⚠️';

  var mesSem=iniSemAnt.getMonth(), anoSem=iniSemAnt.getFullYear();
  var recMes=dados.filter(function(r){return r[1]==='Receita'&&new Date(r[0]).getMonth()===mesSem&&new Date(r[0]).getFullYear()===anoSem;})
    .reduce(function(s,r){return s+(+r[9]);},0);

  var diasLabel=['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
  var linhasDia='';
  for(var i=0;i<7;i++){
    var d=new Date(iniSemAnt); d.setDate(iniSemAnt.getDate()+i);
    var recDia=dadosSem.filter(function(r){var dr=new Date(r[0]);dr.setHours(0,0,0,0);return dr.getTime()===d.getTime()&&r[1]==='Receita';})
      .reduce(function(s,r){return s+(+r[9]);},0);
    linhasDia+=diasLabel[d.getDay()]+' '+Utilities.formatDate(d,Session.getScriptTimeZone(),'dd/MM')+' → R$ '+_fmt(recDia)+'\n';
  }

  var deb =dadosSem.filter(function(r){return r[1]==='Receita';}).reduce(function(s,r){return s+(+r[4]);},0);
  var cred=dadosSem.filter(function(r){return r[1]==='Receita';}).reduce(function(s,r){return s+(+r[5]);},0);
  var pix =dadosSem.filter(function(r){return r[1]==='Receita';}).reduce(function(s,r){return s+(+r[6]);},0);
  var din =dadosSem.filter(function(r){return r[1]==='Receita';}).reduce(function(s,r){return s+(+r[7]);},0);
  var del =dadosSem.filter(function(r){return r[1]==='Receita';}).reduce(function(s,r){return s+(+r[8]);},0);

  var despCat={};
  dadosSem.filter(function(r){return r[1]==='Despesa';}).forEach(function(r){despCat[r[2]]=(despCat[r[2]]||0)+(+r[9]);});
  var linhasDespCat='';
  Object.entries(despCat).sort(function(a,b){return b[1]-a[1];}).forEach(function(c){
    linhasDespCat+='• '+c[0]+': R$ '+_fmt(c[1])+'\n';
  });
  if(!linhasDespCat) linhasDespCat='• Nenhuma despesa lançada\n';

  var linhasContas='';
  if(shCP&&shCP.getLastRow()>1){
    shCP.getRange(2,1,shCP.getLastRow()-1,5).getValues()
      .filter(function(r){return r[0]&&(r[4]==='Pendente'||r[4]==='Vencido');})
      .sort(function(a,b){return new Date(a[3])-new Date(b[3]);})
      .slice(0,5)
      .forEach(function(r){
        var emoji=r[4]==='Vencido'?'🔴':'🟡';
        linhasContas+=emoji+' '+r[0]+' — R$ '+_fmt(+r[2])+' (vence '+Utilities.formatDate(new Date(r[3]),Session.getScriptTimeZone(),'dd/MM')+')\n';
      });
  }
  if(!linhasContas) linhasContas='✅ Nenhuma conta pendente\n';

  var iniStr=Utilities.formatDate(iniSemAnt,Session.getScriptTimeZone(),'dd/MM');
  var fimStr=Utilities.formatDate(fimSemAnt,Session.getScriptTimeZone(),'dd/MM');
  var mesesN=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

  var texto='🍺 *'+NOME_ADEGA+' — Fechamento Semanal*\n'+
'📅 '+iniStr+' a '+fimStr+'\n'+
'━━━━━━━━━━━━━━━━━━━━\n\n'+
'*💰 RECEITAS POR DIA:*\n'+linhasDia+'\n'+
'*💳 POR FORMA DE PAGAMENTO:*\n'+
'• Débito: R$ '+_fmt(deb)+'\n'+
'• Crédito: R$ '+_fmt(cred)+'\n'+
'• PIX: R$ '+_fmt(pix)+'\n'+
'• Dinheiro: R$ '+_fmt(din)+'\n'+
'• Delivery: R$ '+_fmt(del)+'\n\n'+
'💰 *Total da Semana:* R$ '+_fmt(recSem)+'\n'+
'📊 *Semana Anterior:* R$ '+_fmt(recRetras)+'\n'+
varSinal+Math.abs(variacao).toFixed(1)+'% '+varEmoji+'\n'+
'📈 *Acumulado '+mesesN[mesSem]+':* R$ '+_fmt(recMes)+'\n\n'+
'━━━━━━━━━━━━━━━━━━━━\n'+
'*💸 DESPESAS DA SEMANA:*\n'+linhasDespCat+
'💸 *Total Despesas:* R$ '+_fmt(despSem)+'\n\n'+
'━━━━━━━━━━━━━━━━━━━━\n'+
'*📅 CONTAS A PAGAR:*\n'+linhasContas+'\n'+
'━━━━━━━━━━━━━━━━━━━━\n'+
'💹 *Resultado Líquido:* R$ '+_fmt(recSem-despSem)+'\n'+
'━━━━━━━━━━━━━━━━━━━━\n'+
'_Gerado automaticamente — '+NOME_ADEGA+'_ 🤖';

  return {sucesso:true, texto:texto};
}
