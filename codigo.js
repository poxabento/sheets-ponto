/**
 * SISTEMA INTEGRADO DE CONTROLE DE PONTO E BANCO DE HORAS COM DASHBOARD
 * Jornada Padrão: 8h48min por dia útil (528 minutos)
 */

const JORNADA_PADRAO_MINUTOS = (8 * 60) + 48; // 528 minutos / 8h48 por dia útil
const NOME_ABA_DASHBOARD = "📊 Painel Geral";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⏰ Registro de Ponto')
    .addItem('Bater Ponto Agora', 'registrarPonto')
    .addItem('Abrir / Atualizar Dashboard', 'atualizarDashboard')
    .addSeparator()
    .addItem('Gerar Calendário do Mês Atual', 'gerarDatasDoMes')
    .addItem('Recalcular Tudo e Exibir Resumo', 'recalcularETrazerDetalhes')
    .addItem('Gerar e Enviar PDF por E-mail', 'enviarRelatorioPDF')
    .addToUi();
}

/**
 * Converte horário HH:MM para minutos inteiros
 */
function textoHoraParaMinutos(str) {
  if (!str) return null;
  const textoLimpo = String(str).trim();
  if (textoLimpo === "" || textoLimpo === "-") return null;

  const partes = textoLimpo.split(':');
  if (partes.length >= 2) {
    const h = parseInt(partes[0], 10);
    const m = parseInt(partes[1], 10);
    if (!isNaN(h) && !isNaN(m)) {
      return (h * 60) + m;
    }
  }
  return null;
}

/**
 * Formata minutos em string de hora HH:MM
 */
function minutosParaHorasFormatado(minutos, comSinal = false) {
  const sinal = minutos < 0 ? "-" : (comSinal && minutos > 0 ? "+" : "");
  const minsAbs = Math.abs(minutos);
  const hrs = Math.floor(minsAbs / 60);
  const mins = minsAbs % 60;
  return `${sinal}${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
}

/**
 * Localiza com segurança a aba de dados do ponto.
 * Prioriza a aba ativa do usuário e evita escolher a aba errada quando
 * existem folhas extras no arquivo.
 */
function obterAbaPonto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const abaAtiva = ss.getActiveSheet();
  if (abaAtiva && abaAtiva.getName() !== NOME_ABA_DASHBOARD) {
    return abaAtiva;
  }

  for (let s of sheets) {
    if (s.getName() !== NOME_ABA_DASHBOARD) return s;
  }

  return sheets[0] || ss.insertSheet("Ponto");
}

/**
 * Recalcula o ponto e atualiza a Dashboard
 */
function recalcularETrazerDetalhes(exibirAlerta = true) {
  const sheet = obterAbaPonto();
  const ultimaLinha = sheet.getLastRow();
  
  let dadosResumo = {
    diasUteis: 0,
    diasTrabalhados: 0,
    totalTrabalhado: 0,
    totalExtras: 0,
    totalFaltas: 0,
    saldoBanco: 0
  };

  if (ultimaLinha >= 5) {
    const numLinhas = ultimaLinha - 4;
    const dados = sheet.getRange(5, 1, numLinhas, 6).getDisplayValues();

    let matrizResultados = [];
    let coresExtras = [];
    let coresFaltas = [];

    for (let i = 0; i < dados.length; i++) {
      const dataStr = dados[i][0];
      const diaSemana = String(dados[i][1] || '').toLowerCase().trim();
      
      const ent1 = textoHoraParaMinutos(dados[i][2]);
      const sai1 = textoHoraParaMinutos(dados[i][3]);
      const ent2 = textoHoraParaMinutos(dados[i][4]);
      const sai2 = textoHoraParaMinutos(dados[i][5]);

      const ehFimDeSemana = diaSemana.includes("sábado") || diaSemana.includes("sabado") || diaSemana.includes("domingo");
      if (!ehFimDeSemana && dataStr !== "") dadosResumo.diasUteis++;

      if (ent1 === null || sai2 === null) {
        matrizResultados.push(["", "", ""]);
        coresExtras.push(["#ffffff"]);
        coresFaltas.push(["#ffffff"]);
        continue;
      }

      dadosResumo.diasTrabalhados++;
      let minTrabalhados = 0;

      if (sai1 !== null && ent2 !== null) {
        let turno1 = sai1 - ent1;
        let turno2 = sai2 - ent2;
        if (turno1 < 0) turno1 += 1440;
        if (turno2 < 0) turno2 += 1440;
        minTrabalhados = turno1 + turno2;
      } else {
        minTrabalhados = sai2 - ent1;
        if (minTrabalhados < 0) minTrabalhados += 1440;
      }

      const metaDiaMin = ehFimDeSemana ? 0 : JORNADA_PADRAO_MINUTOS;
      const saldoDiaMin = minTrabalhados - metaDiaMin;

      dadosResumo.totalTrabalhado += minTrabalhados;
      dadosResumo.saldoBanco += saldoDiaMin;

      let hExtraStr = "00:00";
      let faltaStr = "00:00";
      let corH = "#ffffff";
      let corI = "#ffffff";

      if (saldoDiaMin > 0) {
        hExtraStr = minutosParaHorasFormatado(saldoDiaMin);
        dadosResumo.totalExtras += saldoDiaMin;
        corH = "#dcfce7";
      } else if (saldoDiaMin < 0) {
        faltaStr = minutosParaHorasFormatado(Math.abs(saldoDiaMin));
        dadosResumo.totalFaltas += Math.abs(saldoDiaMin);
        corI = "#fee2e2";
      }

      matrizResultados.push([minutosParaHorasFormatado(minTrabalhados), hExtraStr, faltaStr]);
      coresExtras.push([corH]);
      coresFaltas.push([corI]);
    }

    const intervaloResultados = sheet.getRange(5, 7, numLinhas, 3);
    intervaloResultados.clearContent();
    intervaloResultados.setNumberFormat("@");
    intervaloResultados.setValues(matrizResultados);
    intervaloResultados.setBorder(true, true, true, true, true, true, "#475569", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    sheet.getRange(5, 8, coresExtras.length, 1).setBackgrounds(coresExtras);
    sheet.getRange(5, 9, coresFaltas.length, 1).setBackgrounds(coresFaltas);

    const intervaloTabelaCompleta = sheet.getRange(5, 1, numLinhas, 9);
    intervaloTabelaCompleta.setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  // Renderiza a Dashboard corrigida
  renderizarDashboard(dadosResumo, sheet);

  if (exibirAlerta) {
    const saldoFinalStr = minutosParaHorasFormatado(dadosResumo.saldoBanco, true);
    const statusBanco = dadosResumo.saldoBanco >= 0 ? "🟢 POSITIVO" : "🔴 NEGATIVO";
    const msg = `📊 RESUMO ATUALIZADO\n\n` +
      `📅 Dias Trabalhados: ${dadosResumo.diasTrabalhados} / ${dadosResumo.diasUteis}\n` +
      `⏱️ Horas Trabalhadas: ${minutosParaHorasFormatado(dadosResumo.totalTrabalhado)} h\n` +
      `➕ Extras: ${minutosParaHorasFormatado(dadosResumo.totalExtras)} h\n` +
      `➖ Faltas: ${minutosParaHorasFormatado(dadosResumo.totalFaltas)} h\n` +
      `🏆 Saldo do Banco: ${saldoFinalStr} h (${statusBanco})`;
    SpreadsheetApp.getUi().alert(msg);
  }

  return dadosResumo;
}

/**
 * DESENHA A DASHBOARD SEM TRAVAMENTOS (COM UNMERGE SEGURO)
 */
function renderizarDashboard(resumo, sheetPonto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName(NOME_ABA_DASHBOARD);

  if (!dash) {
    dash = ss.insertSheet(NOME_ABA_DASHBOARD, 0);
  } else {
    // CORREÇÃO CRÍTICA: Desfaz todas as células mescladas antes de limpar
    try {
      dash.getDataRange().breakApart();
    } catch (err) {}
    dash.clear();
    ss.setActiveSheet(dash);
    ss.moveActiveSheet(1);
  }

  // Garante ao menos 22 colunas e 40 linhas para acomodar os gráficos
  if (dash.getMaxColumns() < 22) dash.insertColumnsAfter(dash.getMaxColumns(), 22 - dash.getMaxColumns());
  if (dash.getMaxRows() < 40) dash.insertRowsAfter(dash.getMaxRows(), 40 - dash.getMaxRows());

  dash.setHiddenGridlines(true);

  // Paleta de Cores Moderna
  const PALETA = {
    canvas:     "#F8FAFC",
    card:       "#FFFFFF",
    borda:      "#E2E8F0",
    escuro:     "#0F172A",
    cinzaTexto: "#475569",
    cinzaClaro: "#94A3B8",
    azul:       "#2563EB",
    verde:      "#16A34A",
    vermelho:   "#DC2626"
  };

  dash.getRange(1, 1, 40, 22).setBackground(PALETA.canvas);

  // Larguras ajustadas
  dash.setColumnWidth(1, 20);  // A
  dash.setColumnWidth(2, 90);  // B
  dash.setColumnWidth(3, 85);  // C
  dash.setColumnWidth(4, 90);  // D
  dash.setColumnWidth(5, 85);  // E
  dash.setColumnWidth(6, 90);  // F
  dash.setColumnWidth(7, 85);  // G
  dash.setColumnWidth(8, 90);  // H
  dash.setColumnWidth(9, 85);  // I
  dash.setColumnWidth(10, 75); // J
  dash.setColumnWidth(11, 75); // K
  dash.setColumnWidth(12, 75); // L
  dash.setColumnWidth(13, 80); // M

  // --- CABEÇALHO ---
  dash.getRange("B2:M2").merge()
    .setValue("CONTROLE DE PONTO & BANCO DE HORAS")
    .setBackground(PALETA.escuro)
    .setFontColor("#FFFFFF")
    .setFontFamily("Segoe UI")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.escuro, SpreadsheetApp.BorderStyle.SOLID);

  const fuso = ss.getSpreadsheetTimeZone();
  const agoraStr = Utilities.formatDate(new Date(), fuso, "dd/MM/yyyy HH:mm");
  dash.getRange("B3:M3").merge()
    .setValue(`Folha Ativa: ${sheetPonto.getName()}   |   Última sincronização: ${agoraStr}   |   Jornada: 08:48/dia útil`)
    .setBackground("#F8FAFC")
    .setFontFamily("Segoe UI")
    .setFontSize(9)
    .setFontColor(PALETA.cinzaTexto)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange("B4:M4").merge()
    .setValue("RESUMO GERAL DO MÊS")
    .setBackground("#E2E8F0")
    .setFontFamily("Segoe UI")
    .setFontSize(9)
    .setFontWeight("bold")
    .setFontColor(PALETA.escuro)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  // --- CARDS DE KPIS ---
  const metaMinutosMes = resumo.diasUteis * JORNADA_PADRAO_MINUTOS;
  const saldoPositivo = resumo.saldoBanco >= 0;
  const corSaldo = saldoPositivo ? PALETA.verde : PALETA.vermelho;

  const cards = [
    {
      cIni: 2, cFim: 4,
      corLinha: PALETA.azul,
      titulo: "HORAS TRABALHADAS",
      valor: minutosParaHorasFormatado(resumo.totalTrabalhado) + "h",
      meta: `Meta: ${minutosParaHorasFormatado(metaMinutosMes)}h`,
      corNum: PALETA.escuro
    },
    {
      cIni: 5, cFim: 7,
      corLinha: corSaldo,
      titulo: saldoPositivo ? "CRÉDITO BANCO" : "DÉBITO BANCO",
      valor: (saldoPositivo ? "+" : "-") + minutosParaHorasFormatado(Math.abs(resumo.saldoBanco)) + "h",
      meta: saldoPositivo ? "Saldo positivo" : "Horas a pagar",
      corNum: corSaldo
    },
    {
      cIni: 8, cFim: 10,
      corLinha: PALETA.verde,
      titulo: "HORAS EXTRAS",
      valor: "+" + minutosParaHorasFormatado(resumo.totalExtras) + "h",
      meta: `${resumo.diasTrabalhados} dias trabalhados`,
      corNum: PALETA.verde
    },
    {
      cIni: 11, cFim: 13,
      corLinha: PALETA.vermelho,
      titulo: "ATRASOS / FALTAS",
      valor: "-" + minutosParaHorasFormatado(resumo.totalFaltas) + "h",
      meta: `A regularizar`,
      corNum: PALETA.vermelho
    }
  ];

  cards.forEach(c => {
    const cols = c.cFim - c.cIni + 1;

    dash.getRange(5, c.cIni, 1, cols).merge().setBackground(c.corLinha);
    dash.getRange(6, c.cIni, 4, cols).setBackground(PALETA.card);

    dash.getRange(6, c.cIni, 1, cols).merge()
      .setValue(c.titulo)
      .setFontFamily("Segoe UI")
      .setFontSize(8)
      .setFontWeight("bold")
      .setFontColor(PALETA.cinzaTexto)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("bottom");

    dash.getRange(7, c.cIni, 1, cols).merge()
      .setValue(c.valor)
      .setFontFamily("Segoe UI")
      .setFontSize(18)
      .setFontWeight("bold")
      .setFontColor(c.corNum)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    dash.getRange(8, c.cIni, 1, cols).merge()
      .setValue(c.meta)
      .setFontFamily("Segoe UI")
      .setFontSize(8)
      .setFontColor(PALETA.cinzaClaro)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("top");

    dash.getRange(5, c.cIni, 4, cols)
      .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);
  });

  // --- BARRA DE STATUS / PROGRESSO ---
  const perc = metaMinutosMes > 0 ? Math.min(100, Math.round((resumo.totalTrabalhado / metaMinutosMes) * 100)) : 0;
  const cheios = Math.round(perc / 10);
  const barraStr = "█".repeat(cheios) + "░".repeat(10 - cheios);
  const statusTexto = perc >= 100 ? "META BATIDA" : "EM ANDAMENTO";

  dash.getRange("B10:M10").merge()
    .setValue(`JORNADA DO MÊS • [ ${barraStr} ] ${perc}% • ${statusTexto} • ${resumo.diasTrabalhados}/${resumo.diasUteis} dias úteis`)
    .setFontFamily("Segoe UI")
    .setFontSize(9)
    .setFontWeight("bold")
    .setFontColor(PALETA.escuro)
    .setBackground(PALETA.card)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  // --- TABELA RESUMO DOS ÚLTIMOS DIAS ---
  dash.getRange("B12:M12").merge()
    .setValue("ÚLTIMAS ATIVIDADES REGISTRADAS")
    .setFontFamily("Segoe UI")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(PALETA.escuro)
    .setBackground("#F8FAFC")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  dash.getRange("B13").setValue("Data");
  dash.getRange("C13:D13").merge().setValue("Dia");
  dash.getRange("E13").setValue("Ent 1");
  dash.getRange("F13").setValue("Saí 1");
  dash.getRange("G13").setValue("Ent 2");
  dash.getRange("H13").setValue("Saí 2");
  dash.getRange("I13:J13").merge().setValue("Total");
  dash.getRange("K13").setValue("Extra");
  dash.getRange("L13").setValue("Falta");
  dash.getRange("M13").setValue("Status");

  dash.getRange("B13:M13")
    .setBackground("#E2E8F0")
    .setFontFamily("Segoe UI")
    .setFontSize(8)
    .setFontWeight("bold")
    .setFontColor(PALETA.escuro)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#94A3B8", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  const ultPonto = sheetPonto.getLastRow();
  let dadosPreview = [];
  if (ultPonto >= 5) {
    const totalPonto = ultPonto - 4;
    const qtd = Math.min(6, totalPonto);
    const ini = Math.max(5, ultPonto - qtd + 1);
    dadosPreview = sheetPonto.getRange(ini, 1, qtd, 9).getDisplayValues();
  }

  for (let i = 0; i < 6; i++) {
    const r = 14 + i;
    const item = dadosPreview[i] || ["--/--", "-", "-", "-", "-", "-", "-", "00:00", "00:00"];

    dash.getRange(`B${r}`).setValue(item[0]);
    dash.getRange(`C${r}:D${r}`).merge().setValue(item[1]);
    dash.getRange(`E${r}`).setValue(item[2] || "-");
    dash.getRange(`F${r}`).setValue(item[3] || "-");
    dash.getRange(`G${r}`).setValue(item[4] || "-");
    dash.getRange(`H${r}`).setValue(item[5] || "-");
    dash.getRange(`I${r}:J${r}`).merge().setValue(item[6] || "-");
    dash.getRange(`K${r}`).setValue(item[7] || "-");
    dash.getRange(`L${r}`).setValue(item[8] || "-");
    dash.getRange(`M${r}`).setValue(item[2] && item[5] ? "✓ Ok" : "Pendente");

    dash.getRange(`B${r}:M${r}`)
      .setBackground(i % 2 === 0 ? PALETA.card : "#FAFAFC")
      .setFontFamily("Segoe UI")
      .setFontSize(8)
      .setFontColor(PALETA.cinzaTexto)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }

  dash.getRange("B13:M19").setBorder(true, true, true, true, true, true, "#64748B", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  adicionarGraficosDashboard(dash, resumo);
}

function adicionarGraficosDashboard(dash, resumo) {
  const metaMinutosMes = resumo.diasUteis * JORNADA_PADRAO_MINUTOS;

  dash.getRange("O1:Q5").setValues([
    ["Categoria", "Minutos", ""],
    ["Meta", metaMinutosMes, ""],
    ["Trabalhadas", resumo.totalTrabalhado, ""],
    ["Extras", resumo.totalExtras, ""],
    ["Faltas", Math.max(0, resumo.totalFaltas), ""]
  ]);

  const graficoComparativo = dash.newChart()
    .asColumnChart()
    .addRange(dash.getRange("O1:Q5"))
    .setPosition(20, 2, 0, 0)
    .setOption('title', 'Comparativo do Mês')
    .setOption('titleTextStyle', { color: '#0F172A', fontSize: 11, bold: true })
    .setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', '#F8FAFC')
    .setOption('colors', ['#94A3B8', '#2563EB', '#16A34A', '#DC2626'])
    .setOption('hAxis', { textStyle: { color: '#475569' }, titleTextStyle: { color: '#475569' } })
    .setOption('vAxis', { textStyle: { color: '#475569' }, format: '##h' })
    .setOption('chartArea', { width: '80%', height: '70%' })
    .build();

  dash.insertChart(graficoComparativo);

  dash.getRange("T1:U4").setValues([
    ["Tipo", "Minutos"],
    ["Trabalhado", resumo.totalTrabalhado],
    ["Extras", resumo.totalExtras],
    ["Faltas", Math.max(0, resumo.totalFaltas)]
  ]);

  const graficoPizza = dash.newChart()
    .asPieChart()
    .addRange(dash.getRange("T1:U4"))
    .setPosition(20, 9, 0, 0)
    .setOption('title', 'Distribuição de Horas')
    .setOption('titleTextStyle', { color: '#0F172A', fontSize: 11, bold: true })
    .setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', '#F8FAFC')
    .setOption('colors', ['#2563EB', '#16A34A', '#DC2626'])
    .setOption('pieSliceText', 'value')
    .setOption('chartArea', { width: '80%', height: '72%' })
    .build();

  dash.insertChart(graficoPizza);
}

function atualizarDashboard() {
  recalcularETrazerDetalhes(false);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(NOME_ABA_DASHBOARD);
  if (dash) ss.setActiveSheet(dash);
}

function gerarDatasDoMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obterAbaPonto();
  const fuso = ss.getSpreadsheetTimeZone();
  
  const hoje = new Date();
  const ano = hoje.getFullYear();
  const mes = hoje.getMonth();
  
  const totalDiasMes = new Date(ano, mes + 1, 0).getDate();
  const diasDaSemanaNomes = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];

  let dadosDatas = [];
  
  for (let dia = 1; dia <= totalDiasMes; dia++) {
    let dataObj = new Date(ano, mes, dia);
    let diaSemanaNome = diasDaSemanaNomes[dataObj.getDay()];
    let dataFormatada = Utilities.formatDate(dataObj, fuso, "dd/MM/yyyy");
    dadosDatas.push([dataFormatada, diaSemanaNome]);
  }

  sheet.getRange(5, 1, dadosDatas.length, 2).setValues(dadosDatas);
  SpreadsheetApp.getUi().alert(`📅 Calendário gerado com sucesso! (${totalDiasMes} dias inseridos)`);
  recalcularETrazerDetalhes(false);
}

function registrarPonto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obterAbaPonto();
  const fuso = ss.getSpreadsheetTimeZone();
  const hoje = new Date();

  const dataHojeStr = Utilities.formatDate(hoje, fuso, "dd/MM/yyyy");
  const horaAtualStr = Utilities.formatDate(hoje, fuso, "HH:mm");
  const diaSemanaNome = Utilities.formatDate(hoje, fuso, "EEEE");

  const dados = sheet.getDataRange().getDisplayValues();
  let linhaEncontrada = -1;

  for (let i = 4; i < dados.length; i++) {
    if (dados[i][0]) {
      if (dados[i][0].trim() === dataHojeStr) {
        linhaEncontrada = i + 1;
        break;
      }
    }
  }

  if (linhaEncontrada === -1) {
    const ultimaLinha = Math.max(5, sheet.getLastRow() + 1);
    sheet.getRange(ultimaLinha, 1, 1, 2).setValues([[dataHojeStr, diaSemanaNome]]);
    linhaEncontrada = ultimaLinha;
  }

  const colunasHorario = [3, 4, 5, 6];
  let pontoRegistrado = false;

  for (let col of colunasHorario) {
    const valorCelula = sheet.getRange(linhaEncontrada, col).getValue();
    if (!valorCelula || valorCelula === "") {
      sheet.getRange(linhaEncontrada, col).setValue(horaAtualStr);
      pontoRegistrado = true;
      break;
    }
  }

  if (pontoRegistrado) {
    recalcularETrazerDetalhes(false);
    SpreadsheetApp.getUi().alert(`✅ Ponto registrado às ${horaAtualStr} com sucesso!`);
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ Todos os 4 horários do dia de hoje já estão preenchidos.");
  }
}

function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() === NOME_ABA_DASHBOARD) return;

  const col = range.getColumn();
  if (range.getRow() >= 5 && col >= 3 && col <= 6) {
    recalcularETrazerDetalhes(false);
  }
}

function enviarRelatorioPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obterAbaPonto();
  const emailDestino = Session.getActiveUser().getEmail();

  const url = ss.getUrl().replace(/edit$/, '') + 'export?';
  const parametrosPDF = {
    exportFormat: 'pdf',
    format: 'pdf',
    size: 'A4',
    portrait: 'false',
    fitw: 'true',
    gridlines: 'true',
    sheetnames: 'false',
    gid: sheet.getSheetId()
  };

  let queryUrl = [];
  for (let key in parametrosPDF) {
    queryUrl.push(key + '=' + parametrosPDF[key]);
  }
  const urlFinal = url + queryUrl.join('&');

  const PDFBlob = UrlFetchApp.fetch(urlFinal, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  }).getBlob().setName(`Espelho_de_Ponto_${sheet.getName()}.pdf`);

  MailApp.sendEmail(
    emailDestino,
    `Espelho de Ponto - ${sheet.getName()}`,
    `Segue em anexo o relatório em PDF com o espelho de ponto atualizado.`,
    { attachments: [PDFBlob] }
  );

  SpreadsheetApp.getUi().alert(`📧 Relatório enviado com sucesso para: ${emailDestino}`);
}
