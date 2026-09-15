/**
 * SISTEMA INTEGRADO DE CONTROLE DE PONTO E BANCO DE HORAS
 * Jornada Padrão: Segunda a Sexta - 09:00 por dia (540 minutos)
 */

//const JORNADA_PADRAO_MINUTOS = 9 * 60; // 540 minutos / 9 horas por dia útil

const JORNADA_PADRAO_MINUTOS = (8 * 60)+48; // 528 minutos / 8h48 por dia útil

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⏰ Registro de Ponto')
    .addItem('Bater Ponto Agora', 'registrarPonto')
    .addSeparator()
    .addItem('Gerar Calendário do Mês Atual', 'gerarDatasDoMes')
    .addItem('Recalcular Tudo e Exibir Resumo', 'recalcularETrazerDetalhes')
    .addItem('Gerar e Enviar PDF por E-mail', 'enviarRelatorioPDF')
    .addToUi();
}

/**
 * Converte qualquer texto de horário HH:MM para minutos inteiros
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
 * Recalcula todas as colunas mantendo e reforçando as bordas da planilha
 */
function recalcularETrazerDetalhes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 5) return;

  const numLinhas = ultimaLinha - 4;
  
  // Lê os valores exibidos na tela para evitar erros de formatação
  const dados = sheet.getRange(5, 1, numLinhas, 6).getDisplayValues();

  let totalMinutosTrabalhados = 0;
  let totalMinutosExtras = 0;
  let totalMinutosFaltas = 0;
  let saldoBancoMinutos = 0;
  let diasTrabalhadosCount = 0;
  let diasUteisCount = 0;

  let matrizResultados = [];
  let coresExtras = [];
  let coresFaltas = [];

  for (let i = 0; i < dados.length; i++) {
    const dataStr = dados[i][0];
    const diaSemana = String(dados[i][1] || '').toLowerCase().trim();
    
    const ent1 = textoHoraParaMinutos(dados[i][2]); // Entrada 1 (C)
    const sai1 = textoHoraParaMinutos(dados[i][3]); // Saída 1 (D)
    const ent2 = textoHoraParaMinutos(dados[i][4]); // Entrada 2 (E)
    const sai2 = textoHoraParaMinutos(dados[i][5]); // Saída 2 (F)

    const ehFimDeSemana = diaSemana.includes("sábado") || diaSemana.includes("sabado") || diaSemana.includes("domingo");
    if (!ehFimDeSemana && dataStr !== "") diasUteisCount++;

    // Requer pelo menos Entrada 1 e Saída Final para calcular
    if (ent1 === null || sai2 === null) {
      matrizResultados.push(["", "", ""]);
      coresExtras.push(["#ffffff"]);
      coresFaltas.push(["#ffffff"]);
      continue;
    }

    diasTrabalhadosCount++;

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

    totalMinutosTrabalhados += minTrabalhados;
    saldoBancoMinutos += saldoDiaMin;

    let hExtraStr = "00:00";
    let faltaStr = "00:00";
    let corH = "#ffffff";
    let corI = "#ffffff";

    if (saldoDiaMin > 0) {
      hExtraStr = minutosParaHorasFormatado(saldoDiaMin);
      totalMinutosExtras += saldoDiaMin;
      corH = "#d4edda"; // Verde Claro
    } else if (saldoDiaMin < 0) {
      faltaStr = minutosParaHorasFormatado(Math.abs(saldoDiaMin));
      totalMinutosFaltas += Math.abs(saldoDiaMin);
      corI = "#f8d7da"; // Vermelho Claro
    }

    const totalTrabStr = minutosParaHorasFormatado(minTrabalhados);
    matrizResultados.push([totalTrabStr, hExtraStr, faltaStr]);
    coresExtras.push([corH]);
    coresFaltas.push([corI]);
  }

  // Intervalo das colunas de resultado (G, H, I)
  const intervaloResultados = sheet.getRange(5, 7, numLinhas, 3);
  
  // Limpa apenas o texto mantendo os estilos anteriores
  intervaloResultados.clearContent();
  intervaloResultados.setNumberFormat("@");

  // Insere os valores calculados
  intervaloResultados.setValues(matrizResultados);

  // Aplica as cores de fundo
  sheet.getRange(5, 8, coresExtras.length, 1).setBackgrounds(coresExtras);
  sheet.getRange(5, 9, coresFaltas.length, 1).setBackgrounds(coresFaltas);

  // Aplica as bordas pretas finas em toda a tabela (Colunas A até I)
  const intervaloTabelaCompleta = sheet.getRange(5, 1, numLinhas, 9);
  intervaloTabelaCompleta.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Exibe o resumo final
  const saldoFinalStr = minutosParaHorasFormatado(saldoBancoMinutos, true);
  const statusBanco = saldoBancoMinutos >= 0 ? "🟢 POSITIVO (Crédito)" : "🔴 NEGATIVO (Débito)";

  const mensagemDetalhes = 
    `📊 RESUMO DETALHADO DO PONTO\n` +
    `--------------------------------------------------\n` +
    `📅 Dias Úteis no Mês: ${diasUteisCount} dias\n` +
    `✅ Dias Efetivamente Trabalhados: ${diasTrabalhadosCount} dias\n\n` +
    `⏱️ Total de Horas Trabalhadas: ${minutosParaHorasFormatado(totalMinutosTrabalhados)} h\n` +
    `➕ Total de Horas Extras: ${minutosParaHorasFormatado(totalMinutosExtras)} h\n` +
    `➖ Total de Horas Devidas / Faltas: ${minutosParaHorasFormatado(totalMinutosFaltas)} h\n` +
    `--------------------------------------------------\n` +
    `🏆 SALDO DO BANCO DE HORAS: ${saldoFinalStr} h\n` +
    `📌 Status: ${statusBanco}\n` +
    `--------------------------------------------------\n` +
    `*(Jornada padrão: 09:00h por dia útil)*`;

  SpreadsheetApp.getUi().alert(mensagemDetalhes);
}

function gerarDatasDoMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
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
  recalcularETrazerDetalhes();
}

function registrarPonto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const fuso = ss.getSpreadsheetTimeZone();
  const hoje = new Date();
  
  const dataHojeStr = Utilities.formatDate(hoje, fuso, "dd/MM/yyyy");
  const horaAtualStr = Utilities.formatDate(hoje, fuso, "HH:mm");

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
    gerarDatasDoMes();
    registrarPonto();
    return;
  }

  const colunasHorario = [3, 4, 5, 6]; 
  let pontoRegistrado = false;

  for (let col of colunasHorario) {
    let valorCelula = sheet.getRange(linhaEncontrada, col).getValue();
    if (!valorCelula || valorCelula === "") {
      sheet.getRange(linhaEncontrada, col).setValue(horaAtualStr);
      pontoRegistrado = true;
      break;
    }
  }

  if (pontoRegistrado) {
    recalcularETrazerDetalhes();
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ Todos os 4 horários do dia de hoje já estão preenchidos.");
  }
}

function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const col = range.getColumn();
  if (range.getRow() >= 5 && col >= 3 && col <= 6) {
    recalcularETrazerDetalhes();
  }
}

function enviarRelatorioPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
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
