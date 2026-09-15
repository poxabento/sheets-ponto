const JORNADA_PADRAO_MINUTOS = (8 * 60) + 48;

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

function textoHoraParaMinutos(str) {
  if (!str) return null;
  const textoLimpo = String(str).trim();
  if (textoLimpo === "" || textoLimpo === "-") return null;
  const partes = textoLimpo.split(':');
  if (partes.length >= 2) {
    const h = parseInt(partes[0], 10);
    const m = parseInt(partes[1], 10);
    if (!isNaN(h) && !isNaN(m)) return (h * 60) + m;
  }
  return null;
}

function minutosParaHorasFormatado(minutos, comSinal = false) {
  const sinal = minutos < 0 ? "-" : (comSinal && minutos > 0 ? "+" : "");
  const minsAbs = Math.abs(minutos);
  const hrs = Math.floor(minsAbs / 60);
  const mins = minsAbs % 60;
  return `${sinal}${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
}

function recalcularETrazerDetalhes(silencioso = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 5) return;

  const numLinhas = ultimaLinha - 4;
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
    
    const ent1 = textoHoraParaMinutos(dados[i][2]);
    const sai1 = textoHoraParaMinutos(dados[i][3]);
    const ent2 = textoHoraParaMinutos(dados[i][4]);
    const sai2 = textoHoraParaMinutos(dados[i][5]);

    const ehFimDeSemana = diaSemana.includes("sábado") || diaSemana.includes("sabado") || diaSemana.includes("domingo");
    if (!ehFimDeSemana && dataStr !== "") diasUteisCount++;

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
      corH = "#d4edda";
    } else if (saldoDiaMin < 0) {
      faltaStr = minutosParaHorasFormatado(Math.abs(saldoDiaMin));
      totalMinutosFaltas += Math.abs(saldoDiaMin);
      corI = "#f8d7da";
    }

    matrizResultados.push([minutosParaHorasFormatado(minTrabalhados), hExtraStr, faltaStr]);
    coresExtras.push([corH]);
    coresFaltas.push([corI]);
  }

  const intervaloResultados = sheet.getRange(5, 7, numLinhas, 3);
  intervaloResultados.clearContent();
  intervaloResultados.setNumberFormat("@");
  intervaloResultados.setValues(matrizResultados);

  sheet.getRange(5, 8, coresExtras.length, 1).setBackgrounds(coresExtras);
  sheet.getRange(5, 9, coresFaltas.length, 1).setBackgrounds(coresFaltas);

  const intervaloTabelaCompleta = sheet.getRange(5, 1, numLinhas, 9);
  intervaloTabelaCompleta.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  if (!silencioso) {
    const saldoFinalStr = minutosParaHorasFormatado(saldoBancoMinutos, true);
    const statusBanco = saldoBancoMinutos >= 0 ? "🟢 POSITIVO (Crédito)" : "🔴 NEGATIVO (Débito)";
    const msg = 
      `📊 RESUMO DETALHADO DO PONTO\n` +
      `--------------------------------------------------\n` +
      `📅 Dias Úteis no Mês: ${diasUteisCount} dias\n` +
      `✅ Dias Trabalhados: ${diasTrabalhadosCount} dias\n\n` +
      `⏱️ Total Trabalhado: ${minutosParaHorasFormatado(totalMinutosTrabalhados)} h\n` +
      `➕ Horas Extras: ${minutosParaHorasFormatado(totalMinutosExtras)} h\n` +
      `➖ Faltas/Devidas: ${minutosParaHorasFormatado(totalMinutosFaltas)} h\n` +
      `--------------------------------------------------\n` +
      `🏆 BANCO DE HORAS: ${saldoFinalStr} h\n` +
      `📌 Status: ${statusBanco}`;
    SpreadsheetApp.getUi().alert(msg);
  }
}

function gerarDatasDoMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const fuso = ss.getSpreadsheetTimeZone();
  const hoje = new Date();
  const ano = hoje.getFullYear();
  const mes = hoje.getMonth();
  
  const totalDiasMes = new Date(ano, mes + 1, 0).getDate();
  const diasDaSemana = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];

  let dadosDatas = [];
  for (let dia = 1; dia <= totalDiasMes; dia++) {
    let dataObj = new Date(ano, mes, dia);
    dadosDatas.push([Utilities.formatDate(dataObj, fuso, "dd/MM/yyyy"), diasDaSemana[dataObj.getDay()]]);
  }

  sheet.getRange(5, 1, dadosDatas.length, 2).setValues(dadosDatas);
  SpreadsheetApp.getUi().alert(`📅 Calendário gerado! (${totalDiasMes} dias inseridos)`);
  recalcularETrazerDetalhes(true);
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
    if (dados[i][0] && dados[i][0].trim() === dataHojeStr) {
      linhaEncontrada = i + 1;
      break;
    }
  }

  if (linhaEncontrada === -1) {
    gerarDatasDoMes();
    registrarPonto();
    return;
  }

  const colunas = [3, 4, 5, 6];
  let pontoRegistrado = false;
  for (let col of colunas) {
    let val = sheet.getRange(linhaEncontrada, col).getValue();
    if (!val || val === "") {
      sheet.getRange(linhaEncontrada, col).setValue(horaAtualStr);
      pontoRegistrado = true;
      break;
    }
  }

  if (pontoRegistrado) {
    recalcularETrazerDetalhes(true);
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ Todos os 4 horários de hoje já estão preenchidos.");
  }
}

function enviarRelatorioPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const emailDestino = Session.getActiveUser().getEmail();

  const url = ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf&size=A4&portrait=false&fitw=true&gridlines=true&sheetnames=false&gid=' + sheet.getSheetId();

  const PDFBlob = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  }).getBlob().setName(`Espelho_Ponto_${sheet.getName()}.pdf`);

  MailApp.sendEmail(
    emailDestino,
    `Espelho de Ponto - ${sheet.getName()}`,
    `Segue em anexo o relatório em PDF atualizado.`,
    { attachments: [PDFBlob] }
  );

  SpreadsheetApp.getUi().alert(`📧 Relatório enviado para: ${emailDestino}`);
}
