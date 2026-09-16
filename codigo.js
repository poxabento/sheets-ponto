/**
 * SISTEMA INTEGRADO DE CONTROLE DE PONTO E BANCO DE HORAS COM DASHBOARD
 * Jornada Padrão: 8h48min por dia útil (528 minutos)
 *
 * v3 — Melhorias sobre a v2:
 *  - Processa TODAS as abas de mês da planilha (não só a ativa), aplicando o
 *    mesmo cálculo e formatação em cada uma.
 *  - Aba "Feriados" (criada automaticamente) — datas nela contam como não-úteis,
 *    igual fim de semana, tanto no cálculo quanto na meta do dia.
 *  - Histórico mensal: gráfico comparando o saldo do banco de horas de cada
 *    aba/mês processada.
 *  - Alerta automático por e-mail quando o saldo do banco fica abaixo de um
 *    limite (padrão: -3h), no máximo uma vez por dia.
 *  - Seta de tendência no card do banco de horas (▲/▼) comparando o mês atual
 *    com a média dos meses anteriores.
 *
 * v2 já trazia: layout construído uma única vez (performance), gráfico de
 * tendência do saldo diário, indicadores de projeção/médias.
 */

const JORNADA_PADRAO_MINUTOS = (8 * 60) + 48; // 528 minutos / 8h48 por dia útil
const NOME_ABA_DASHBOARD = "📊 Painel Geral";
const NOME_ABA_FERIADOS = "Feriados";
const MARCADOR_LAYOUT = "layoutV3";
const MAX_DIAS_TENDENCIA = 31;
const MAX_MESES_HISTORICO = 12;
const LIMITE_ALERTA_SALDO_MIN = -180; // -3h — abaixo disso, dispara e-mail
const PROP_ULTIMO_ALERTA = "ULTIMO_ALERTA_SALDO";

const PALETA = {
  canvas:     "#F8FAFC",
  card:       "#FFFFFF",
  cardAlt:    "#F5F7FF",
  borda:      "#E2E8F0",
  escuro:     "#0F172A",
  cinzaTexto: "#475569",
  cinzaClaro: "#94A3B8",
  azul:       "#2563EB",
  indigo:     "#4F46E5",
  verde:      "#16A34A",
  vermelho:   "#DC2626"
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⏰ Registro de Ponto')
    .addItem('Bater Ponto Agora', 'registrarPonto')
    .addItem('Abrir / Atualizar Dashboard', 'atualizarDashboard')
    .addSeparator()
    .addItem('Gerar Calendário do Mês Atual', 'gerarDatasDoMes')
    .addItem('Recalcular Tudo e Exibir Resumo', 'recalcularETrazerDetalhes')
    .addSeparator()
    .addItem('🔄 Reconstruir Layout do Zero', 'forcarReconstrucaoLayout')
    .addItem('Gerar e Enviar PDF por E-mail', 'enviarRelatorioPDF')
    .addToUi();
}

/** Converte horário HH:MM para minutos inteiros */
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

/** Formata minutos em string de hora HH:MM */
function minutosParaHorasFormatado(minutos, comSinal = false) {
  const sinal = minutos < 0 ? "-" : (comSinal && minutos > 0 ? "+" : "");
  const minsAbs = Math.abs(Math.round(minutos));
  const hrs = Math.floor(minsAbs / 60);
  const mins = minsAbs % 60;
  return `${sinal}${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
}

/**
 * Lê (e cria, se não existir) a aba "Feriados" com uma lista inicial de
 * feriados nacionais. O usuário pode editar/completar essa aba livremente —
 * qualquer data lá dentro passa a contar como não-útil no cálculo.
 */
function obterFeriados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NOME_ABA_FERIADOS);

  if (!sheet) {
    sheet = ss.insertSheet(NOME_ABA_FERIADOS);
    sheet.getRange("A1:B1").setValues([["Data (dd/mm/aaaa)", "Descrição"]]).setFontWeight("bold");
    sheet.getRange("A2:B14").setValues([
      ["01/01/2026", "Confraternização Universal"],
      ["16/02/2026", "Carnaval"],
      ["17/02/2026", "Carnaval"],
      ["03/04/2026", "Sexta-feira Santa"],
      ["21/04/2026", "Tiradentes"],
      ["01/05/2026", "Dia do Trabalho"],
      ["04/06/2026", "Corpus Christi"],
      ["07/09/2026", "Independência do Brasil"],
      ["12/10/2026", "Nossa Senhora Aparecida"],
      ["02/11/2026", "Finados"],
      ["15/11/2026", "Proclamação da República"],
      ["20/11/2026", "Consciência Negra"],
      ["25/12/2026", "Natal"]
    ]);
    sheet.autoResizeColumns(1, 2);
  }

  const ultimaLinha = sheet.getLastRow();
  const feriados = new Set();
  if (ultimaLinha >= 2) {
    const valores = sheet.getRange(2, 1, ultimaLinha - 1, 1).getDisplayValues();
    valores.forEach(l => { if (l[0]) feriados.add(l[0].trim()); });
  }
  return feriados;
}

/**
 * Localiza com segurança a aba de dados do ponto ATIVA (usada como "mês
 * atual" para os cards principais e projeções).
 */
function obterAbaPonto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abasIgnoradas = [NOME_ABA_DASHBOARD, NOME_ABA_FERIADOS];
  const sheets = ss.getSheets();

  const abaAtiva = ss.getActiveSheet();
  if (abaAtiva && !abasIgnoradas.includes(abaAtiva.getName())) return abaAtiva;

  for (let s of sheets) {
    if (!abasIgnoradas.includes(s.getName())) return s;
  }
  return sheets[0] || ss.insertSheet("Ponto");
}

/**
 * Calcula o resumo de UMA aba de mês: percorre as linhas de ponto, escreve
 * de volta as colunas de resultado (G/H/I) com as cores, e devolve o objeto
 * de resumo (incluindo a série para o gráfico de tendência diária).
 */
function calcularResumoAba(sheet, feriados) {
  const ultimaLinha = sheet.getLastRow();
  let resumo = {
    nome: sheet.getName(),
    diasUteis: 0,
    diasTrabalhados: 0,
    totalTrabalhado: 0,
    totalExtras: 0,
    totalFaltas: 0,
    saldoBanco: 0,
    serieTendencia: []
  };

  if (ultimaLinha < 5) return resumo;

  const numLinhas = ultimaLinha - 4;
  const dados = sheet.getRange(5, 1, numLinhas, 6).getDisplayValues();

  let matrizResultados = [];
  let coresExtras = [];
  let coresFaltas = [];
  let saldoAcumulado = 0;

  for (let i = 0; i < dados.length; i++) {
    const dataStr = dados[i][0];
    const diaSemana = String(dados[i][1] || '').toLowerCase().trim();

    const ent1 = textoHoraParaMinutos(dados[i][2]);
    const sai1 = textoHoraParaMinutos(dados[i][3]);
    const ent2 = textoHoraParaMinutos(dados[i][4]);
    const sai2 = textoHoraParaMinutos(dados[i][5]);

    const ehFimDeSemana = diaSemana.includes("sábado") || diaSemana.includes("sabado") || diaSemana.includes("domingo");
    const ehFeriado = dataStr !== "" && feriados.has(dataStr.trim());
    const ehDiaNaoUtil = ehFimDeSemana || ehFeriado;

    if (!ehDiaNaoUtil && dataStr !== "") resumo.diasUteis++;

    if (ent1 === null || sai2 === null) {
      matrizResultados.push(["", "", ""]);
      coresExtras.push(["#ffffff"]);
      coresFaltas.push(["#ffffff"]);
      continue;
    }

    resumo.diasTrabalhados++;
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

    const metaDiaMin = ehDiaNaoUtil ? 0 : JORNADA_PADRAO_MINUTOS;
    const saldoDiaMin = minTrabalhados - metaDiaMin;

    resumo.totalTrabalhado += minTrabalhados;
    resumo.saldoBanco += saldoDiaMin;
    saldoAcumulado += saldoDiaMin;

    const labelDia = dataStr ? dataStr.substring(0, 5) : `#${i + 1}`;
    resumo.serieTendencia.push({ label: labelDia, saldoAcumulado: saldoAcumulado });

    let hExtraStr = "00:00";
    let faltaStr = "00:00";
    let corH = "#ffffff";
    let corI = "#ffffff";

    if (saldoDiaMin > 0) {
      hExtraStr = minutosParaHorasFormatado(saldoDiaMin);
      resumo.totalExtras += saldoDiaMin;
      corH = "#dcfce7";
    } else if (saldoDiaMin < 0) {
      faltaStr = minutosParaHorasFormatado(Math.abs(saldoDiaMin));
      resumo.totalFaltas += Math.abs(saldoDiaMin);
      corI = ehFeriado ? "#fee2e2" : "#fee2e2";
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

  sheet.getRange(5, 1, numLinhas, 9)
    .setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  return resumo;
}

/**
 * Processa TODAS as abas de mês da planilha (ignorando Dashboard e
 * Feriados), devolvendo o resumo de cada uma e identificando qual é a
 * "atual" (a aba ativa no momento da chamada).
 */
function processarTodasAsAbas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feriados = obterFeriados();
  const abasIgnoradas = [NOME_ABA_DASHBOARD, NOME_ABA_FERIADOS];
  const abaAtiva = obterAbaPonto();

  let historico = [];
  let resumoAtual = null;

  ss.getSheets().forEach(sheet => {
    if (abasIgnoradas.includes(sheet.getName())) return;
    const resumo = calcularResumoAba(sheet, feriados);
    historico.push(resumo);
    if (sheet.getSheetId() === abaAtiva.getSheetId()) resumoAtual = resumo;
  });

  if (!resumoAtual && historico.length > 0) resumoAtual = historico[historico.length - 1];

  return { resumoAtual, historico, abaAtiva };
}

/**
 * Recalcula todas as abas, monta projeções para o mês atual, verifica o
 * alerta de saldo negativo e atualiza a Dashboard.
 */
function recalcularETrazerDetalhes(exibirAlerta = true) {
  const { resumoAtual, historico, abaAtiva } = processarTodasAsAbas();

  // --- Projeções e médias (mês atual) ---
  resumoAtual.diasUteisRestantes = Math.max(0, resumoAtual.diasUteis - resumoAtual.diasTrabalhados);
  resumoAtual.mediaSaldoPorDiaTrabalhado = resumoAtual.diasTrabalhados > 0 ? (resumoAtual.saldoBanco / resumoAtual.diasTrabalhados) : 0;
  resumoAtual.mediaTrabalhadaDiaria = resumoAtual.diasTrabalhados > 0 ? (resumoAtual.totalTrabalhado / resumoAtual.diasTrabalhados) : 0;
  resumoAtual.projecaoFechamento = resumoAtual.saldoBanco + (resumoAtual.mediaSaldoPorDiaTrabalhado * resumoAtual.diasUteisRestantes);
  resumoAtual.ajusteDiarioNecessario = resumoAtual.diasUteisRestantes > 0 ? (-resumoAtual.saldoBanco / resumoAtual.diasUteisRestantes) : null;

  // --- Tendência vs. meses anteriores ---
  const outrosMeses = historico.filter(h => h.nome !== resumoAtual.nome);
  if (outrosMeses.length > 0) {
    const mediaHistorica = outrosMeses.reduce((soma, h) => soma + h.saldoBanco, 0) / outrosMeses.length;
    resumoAtual.tendenciaVsHistorico = resumoAtual.saldoBanco > mediaHistorica ? "melhor" : (resumoAtual.saldoBanco < mediaHistorica ? "pior" : "igual");
  } else {
    resumoAtual.tendenciaVsHistorico = null;
  }

  verificarEEnviarAlerta(resumoAtual);
  renderizarDashboard(resumoAtual, historico, abaAtiva);

  if (exibirAlerta) {
    const saldoFinalStr = minutosParaHorasFormatado(resumoAtual.saldoBanco, true);
    const statusBanco = resumoAtual.saldoBanco >= 0 ? "🟢 POSITIVO" : "🔴 NEGATIVO";
    const msg = `📊 RESUMO ATUALIZADO — ${resumoAtual.nome}\n\n` +
      `📅 Dias Trabalhados: ${resumoAtual.diasTrabalhados} / ${resumoAtual.diasUteis}\n` +
      `⏱️ Horas Trabalhadas: ${minutosParaHorasFormatado(resumoAtual.totalTrabalhado)} h\n` +
      `➕ Extras: ${minutosParaHorasFormatado(resumoAtual.totalExtras)} h\n` +
      `➖ Faltas: ${minutosParaHorasFormatado(resumoAtual.totalFaltas)} h\n` +
      `🏆 Saldo do Banco: ${saldoFinalStr} h (${statusBanco})\n` +
      `🔮 Projeção de Fechamento: ${minutosParaHorasFormatado(resumoAtual.projecaoFechamento, true)} h\n` +
      `📚 Abas processadas: ${historico.map(h => h.nome).join(", ")}`;
    SpreadsheetApp.getUi().alert(msg);
  }

  return resumoAtual;
}

/** Dispara um e-mail de alerta se o saldo do banco estiver abaixo do limite (no máx. 1x/dia). */
function verificarEEnviarAlerta(resumo) {
  if (resumo.saldoBanco >= LIMITE_ALERTA_SALDO_MIN) return;

  const props = PropertiesService.getScriptProperties();
  const fuso = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const hojeStr = Utilities.formatDate(new Date(), fuso, "yyyy-MM-dd");
  if (props.getProperty(PROP_ULTIMO_ALERTA) === hojeStr) return;

  const email = Session.getActiveUser().getEmail();
  if (!email) return;

  MailApp.sendEmail(
    email,
    `⚠️ Banco de horas negativo — ${resumo.nome}`,
    `Seu banco de horas em "${resumo.nome}" está em ${minutosParaHorasFormatado(resumo.saldoBanco, true)}h, ` +
    `abaixo do limite de alerta (${minutosParaHorasFormatado(LIMITE_ALERTA_SALDO_MIN)}h).\n\n` +
    `Acesse a planilha para ver os detalhes e reorganizar seus próximos dias.`
  );
  props.setProperty(PROP_ULTIMO_ALERTA, hojeStr);
}

/** Força a reconstrução completa do layout na próxima atualização. */
function forcarReconstrucaoLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(NOME_ABA_DASHBOARD);
  if (dash) dash.getRange("A1").setNote("");
  atualizarDashboard();
}

function atualizarDashboard() {
  recalcularETrazerDetalhes(false);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(NOME_ABA_DASHBOARD);
  if (dash) ss.setActiveSheet(dash);
}

/**
 * Ponto de entrada da renderização: garante que o layout estático exista
 * (constrói só na primeira vez ou quando forçado) e então apenas atualiza
 * os valores dinâmicos.
 */
function renderizarDashboard(resumoAtual, historico, sheetAtual) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName(NOME_ABA_DASHBOARD);
  const layoutPresente = dash && dash.getRange("A1").getNote() === MARCADOR_LAYOUT;

  if (!dash) dash = ss.insertSheet(NOME_ABA_DASHBOARD, 0);
  if (!layoutPresente) construirLayoutBase(dash);

  atualizarValoresDashboard(dash, resumoAtual, historico, sheetAtual);
}

/** Constrói toda a parte estática da dashboard. Roda só uma vez. */
function construirLayoutBase(dash) {
  try { dash.getDataRange().breakApart(); } catch (err) {}
  dash.clear();
  dash.getCharts().forEach(c => dash.removeChart(c));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(dash);
  ss.moveActiveSheet(1);

  dash.setHiddenGridlines(true);

  const TOTAL_LINHAS = 70;
  const TOTAL_COLUNAS = 27; // até AA — O em diante reservadas para dados dos gráficos
  if (dash.getMaxColumns() < TOTAL_COLUNAS) dash.insertColumnsAfter(dash.getMaxColumns(), TOTAL_COLUNAS - dash.getMaxColumns());
  if (dash.getMaxRows() < TOTAL_LINHAS) dash.insertRowsAfter(dash.getMaxRows(), TOTAL_LINHAS - dash.getMaxRows());

  dash.getRange(1, 1, TOTAL_LINHAS, TOTAL_COLUNAS).setBackground(PALETA.canvas);

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
    .setBackground(PALETA.escuro).setFontColor("#FFFFFF").setFontFamily("Segoe UI")
    .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.escuro, SpreadsheetApp.BorderStyle.SOLID)
    .setValue("CONTROLE DE PONTO & BANCO DE HORAS");

  dash.getRange("B3:M3").merge()
    .setBackground("#F8FAFC").setFontFamily("Segoe UI").setFontSize(9).setFontColor(PALETA.cinzaTexto)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange("B4:M4").merge()
    .setValue("RESUMO GERAL DO MÊS")
    .setBackground("#E2E8F0").setFontFamily("Segoe UI").setFontSize(9).setFontWeight("bold")
    .setFontColor(PALETA.escuro).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  // --- CARDS PRINCIPAIS ---
  construirEstruturaCard(dash, [
    { cIni: 2, cFim: 4,  titulo: "HORAS TRABALHADAS" },
    { cIni: 5, cFim: 7,  titulo: "BANCO DE HORAS" },
    { cIni: 8, cFim: 10, titulo: "HORAS EXTRAS" },
    { cIni: 11, cFim: 13, titulo: "ATRASOS / FALTAS" }
  ], 5, PALETA.card);

  // --- BARRA DE PROGRESSO ---
  dash.getRange("B10:M10").merge()
    .setFontFamily("Segoe UI").setFontSize(9).setFontWeight("bold").setFontColor(PALETA.escuro)
    .setBackground(PALETA.card).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  // --- LABEL PROJEÇÕES ---
  dash.getRange("B12:M12").merge()
    .setValue("PROJEÇÕES E MÉDIAS")
    .setBackground("#E2E8F0").setFontFamily("Segoe UI").setFontSize(9).setFontWeight("bold")
    .setFontColor(PALETA.escuro).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  // --- MINI CARDS DE PROJEÇÃO ---
  construirEstruturaCard(dash, [
    { cIni: 2, cFim: 4,  titulo: "DIAS ÚTEIS RESTANTES" },
    { cIni: 5, cFim: 7,  titulo: "MÉDIA DIÁRIA TRABALHADA" },
    { cIni: 8, cFim: 10, titulo: "AJUSTE NECESSÁRIO/DIA" },
    { cIni: 11, cFim: 13, titulo: "PROJEÇÃO DE FECHAMENTO" }
  ], 13, PALETA.cardAlt);

  // --- LABEL TENDÊNCIA DIÁRIA ---
  dash.getRange("B18:M18").merge()
    .setValue("TENDÊNCIA DO SALDO DIÁRIO (MÊS ATUAL)")
    .setBackground("#E2E8F0").setFontFamily("Segoe UI").setFontSize(9).setFontWeight("bold")
    .setFontColor(PALETA.escuro).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange("B19:M34").setBackground(PALETA.card)
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  // --- LABEL HISTÓRICO MENSAL ---
  dash.getRange("B36:M36").merge()
    .setValue("HISTÓRICO — SALDO DO BANCO POR MÊS (TODAS AS ABAS)")
    .setBackground("#E2E8F0").setFontFamily("Segoe UI").setFontSize(9).setFontWeight("bold")
    .setFontColor(PALETA.escuro).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange("B37:M52").setBackground(PALETA.card)
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  // --- LABEL COMPARATIVO / DISTRIBUIÇÃO ---
  dash.getRange("B54:M54").merge()
    .setValue("COMPARATIVO E DISTRIBUIÇÃO (MÊS ATUAL)")
    .setBackground("#E2E8F0").setFontFamily("Segoe UI").setFontSize(9).setFontWeight("bold")
    .setFontColor(PALETA.escuro).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange("B55:M69").setBackground(PALETA.card)
    .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);

  // --- Faixas de dados fixas para os gráficos (ocultas) ---
  dash.getRange("O1").setValue("Dia");
  dash.getRange("P1").setValue("Saldo Acumulado (h)");
  dash.getRange("R1:S5").setValues([["Categoria", "Horas"], ["Meta", 0], ["Trabalhadas", 0], ["Extras", 0], ["Faltas", 0]]);
  dash.getRange("U1:V4").setValues([["Tipo", "Horas"], ["Trabalhado", 0], ["Extras", 0], ["Faltas", 0]]);
  dash.getRange("X1:Y1").setValues([["Mês", "Saldo (h)"]]);
  dash.hideColumns(15, 11); // O até Y

  // Gráfico de tendência diária (linha) — mês atual
  const graficoTendencia = dash.newChart().asLineChart()
    .addRange(dash.getRange(1, 15, MAX_DIAS_TENDENCIA + 1, 2))
    .setPosition(19, 2, 0, 0)
    .setOption('title', 'Saldo Acumulado do Banco de Horas — Dia a Dia')
    .setOption('titleTextStyle', { color: PALETA.escuro, fontSize: 11, bold: true })
    .setOption('legend', { position: 'none' })
    .setOption('backgroundColor', PALETA.card)
    .setOption('colors', [PALETA.indigo])
    .setOption('curveType', 'function')
    .setOption('pointSize', 4)
    .setOption('hAxis', { textStyle: { color: PALETA.cinzaTexto, fontSize: 8 } })
    .setOption('vAxis', { textStyle: { color: PALETA.cinzaTexto }, format: '#,##0.0"h"', gridlines: { color: PALETA.borda } })
    .setOption('chartArea', { width: '85%', height: '70%' })
    .setOption('width', 760).setOption('height', 290)
    .build();
  dash.insertChart(graficoTendencia);

  // Gráfico de histórico mensal (barras) — todas as abas processadas
  const graficoHistorico = dash.newChart().asColumnChart()
    .addRange(dash.getRange(1, 24, MAX_MESES_HISTORICO + 1, 2))
    .setPosition(37, 2, 0, 0)
    .setOption('title', 'Saldo do Banco de Horas por Mês')
    .setOption('titleTextStyle', { color: PALETA.escuro, fontSize: 11, bold: true })
    .setOption('legend', { position: 'none' })
    .setOption('backgroundColor', PALETA.card)
    .setOption('series', { 0: { color: PALETA.indigo } })
    .setOption('hAxis', { textStyle: { color: PALETA.cinzaTexto, fontSize: 9 } })
    .setOption('vAxis', { textStyle: { color: PALETA.cinzaTexto }, format: '#,##0.0"h"', gridlines: { color: PALETA.borda } })
    .setOption('chartArea', { width: '85%', height: '68%' })
    .setOption('width', 760).setOption('height', 280)
    .build();
  dash.insertChart(graficoHistorico);

  // Gráfico comparativo (barras) — mês atual
  const graficoComparativo = dash.newChart().asColumnChart()
    .addRange(dash.getRange("R1:S5"))
    .setPosition(55, 2, 0, 0)
    .setOption('title', 'Comparativo do Mês')
    .setOption('titleTextStyle', { color: PALETA.escuro, fontSize: 11, bold: true })
    .setOption('legend', { position: 'none' })
    .setOption('backgroundColor', PALETA.card)
    .setOption('colors', [PALETA.azul])
    .setOption('hAxis', { textStyle: { color: PALETA.cinzaTexto } })
    .setOption('vAxis', { textStyle: { color: PALETA.cinzaTexto }, format: '#,##0.0"h"' })
    .setOption('chartArea', { width: '80%', height: '65%' })
    .setOption('width', 380).setOption('height', 260)
    .build();
  dash.insertChart(graficoComparativo);

  // Gráfico de distribuição (pizza) — mês atual
  const graficoPizza = dash.newChart().asPieChart()
    .addRange(dash.getRange("U1:V4"))
    .setPosition(55, 9, 0, 0)
    .setOption('title', 'Distribuição de Horas')
    .setOption('titleTextStyle', { color: PALETA.escuro, fontSize: 11, bold: true })
    .setOption('legend', { position: 'bottom' })
    .setOption('backgroundColor', PALETA.card)
    .setOption('colors', [PALETA.azul, PALETA.verde, PALETA.vermelho])
    .setOption('pieSliceText', 'value')
    .setOption('chartArea', { width: '85%', height: '72%' })
    .setOption('width', 380).setOption('height', 260)
    .build();
  dash.insertChart(graficoPizza);

  dash.getRange("A1").setNote(MARCADOR_LAYOUT);
}

/** Cria a moldura visual de um conjunto de cards, sem preencher valores. */
function construirEstruturaCard(dash, cards, linhaInicial, corFundo) {
  cards.forEach(c => {
    const cols = c.cFim - c.cIni + 1;

    dash.getRange(linhaInicial, c.cIni, 1, cols).merge();
    dash.getRange(linhaInicial + 1, c.cIni, 3, cols).setBackground(corFundo);

    dash.getRange(linhaInicial + 1, c.cIni, 1, cols).merge()
      .setValue(c.titulo).setFontFamily("Segoe UI").setFontSize(8).setFontWeight("bold")
      .setFontColor(PALETA.cinzaTexto).setHorizontalAlignment("center").setVerticalAlignment("bottom");

    dash.getRange(linhaInicial + 2, c.cIni, 1, cols).merge()
      .setFontFamily("Segoe UI").setFontSize(16).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    dash.getRange(linhaInicial + 3, c.cIni, 1, cols).merge()
      .setFontFamily("Segoe UI").setFontSize(8).setFontColor(PALETA.cinzaClaro)
      .setHorizontalAlignment("center").setVerticalAlignment("top");

    dash.getRange(linhaInicial, c.cIni, 4, cols)
      .setBorder(true, true, true, true, false, false, PALETA.borda, SpreadsheetApp.BorderStyle.SOLID);
  });
}

/** Sobrescreve apenas os valores dinâmicos da dashboard (rápido). */
function atualizarValoresDashboard(dash, resumo, historico, sheetAtual) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fuso = ss.getSpreadsheetTimeZone();
  const agoraStr = Utilities.formatDate(new Date(), fuso, "dd/MM/yyyy HH:mm");
  dash.getRange("B3").setValue(`Mês Atual: ${sheetAtual.getName()}   |   Sincronizado: ${agoraStr}   |   Jornada: 08:48/dia útil   |   ${historico.length} abas processadas`);

  const metaMinutosMes = resumo.diasUteis * JORNADA_PADRAO_MINUTOS;
  const saldoPositivo = resumo.saldoBanco >= 0;
  const corSaldo = saldoPositivo ? PALETA.verde : PALETA.vermelho;

  const setaTendencia = resumo.tendenciaVsHistorico === "melhor" ? " ▲"
    : (resumo.tendenciaVsHistorico === "pior" ? " ▼" : "");
  const metaBanco = (saldoPositivo ? "Saldo positivo" : "Horas a pagar") +
    (resumo.tendenciaVsHistorico ? `${setaTendencia} vs média histórica` : "");

  preencherValoresCard(dash, [
    { cIni: 2, valor: minutosParaHorasFormatado(resumo.totalTrabalhado) + "h", meta: `Meta: ${minutosParaHorasFormatado(metaMinutosMes)}h`, cor: PALETA.escuro },
    { cIni: 5, valor: (saldoPositivo ? "+" : "-") + minutosParaHorasFormatado(Math.abs(resumo.saldoBanco)) + "h", meta: metaBanco, cor: corSaldo },
    { cIni: 8, valor: "+" + minutosParaHorasFormatado(resumo.totalExtras) + "h", meta: `${resumo.diasTrabalhados} dias trabalhados`, cor: PALETA.verde },
    { cIni: 11, valor: "-" + minutosParaHorasFormatado(resumo.totalFaltas) + "h", meta: `A regularizar`, cor: PALETA.vermelho }
  ], 5);

  const ajusteTexto = resumo.ajusteDiarioNecessario === null ? "—"
    : (resumo.ajusteDiarioNecessario >= 0 ? `+${minutosParaHorasFormatado(resumo.ajusteDiarioNecessario)}h/dia` : `${minutosParaHorasFormatado(resumo.ajusteDiarioNecessario)}h/dia`);
  const projecaoPositiva = resumo.projecaoFechamento >= 0;

  preencherValoresCard(dash, [
    { cIni: 2, valor: String(resumo.diasUteisRestantes), meta: "dias úteis até o fim do mês", cor: PALETA.escuro },
    { cIni: 5, valor: minutosParaHorasFormatado(resumo.mediaTrabalhadaDiaria) + "h", meta: "por dia trabalhado", cor: PALETA.azul },
    { cIni: 8, valor: ajusteTexto, meta: resumo.ajusteDiarioNecessario === null ? "sem dias restantes" : "p/ zerar o banco", cor: resumo.ajusteDiarioNecessario > 0 ? PALETA.vermelho : PALETA.verde },
    { cIni: 11, valor: (projecaoPositiva ? "+" : "") + minutosParaHorasFormatado(resumo.projecaoFechamento, true) + "h", meta: "saldo estimado no fim do mês", cor: projecaoPositiva ? PALETA.verde : PALETA.vermelho }
  ], 13);

  const perc = metaMinutosMes > 0 ? Math.min(100, Math.round((resumo.totalTrabalhado / metaMinutosMes) * 100)) : 0;
  const cheios = Math.round(perc / 10);
  const barraStr = "█".repeat(cheios) + "░".repeat(10 - cheios);
  const statusTexto = perc >= 100 ? "META BATIDA" : "EM ANDAMENTO";
  dash.getRange("B10").setValue(`JORNADA DO MÊS • [ ${barraStr} ] ${perc}% • ${statusTexto} • ${resumo.diasTrabalhados}/${resumo.diasUteis} dias úteis`);

  atualizarDadosGraficos(dash, resumo, historico);
}

function preencherValoresCard(dash, valores, linhaInicial) {
  valores.forEach(v => {
    dash.getRange(linhaInicial + 2, v.cIni).setValue(v.valor).setFontColor(v.cor);
    dash.getRange(linhaInicial + 3, v.cIni).setValue(v.meta);
  });
}

/** Atualiza só os dados que alimentam os gráficos já existentes. */
function atualizarDadosGraficos(dash, resumo, historico) {
  const metaMinutosMes = resumo.diasUteis * JORNADA_PADRAO_MINUTOS;

  // Tendência diária (mês atual), em horas decimais
  let linhasTendencia = [];
  for (let i = 0; i < MAX_DIAS_TENDENCIA; i++) {
    const ponto = resumo.serieTendencia[i];
    linhasTendencia.push(ponto ? [ponto.label, ponto.saldoAcumulado / 60] : ["", null]);
  }
  dash.getRange(2, 15, MAX_DIAS_TENDENCIA, 2).setValues(linhasTendencia);

  // Histórico mensal (todas as abas), em horas decimais
  let linhasHistorico = [];
  for (let i = 0; i < MAX_MESES_HISTORICO; i++) {
    const mes = historico[i];
    linhasHistorico.push(mes ? [mes.nome, mes.saldoBanco / 60] : ["", null]);
  }
  dash.getRange(2, 24, MAX_MESES_HISTORICO, 2).setValues(linhasHistorico);

  // Comparativo (barras) — mês atual, em horas decimais
  dash.getRange("R2:S5").setValues([
    ["Meta", metaMinutosMes / 60],
    ["Trabalhadas", resumo.totalTrabalhado / 60],
    ["Extras", resumo.totalExtras / 60],
    ["Faltas", Math.max(0, resumo.totalFaltas) / 60]
  ]);

  // Distribuição (pizza) — mês atual, em horas decimais
  dash.getRange("U2:V4").setValues([
    ["Trabalhado", resumo.totalTrabalhado / 60],
    ["Extras", resumo.totalExtras / 60],
    ["Faltas", Math.max(0, resumo.totalFaltas) / 60]
  ]);
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
    if (dados[i][0] && dados[i][0].trim() === dataHojeStr) {
      linhaEncontrada = i + 1;
      break;
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
  if (sheet.getName() === NOME_ABA_DASHBOARD || sheet.getName() === NOME_ABA_FERIADOS) return;

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
    exportFormat: 'pdf', format: 'pdf', size: 'A4', portrait: 'false',
    fitw: 'true', gridlines: 'true', sheetnames: 'false', gid: sheet.getSheetId()
  };

  let queryUrl = [];
  for (let key in parametrosPDF) queryUrl.push(key + '=' + parametrosPDF[key]);
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
