/**
 * SISTEMA INTEGRADO DE CONTROLE DE PONTO E BANCO DE HORAS COM DASHBOARD
 * Jornada Padrão: 8h48min por dia útil (528 minutos)
 *
 * v4 — Melhorias sobre a v3:
 *  - CORREÇÃO IMPORTANTE: o alerta por e-mail rodava ANTES de desenhar a
 *    dashboard. Como gatilhos automáticos de edição (onEdit) não têm
 *    permissão para usar MailApp, isso quebrava a execução e a dashboard
 *    nunca era atualizada sozinha. Agora a dashboard sempre é desenhada
 *    primeiro, e o alerta roda depois, dentro de um try/catch — se falhar
 *    (por ser um gatilho automático sem permissão), a dashboard já foi
 *    atualizada mesmo assim. O e-mail em si só sai com certeza quando você
 *    roda algo pelo menu (que já é uma execução autorizada).
 *  - onEdit agora detecta edições em bloco/colagens (olha a faixa inteira,
 *    não só a primeira célula) e também recalcula quando você edita a aba
 *    "Feriados".
 *  - Visual: ícones nos títulos dos cards, e o histórico mensal agora usa
 *    barras verdes (saldo positivo) e vermelhas (saldo negativo) em vez de
 *    uma cor única.
 *
 * v3 trouxe: feriados, histórico entre todas as abas, alerta automático.
 * v2 trouxe: layout construído uma única vez (performance), gráfico de
 * tendência do saldo diário, indicadores de projeção/médias.
 */
