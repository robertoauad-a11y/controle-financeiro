import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Plus, Trash2, X, Wallet, TrendingUp, TrendingDown, ClipboardList,
  Calendar, Check, AlertCircle, Clock, Filter, ChevronDown, PiggyBank,
  Home as HomeIcon, ArrowRight, Download, Pencil, Eye, EyeOff, Sun, Moon, Bell
} from "lucide-react";
import * as XLSX from "xlsx";
import { auth, db } from "./firebase.js";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { collection, doc, addDoc, updateDoc, deleteDoc, onSnapshot, writeBatch } from "firebase/firestore";
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, ReferenceLine, AreaChart, Area, LabelList, CartesianGrid } from "recharts";

/* ------------------------------------------------------------------ */
/*  Firestore collections — one document per record. Never write an   */
/*  auto-generated "empty" document; only explicit user actions       */
/*  (add / edit / delete / import) touch the database.                */
/* ------------------------------------------------------------------ */
const custosCol = collection(db, "custos");
const receitasCol = collection(db, "receitas");
const demandasCol = collection(db, "demandas");
const fixosCol = collection(db, "recorrentes");   // list of monthly fixed costs (only changed by explicit action)

const MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const FONTES = ["Conta Salário","Conta Corrente","Dinheiro","PIX","Transferência","Débito Automático","BB","Transferência (Corrente)","Transferência (Salário)"];
// PIX and Débito Automático always leave the Conta Corrente, so they count as it in totals
// and charts (the record keeps its own fonte). "BB" is just the bank's abbreviation, not an
// account, so it is no longer offered for new entries.
const FONTES_FORM = FONTES.filter(f => f !== "BB" && f !== "Transferência");
// Transferência leaves either sub-account of the same bank account, so it is split in two options.
const contaDaFonte = (f) => (f === "PIX" || f === "Débito Automático" || f === "Transferência (Corrente)") ? "Conta Corrente" : f === "Transferência (Salário)" ? "Conta Salário" : f;
// "Depósito …" (cash put into the bank): not income, money just moves from Dinheiro to the receita's conta.
const ehDeposito = (r) => /^\s*dep[óo]sito\b/i.test(r.descricao || "");
// Where a receita normally lands: salário goes to Conta Salário; aluguéis (and deposits) go to Conta Corrente.
const destinoPadraoReceita = (descricao) => /sal[áa]rio/i.test(descricao || "") ? "Conta Salário" : /aluguel|dep[óo]sito/i.test(descricao || "") ? "Conta Corrente" : null;
const fonteCombina = (c, filtro) => filtro === "Todas" || c.fonte === filtro || contaDaFonte(c.fonte) === filtro;
// PIX is a payment method, not an account: this only counts how many were registered and their sum.
const resumoPix = (lista, mes) => {
  const itens = lista.filter(c => c.mes === mes && c.fonte === "PIX");
  const pagos = itens.filter(c => c.status === "Pago");
  return {
    n: pagos.length, total: pagos.reduce((t, c) => t + Number(c.valorPago || 0), 0),
    nPend: itens.length - pagos.length, totalPend: itens.filter(c => c.status !== "Pago").reduce((t, c) => t + Number(c.valorPrevisto || 0), 0),
    itens: [...pagos].sort((a, b) => String(b.dataPagamento || "").localeCompare(String(a.dataPagamento || "")))
  };
};
const notasMeios = (lista) => {
  const g = {};
  lista.filter(c => c.fonte !== contaDaFonte(c.fonte)).forEach(c => { const k = contaDaFonte(c.fonte); g[k] = g[k] || []; if (!g[k].includes(c.fonte)) g[k].push(c.fonte); });
  return Object.entries(g).map(([k, m]) => `${k} inclui ${m.join(" e ")}`);
};
const opcoesFonte = (atual) => (atual && !FONTES_FORM.includes(atual)) ? [...FONTES_FORM, atual] : FONTES_FORM;
// Só na apresentação: custo não pago cujo vencimento já passou conta como "Atrasado" (o dado salvo não muda).
const statusEfetivo = (c) => (c.status !== "Pago" && c.vencimento && c.vencimento < localISO() ? "Atrasado" : c.status);
const STATUS_CUSTO = ["Pago","Pendente","Atrasado"];
const STATUS_RECEITA = ["Recebido","Parcial","Pendente"];
const PRIORIDADES = ["Alta","Média","Baixa"];
const STATUS_DEMANDA = ["Pendente","Em Andamento","Concluída"];

const fmtBRL = (v) => {
  const n = Number(v || 0);
  return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
};
const fmtDate = (s) => {
  if (!s) return "—";
  const [y,m,d] = s.split("-");
  if (!y || !m || !d) return s;
  return `${d}/${m}/${y}`;
};
const monthLabel = (mes) => {
  if (!mes) return "—";
  const [y,m] = mes.split("-");
  return `${MESES[parseInt(m,10)-1]}/${y}`;
};
const parseMonthLabel = (label) => {
  if (!label) return null;
  const [nome, y] = String(label).split("/");
  const idx = MESES.findIndex(m => m.toLowerCase() === (nome||"").trim().toLowerCase());
  if (idx === -1 || !y) return null;
  return `${y.trim()}-${String(idx+1).padStart(2,"0")}`;
};
const parseDateBR = (s) => {
  if (!s || s === "—") return null;
  const str = String(s).trim();
  const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
  // Handle Excel serial date numbers or Date objects passed through by SheetJS
  if (str.match(/^\d{4}-\d{2}-\d{2}/)) return str.slice(0,10);
  return null;
};
const uid = (p) => `${p}${Date.now().toString(36)}${Math.random().toString(36).slice(2,7)}`;

// Turns free text like "aluguel 2675 dia 10" or "recebi 800 freelance dia 15"
// into a draft entry. Deliberately forgiving — it only has to get close;
// the person confirms/edits everything before it's saved.
// Fonte keywords the quick-entry text is scanned for — longer/more specific
// phrases first, so "conta corrente" matches before a bare "conta" would.
const FONTE_KEYWORDS = [
  { re: /\bconta\s+corrente\b/i, fonte: "Conta Corrente" },
  { re: /\bconta\s+sal[áa]rio\b/i, fonte: "Conta Salário" },
  { re: /\bd[ée]bito\s+autom[áa]tico\b/i, fonte: "Débito Automático" },
  { re: /\btransfer[êe]ncia\b/i, fonte: "Transferência" },
  { re: /\bdinheiro\b/i, fonte: "Dinheiro" },
  { re: /\bpix\b/i, fonte: "PIX" },
];
const fonteEmTexto = (segment) => {
  for (const { re, fonte } of FONTE_KEYWORDS) if (re.test(segment)) return fonte;
  return null;
};

// Parses free text from the "+" quick-entry FAB. Normally returns a single
// draft entry, but when the sentence names two or more amounts each tied to
// their OWN fonte ("1500 em dinheiro e 299 em pix") — a split payment across
// sources, same as launching separate custo records by hand — it returns one
// draft per amount+fonte instead of silently keeping only the largest number.
function parseQuickEntry(rawText) {
  let text = rawText.trim();
  const lower = text.toLowerCase();
  let dataStr = null;

  const fullDate = text.match(/(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/);
  if (fullDate) {
    const d = fullDate[1].padStart(2, "0");
    const m = fullDate[2].padStart(2, "0");
    let y = fullDate[3] || String(new Date().getFullYear());
    if (y.length === 2) y = "20" + y;
    dataStr = `${y}-${m}-${d}`;
    text = text.replace(fullDate[0], " ");
  } else {
    const diaMatch = lower.match(/\bdia\s+(\d{1,2})\b/);
    if (diaMatch) {
      const now = new Date();
      const dia = String(parseInt(diaMatch[1], 10)).padStart(2, "0");
      const mes = String(now.getMonth() + 1).padStart(2, "0");
      dataStr = `${now.getFullYear()}-${mes}-${dia}`;
      text = text.replace(new RegExp(diaMatch[0].replace(/\s+/g, "\\s+"), "i"), " ");
    }
  }
  // Horário ("10:00", "10h", "10h30", "às 10") — nunca é valor em R$. Sai do texto
  // antes de procurar valores e, em demanda, vai para os Detalhes.
  let hora = "";
  const horaMatch = text.match(/\b(\d{1,2})\s*(?::|h)\s*(\d{2})?\b(?!\s*(?:reais|r\$))/i) || text.match(/\bàs\s+(\d{1,2})\b()/i);
  if (horaMatch && +horaMatch[1] <= 23 && (!horaMatch[2] || +horaMatch[2] <= 59)) {
    hora = `${horaMatch[1].padStart(2, "0")}:${horaMatch[2] || "00"}`;
    text = text.replace(horaMatch[0], " ");
  }
  // Mês por extenso: "19 de novembro", "dia 19 ... novembro", "em novembro" (sem dia).
  const MESES_ABREV = ["janeiro","fevereiro","marco","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
  let semDiaNoMes = false;
  if (!fullDate) {
    const mm = lower.normalize("NFD").replace(/[\u0300-\u036f]/g, "").match(/\b(janeiro|fevereiro|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\b/);
    if (mm) {
      const mi = MESES_ABREV.indexOf(mm[1]);
      const dm = lower.match(/\b(?:dia\s+)?(\d{1,2})\s+de\b/) || lower.match(/\bdia\s+(\d{1,2})\b/);
      const now = new Date(); let ano = now.getFullYear();
      if (mi < now.getMonth()) ano += 1;
      if (dm) {
        dataStr = `${ano}-${String(mi + 1).padStart(2, "0")}-${String(+dm[1]).padStart(2, "0")}`;
        text = text.replace(new RegExp("\\b(?:dia\\s+)?" + dm[1] + "(?:\\s+de)?(?![\\d])", "i"), " ");
      } else { semDiaNoMes = true; }
      text = text.replace(new RegExp("\\b" + mm[1].replace("marco", "mar[çc]o") + "\\b", "i"), " ");
    }
  }
  if (!dataStr) dataStr = localISO();
  const textoSemHora = text;

  // "salário" alone suggests income ("recebi meu salário"), but "conta
  // salário" is just the fonte name — strip recognized fonte phrases first so
  // mentioning that account doesn't misclassify a custo as a receita.
  let lowerSemFonte = lower;
  for (const { re } of FONTE_KEYWORDS) lowerSemFonte = lowerSemFonte.replace(re, " ");
  const isReceita = /receb|sal[áa]rio|entrada|venda|renda|dep[óo]sito/i.test(lowerSemFonte);
  const jaRecebido = /\breceb/i.test(lowerSemFonte);
  const jaPago = /\bpagu?ei\b|\bpago\b|\bfeito\b|\bfeita\b|\bquitei\b|\bquitado\b|\bliquidei\b|\bliquidado\b|\bfiz\b/i.test(lowerSemFonte);
  const status = isReceita ? (jaRecebido ? "Recebido" : "Pendente") : (jaPago ? "Pago" : "Pendente");
  // Demanda = compromisso/tarefa: palavra de agenda e NENHUM sinal de dinheiro
  // (R$, "reais", fonte de pagamento, "paguei/recebi"). Ex.: "consulta médica 19/11 10:00".
  const temDinheiro = /r\$|\breais\b|\bvalor\b/i.test(lower) || FONTE_KEYWORDS.some(({ re }) => re.test(lower)) || jaPago || jaRecebido;
  const palavraAgenda = /\b(consulta|exame|retorno|agendar|agende|agendamento|marcar|marque|reuni[ãa]o|entrevista|lembrar|lembrete|ligar|providenciar|renovar|dentista|m[ée]dic[oa]|vistoria|ensaio|show|compromisso|demanda|tarefa)\b/i.test(lower);
  const valoresSoltos = [...text.matchAll(/\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d+(?:,\d{1,2})?/g)];
  const isDemanda = palavraAgenda && !temDinheiro && !isReceita && valoresSoltos.length === 0;
  const tipo = isDemanda ? "demanda" : isReceita ? "receita" : "custo";
  if (isDemanda) {
    const nome = text.replace(/(?<![\p{L}\d])(no|na|em|para|dia|de|do|da|às|as|à)(?![\p{L}\d])/giu, " ").replace(/[,.]/g, " ").replace(/\s+/g, " ").trim();
    return [{ tipo, descricao: (nome || text.trim()).toUpperCase(), valor: 0, data: semDiaNoMes ? "" : dataStr, status: "Pendente", fonte: "", hora }];
  }

  // The dot-grouped form only applies when a ".XXX" group actually follows
  // (the "+" below, not "*") — otherwise a plain 4+ digit run like "2675" or
  // "1500" was being cut to its first 3 digits ("267"/"150") by the grouped
  // alternative matching before the plain one got a chance to run.
  const valorMatches = [...text.matchAll(/\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d+(?:,\d{1,2})?/g)]
    .map(m => ({ raw: m[0], value: parseValorBR(m[0]), index: m.index }));

  // The fonte "for" a given amount is whatever fonte keyword sits between it
  // and the next amount (or the end of the text) — e.g. in "1500 em dinheiro
  // e 299 em pix", "dinheiro" belongs to 1500 and "pix" belongs to 299.
  const comFonte = valorMatches.map((v, i) => {
    const fim = i + 1 < valorMatches.length ? valorMatches[i + 1].index : text.length;
    return { ...v, fonte: fonteEmTexto(text.slice(v.index, fim)) };
  }).filter(v => v.fonte);
  const fontesDistintas = new Set(comFonte.map(v => v.fonte));

  const buildDescricao = () => {
    let t = text;
    for (const v of valorMatches) t = t.replace(v.raw, " ");
    for (const { re } of FONTE_KEYWORDS) t = t.replace(re, " ");
    // Drop connector/filler words so only the subject remains — "Recebi aluguel do
    // VIA IPIRANGA, no dia 10, no valor de 3500" -> "ALUGUEL VIA IPIRANGA".
    t = t
      .replace(/\br\$/gi, " ").replace(/\breais?\b/gi, " ")
      .replace(/\bno valor de\b/gi, " ").replace(/\bvalor de\b/gi, " ")
      .replace(/\bno dia\b/gi, " ").replace(/\bpagamento( de)?\b/gi, " ")
      .replace(/\brecebi\b/gi, " ").replace(/\bpaguei\b/gi, " ").replace(/\bem\b/gi, " ")
      .replace(/[,.]/g, " ");
    // JS's \b treats accented letters (ã, ç, é…) as non-word characters, so
    // "Colchão" would lose its final "o" to a plain \bo\b. Short filler words
    // are matched with Unicode-aware lookarounds instead.
    const palavra = (alts) => new RegExp(`(?<![\\p{L}\\d])(?:${alts})(?![\\p{L}\\d])`, "giu");
    t = t
      .replace(palavra("hoje|no|na|nos|nas"), " ")
      // A trailing "a"/"o" is part of the name ("Plano A"), not an article.
      .replace(palavra("o|a|os|as"), (m, off, whole) => (whole.slice(off + m.length).trim() === "" ? m : " "))
      .replace(palavra("feito|feita|quitei|quitado|liquidei|liquidado|fiz"), " ")
      .replace(palavra("de|da|do|e"), " ")
      .replace(/\s+/g, " ").trim();
    return t.toUpperCase() || "NOVO LANÇAMENTO";
  };
  const descricao = buildDescricao();

  // Split into several drafts only when at least two amounts each resolved to
  // a DIFFERENT fonte — a confident "part here, part there" signal. A lone
  // second number with no fonte word (an installment count, a day already
  // consumed above, etc.) stays a single entry, same as before.
  if (comFonte.length >= 2 && fontesDistintas.size >= 2) {
    return comFonte.map(v => ({ tipo, descricao, valor: v.value, data: dataStr, status, fonte: v.fonte }));
  }

  const valor = valorMatches.length ? valorMatches.reduce((a, b) => (b.value > a.value ? b : a)).value : 0;
  const fonte = fonteEmTexto(text);
  return [{ tipo, descricao, valor, data: dataStr, status, fonte }];
}

// Words too common to mean anything when comparing a custo's descrição
// against a demanda's nome — filtered out of both sides before comparing.
const STOPWORDS_PT = new Set(["DE","DA","DO","DOS","DAS","E","A","O","OS","AS","EM","NO","NA","NOS","NAS","PARA","COM","UM","UMA"]);
const normTokens = (s) => (s || "")
  .normalize("NFD").replace(/[̀-ͯ]/g, "")
  .toUpperCase().replace(/[^A-Z0-9\s]/g, " ")
  .split(/\s+/).filter(w => w.length > 2 && !STOPWORDS_PT.has(w));

// Suggests an open demanda (Pendente/Em Andamento) that a quick-entry custo
// likely resolves — e.g. a "Colchão Rubem" custo matching the "Colchão
// Rubem" demanda. Conservative on purpose: only suggests when EVERY
// significant word of the demanda's nome shows up in the custo's descrição,
// so a partial/loose overlap never triggers a wrong suggestion. The person
// still confirms it with a checkbox before anything closes.
const matchDemanda = (descricao, demandas) => {
  const tokensDesc = new Set(normTokens(descricao));
  if (tokensDesc.size === 0) return null;
  for (const d of demandas) {
    if (d.status === "Concluída") continue;
    const tokensNome = normTokens(d.nome);
    if (tokensNome.length === 0) continue;
    if (tokensNome.every(t => tokensDesc.has(t))) return d;
  }
  return null;
};
// Local calendar date (not UTC): in UTC-4 the UTC date flips to "tomorrow" at 20:00,
// which would stamp new records with the wrong day and the wrong month on the last evening of a month.
const localISO = (d = new Date()) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const todayISO = () => localISO();
const currentMonth = () => localISO().slice(0, 7);
const parseValorBR = (raw) => {
  if (raw === null || raw === undefined || raw === "") return 0;
  let s = String(raw).trim().replace(/[^\d.,-]/g, "");
  if (!s) return 0;
  if (s.includes(",")) {
    // Vírgula presente: pontos são milhar, vírgula é decimal — "5.806,00" -> 5806.00
    s = s.replace(/\./g, "").replace(",", ".");
  } else if ((s.match(/\./g) || []).length > 1 || /\.\d{3}$/.test(s)) {
    // Só pontos, em grupos de 3 (ou mais de um ponto) -> são todos separadores de milhar — "5.806" -> 5806
    s = s.replace(/\./g, "");
  }
  // Caso contrário (um único ponto com 1-2 dígitos depois, ex. "5.8") já é o decimal certo.
  return parseFloat(s) || 0;
};
// Value shown inside editable inputs, in Brazilian format ("2.675,00"). Round-
// trips safely through parseValorBR (comma present → dots are thousands).
const fmtValorInput = (n) => {
  if (n === null || n === undefined || n === "") return "";
  const num = Number(n);
  if (!isFinite(num)) return "";
  return num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};
// "Contar a partir de": month ("YYYY-MM") the yearly balance starts counting
// from — lets months with incomplete history stay out of the Saldo do Ano /
// Saldo acumulado. Kept on this device; storage can be blocked, so every
// access is wrapped.
const INICIO_KEY = "cf_inicio_contagem";
// Real receitas only started being logged in September 2026, so earlier months
// would distort the yearly balance — that's the default until changed.
const INICIO_PADRAO = "2026-09";
const TEMA_KEY = "cf_tema";   // "escuro" (default) | "claro" | "auto"
const lerTemaPref = () => { try { const v = localStorage.getItem(TEMA_KEY); return v === "claro" || v === "auto" ? v : "escuro"; } catch (e) { return "escuro"; } };
const salvarTemaPref = (v) => { try { localStorage.setItem(TEMA_KEY, v); } catch (e) { /* storage unavailable — choice just won't persist */ } };
const temaEfetivo = (pref) => pref === "auto"
  ? ((typeof window !== "undefined" && window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches) ? "claro" : "escuro")
  : pref;
const lerInicioContagem = () => { try { return localStorage.getItem(INICIO_KEY) || INICIO_PADRAO; } catch (e) { return INICIO_PADRAO; } };
const salvarInicioContagem = (v) => { try { localStorage.setItem(INICIO_KEY, v); } catch (e) { /* storage unavailable — setting just won't persist */ } };
// Index (0–11) of the first counted month for the given year; 0 when nothing
// is set or the saved start belongs to a different year.
const inicioIdxDoAno = (inicio, ano) => {
  if (!inicio) return 0;
  const [y, m] = inicio.split("-").map(Number);
  return y === Number(ano) && m >= 1 && m <= 12 ? m - 1 : 0;
};
const shiftMonth = (mes, delta) => {
  const [y, m] = mes.split("-").map(Number);
  const d = new Date(y, (m - 1) + delta, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
};
const daysInMonthOf = (mes) => {
  const [y, m] = mes.split("-").map(Number);
  return new Date(y, m, 0).getDate();
};
// Sum of valorPago/valorRecebido up to a given day-of-month — used to compare
// "this month so far" against "last month, same point in the month" instead of
// against a full previous month (which always makes the current partial month
// look artificially better early on, and worse near the end).
const pagoAteDia = (custos, mes, dia) => custos
  .filter(c => c.mes === mes && c.dataPagamento && Number(c.dataPagamento.split("-")[2]) <= dia)
  .reduce((s, c) => s + Number(c.valorPago || 0), 0);
const recebidoAteDia = (receitas, mes, dia) => receitas
  .filter(r => r.mes === mes && r.dataRecebimento && Number(r.dataRecebimento.split("-")[2]) <= dia)
  .reduce((s, r) => s + Number(r.valorRecebido || 0), 0);

/* ------------------------------------------------------------------ */
/*  Validação de dados                                                 */
/* ------------------------------------------------------------------ */
// Item 1: valor precisa ser > 0 — bloqueia o salvamento.
const isValorValido = (n) => Number.isFinite(n) && n > 0;

// Item 2: data muito distante do presente — apenas avisa, não bloqueia.
const isDateImplausible = (dateStr) => {
  if (!dateStr) return false;
  const d = new Date(`${dateStr}T00:00:00`);
  if (isNaN(d.getTime())) return false;
  const diffYears = Math.abs(d.getFullYear() - new Date().getFullYear());
  return diffYears >= 5;
};
const confirmDateSanity = async (dateStr, label = "data") => {
  if (!isDateImplausible(dateStr)) return true;
  return confirmar(
    `A ${label} informada (${fmtDate(dateStr)}) parece muito distante da data de hoje. Confirma que está correta?`,
    { titulo: "Data estranha", okLabel: "Está correta" }
  );
};

function exportToExcel(data, prefixo = "Controle_Financeiro") {
  const wb = XLSX.utils.book_new();

  const custosRows = data.custos.map(c => ({
    "Mês": monthLabel(c.mes), "Ano": c.ano, "Descrição / Categoria": c.descricao,
    "Vencimento": fmtDate(c.vencimento), "Data Pagamento": fmtDate(c.dataPagamento),
    "Valor Previsto (R$)": Number(c.valorPrevisto || 0), "Valor Pago (R$)": Number(c.valorPago || 0),
    "Diferença (R$)": Number(c.valorPago || 0) - Number(c.valorPrevisto || 0),
    "Fonte do Pagamento": c.fonte, "Status": c.status, "Observações": c.obs || ""
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(custosRows), "Custos Mensais");

  const receitasRows = data.receitas.map(r => ({
    "Mês": monthLabel(r.mes), "Ano": r.ano, "Descrição / Origem": r.descricao,
    "Data Recebimento": fmtDate(r.dataRecebimento),
    "Valor Previsto (R$)": Number(r.valorPrevisto || 0), "Valor Recebido (R$)": Number(r.valorRecebido || 0),
    "Diferença (R$)": Number(r.valorRecebido || 0) - Number(r.valorPrevisto || 0),
    "Categoria": r.categoria, "Conta de Destino": r.contaDestino, "Status": r.status, "Observações": r.obs || ""
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(receitasRows), "Receitas Mensais");

  const anos = Array.from(new Set([...data.custos, ...data.receitas].map(x => x.ano).filter(Boolean))).sort();
  const resumoRows = [];
  anos.forEach(ano => {
    MESES.forEach((nome, idx) => {
      const mesKey = `${ano}-${String(idx + 1).padStart(2, "0")}`;
      const cs = data.custos.filter(c => c.mes === mesKey);
      const rs = data.receitas.filter(r => r.mes === mesKey);
      if (cs.length === 0 && rs.length === 0) return;
      const totalPrevistoCusto = cs.reduce((s, c) => s + Number(c.valorPrevisto || 0), 0);
      const totalPagoCusto = cs.reduce((s, c) => s + Number(c.valorPago || 0), 0);
      const totalRecebido = rs.reduce((s, r) => s + Number(r.valorRecebido || 0), 0);
      resumoRows.push({
        "Ano": ano, "Mês": nome,
        "Previsto (Custos)": totalPrevistoCusto, "Pago (Custos)": totalPagoCusto,
        "Diferença": totalPagoCusto - totalPrevistoCusto,
        "Recebido": totalRecebido, "Saldo do Mês": totalRecebido - totalPagoCusto,
        "Status": cs.some(c => c.status !== "Pago") ? "Pendências" : (cs.length ? "Em dia" : "—")
      });
    });
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumoRows), "Resumo Anual");

  const demandasRows = data.demandas.map(d => ({
    "Nome da Demanda": d.nome, "Contato": d.contato || "", "Detalhes": d.detalhes || "",
    "Prazo": fmtDate(d.prazo), "Prioridade": d.prioridade, "Status": d.status, "Observações": d.obs || ""
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(demandasRows), "Demandas");

  XLSX.writeFile(wb, `${prefixo}_${todayISO()}.xlsx`);
}

function importFromExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Falha ao ler o arquivo"));
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const custos = [];
        const receitas = [];
        const demandas = [];

        const custosSheet = wb.Sheets["Custos Mensais"];
        if (custosSheet) {
          XLSX.utils.sheet_to_json(custosSheet).forEach((row, i) => {
            const mes = parseMonthLabel(row["Mês"]);
            if (!mes || !row["Descrição / Categoria"]) return;
            custos.push({
              id: uid("c"), mes, ano: parseInt(mes.split("-")[0], 10),
              descricao: row["Descrição / Categoria"], vencimento: parseDateBR(row["Vencimento"]),
              dataPagamento: parseDateBR(row["Data Pagamento"]),
              valorPrevisto: Number(row["Valor Previsto (R$)"]) || 0,
              valorPago: row["Valor Pago (R$)"] === undefined || row["Valor Pago (R$)"] === "" ? null : Number(row["Valor Pago (R$)"]),
              fonte: row["Fonte do Pagamento"] || FONTES[0], status: row["Status"] || "Pendente",
              obs: row["Observações"] || ""
            });
          });
        }

        const receitasSheet = wb.Sheets["Receitas Mensais"];
        if (receitasSheet) {
          XLSX.utils.sheet_to_json(receitasSheet).forEach((row) => {
            const mes = parseMonthLabel(row["Mês"]);
            if (!mes || !row["Descrição / Origem"]) return;
            receitas.push({
              id: uid("r"), mes, ano: parseInt(mes.split("-")[0], 10),
              descricao: row["Descrição / Origem"], dataRecebimento: parseDateBR(row["Data Recebimento"]),
              valorPrevisto: Number(row["Valor Previsto (R$)"]) || 0,
              valorRecebido: row["Valor Recebido (R$)"] === undefined || row["Valor Recebido (R$)"] === "" ? null : Number(row["Valor Recebido (R$)"]),
              categoria: row["Categoria"] || "", contaDestino: row["Conta de Destino"] || FONTES[1],
              status: row["Status"] || "Pendente", obs: row["Observações"] || ""
            });
          });
        }

        const demandasSheet = wb.Sheets["Demandas"];
        if (demandasSheet) {
          XLSX.utils.sheet_to_json(demandasSheet).forEach((row) => {
            if (!row["Nome da Demanda"]) return;
            demandas.push({
              id: uid("d"), nome: row["Nome da Demanda"], contato: row["Contato"] || "",
              detalhes: row["Detalhes"] || "", prazo: parseDateBR(row["Prazo"]),
              prioridade: row["Prioridade"] || "Alta", status: row["Status"] || "Pendente",
              obs: row["Observações"] || ""
            });
          });
        }

        resolve({ custos, receitas, demandas });
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

/* ------------------------------------------------------------------ */
/*  Design tokens                                                      */
/* ------------------------------------------------------------------ */
const C = {
  ink: "#F3F5FA", inkSoft: "#B8BFDB", page: "#0A0E1C", card: "#141928",
  muted: "#8B93B5", faint: "#5C6485", border: "#252B45",
  emerald: "#00E0A8", emeraldSoft: "#00E0A82E", emeraldDeep: "#3FF0C0", emeraldLight: "#6EE7C3", emeraldGlow: "rgba(0,224,168,0.4)",
  coral: "#FF5A68", coralSoft: "#FF5A682E", coralDeep: "#FF8B94", coralLight: "#FF9AA3",
  amber: "#FFB929", amberSoft: "#FFB9292E", amberDeep: "#FFCB66", amberLight: "#FFCB66",
  lavender: "#9AA5D1",
  // Text and chips that sit directly on the page wallpaper (outside any card). Same as the card tokens in the dark theme.
  pageInk: "#F3F5FA", pageMuted: "#8B93B5", pageFaint: "#5C6485", pageChip: "#252B45",
};
let shadowSm = "0 1px 3px rgba(0,0,0,0.35), 0 0 0 1px rgba(255,255,255,0.04)";
let shadowMd = "0 10px 30px rgba(0,0,0,0.45), 0 0 0 1px rgba(255,255,255,0.05)";
let shadowNav = "0 -8px 24px rgba(0,0,0,0.35)"; // sombra acima do menu inferior (no claro fica bem mais leve, senão vira uma faixa cinza)
// Themes. "escuro" is the app's identity and the default. "claro": cards, text and accents all light
// (the Saldo card on Início keeps its dark gradient in both themes).
const C_ESCURO = { ...C, barCinza: "#5C6485", gradOut: ["#FF7A85", "#E8434F"], gradIn: ["#2CF0BD", "#00B98A"] };
const PAL_OUT_ESCURO = ["#FF5A68", "#FF8A3D", "#FFC53D", "#F472B6", "#A78BFA", "#38BDF8", "#2DD4BF", "#9AA5D1"];
const PAL_IN_ESCURO = ["#00E0A8", "#22D3EE", "#4FA3FF", "#B5E655", "#818CF8", "#FDE047", "#F0ABFC", "#86EFAC"];
// Claro: tons mais calmos e maduros (menos saturados), ainda bem distintos entre si.
const PAL_OUT_CLARO = ["#D9606A", "#DE8A55", "#D9AE4A", "#C879A0", "#8C82C9", "#5E9FC4", "#4FA99B", "#8E97B8"];
const PAL_IN_CLARO = ["#2EA98A", "#4A9DB5", "#5B7FD0", "#9DBB56", "#C98A4B", "#D2B94E", "#B98AC4", "#E07A8A"];
const SHADOWS_ESCURO = [shadowSm, shadowMd];
const TEMAS = {
  escuro: {},
  claro: {
    page: "#EEF1F8", card: "#FFFFFF", ink: "#141B34", inkSoft: "#3A4466", muted: "#5A6385", faint: "#7A83A3", border: "#DCE1EE",
    emerald: "#00A67E", emeraldSoft: "#00A67E24", emeraldDeep: "#008A69", emeraldLight: "#00A67E",
    coral: "#E0414F", coralSoft: "#E0414F22", coralDeep: "#C73341", coralLight: "#E0414F",
    amber: "#C98300", amberSoft: "#C9830022", amberDeep: "#A86C00", amberLight: "#C98300", lavender: "#6B76A8",
    pageInk: "#141B34", pageMuted: "#5A6385", pageFaint: "#7A83A3", pageChip: "#DCE1EE",
    barCinza: "#AEB6D0", gradOut: ["#E5808A", "#D9606A"], gradIn: ["#4DBFA2", "#2EA98A"],
  },
};
function aplicarTema(modo) {
  Object.assign(C, C_ESCURO, TEMAS[modo] || {});
  const claro = modo === "claro";
  PALETTE_OUT = claro ? PAL_OUT_CLARO : PAL_OUT_ESCURO;
  PALETTE_IN = claro ? PAL_IN_CLARO : PAL_IN_ESCURO;
  shadowSm = claro ? "0 1px 2px rgba(20,27,52,0.08), 0 0 0 1px rgba(20,27,52,0.05)" : SHADOWS_ESCURO[0];
  shadowMd = claro ? "0 10px 28px rgba(20,27,52,0.14), 0 0 0 1px rgba(20,27,52,0.06)" : SHADOWS_ESCURO[1];
  shadowNav = claro ? "0 -4px 14px rgba(20,27,52,0.06)" : "0 -8px 24px rgba(0,0,0,0.35)";
  // Style objects built once at load: refresh the parts that use the palette.
  Object.assign(inputStyle, { border: "1px solid " + C.border, background: C.card, color: C.ink });
  Object.assign(monthNavBtnStyle, { border: `1px solid ${C.border}`, background: C.card, color: C.ink });
  Object.assign(menuItemStyle, { color: C.ink });
}
// Closed tonal system: every chart color is a shade of emerald/amber/coral/ink —
// no unrelated hues, so multi-slice charts still read as "one app", not a mix.
// Saídas (custos): tons quentes primeiro; Entradas (receitas): tons frios/verdes primeiro.
// Cada paleta tem 8 cores bem distintas entre si, para nenhuma fatia se confundir com outra.
let PALETTE_OUT = ["#FF5A68", "#FF8A3D", "#FFC53D", "#F472B6", "#A78BFA", "#38BDF8", "#2DD4BF", "#9AA5D1"];
let PALETTE_IN = ["#00E0A8", "#22D3EE", "#4FA3FF", "#B5E655", "#818CF8", "#FDE047", "#F0ABFC", "#86EFAC"];
// Cor fixa por fonte: "Conta Salário" tem sempre a mesma cor em qualquer gráfico ou mês.
const fonteColor = (name) => PALETTE_OUT[Math.max(0, FONTES.indexOf(contaDaFonte(name))) % PALETTE_OUT.length];
const fonteColorIn = (name) => PALETTE_IN[Math.max(0, FONTES.indexOf(contaDaFonte(name))) % PALETTE_IN.length];
const STATUS_ACCENT = {
  "Pago": C.emerald, "Recebido": C.emerald, "Concluída": C.emerald,
  "Parcial": C.amber, "Em Andamento": C.amber, "Pendente": C.amber,
  "Atrasado": C.coral,
};

function StatusPill({ status }) {
  const dotMap = {
    "Pago": C.emerald, "Recebido": C.emerald, "Concluída": C.emerald,
    "Parcial": C.amber, "Em Andamento": C.amber, "Pendente": C.amber,
    "Atrasado": C.coral,
  };
  const dot = dotMap[status] || C.faint;
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 6,
      background: "transparent", color: C.muted, fontSize: 11.5, fontWeight: 600,
      padding: "3px 2px", borderRadius: 20, whiteSpace: "nowrap"
    }}>
      <span style={{ width: 6, height: 6, borderRadius: "50%", background: dot, display: "inline-block" }} />
      {status}
    </span>
  );
}

function Avatar({ label, color, icon }) {
  return (
    <div style={{
      width: 40, height: 40, borderRadius: 12, background: `${color}17`,
      display: "flex", alignItems: "center", justifyContent: "center",
      flexShrink: 0, fontWeight: 800, fontSize: 15, color
    }}>
      {icon || (label ? label.charAt(0).toUpperCase() : "?")}
    </div>
  );
}

function ProgressBar({ pct, color = C.emerald, trackColor = C.border, height = 8, overColor = C.coral }) {
  const clamped = Math.max(0, Math.min(100, pct));
  const over = pct > 100;
  return (
    <div style={{ width: "100%", height, borderRadius: height, background: trackColor, overflow: "hidden" }}>
      <div style={{
        width: `${clamped}%`, height: "100%", borderRadius: height,
        background: over ? overColor : color, transition: "width .3s ease"
      }} />
    </div>
  );
}

// Compact chart/list tooltip that shows each series' value, its color dot,
// and its share of the total — a bit more context than a bare number.
// `modo` picks what the extra line says: default = each series' share of the total
// (right for donuts); "saldo" = recebido minus pago; "variacao" = change of Atual vs Anterior.
function RichTooltip({ active, payload, label, total, colorMap, modo }) {
  if (!active || !payload || !payload.length) return null;
  const sum = total ?? payload.reduce((s, p) => s + (Number(p.value) || 0), 0);
  const val = (k) => Number((payload.find(p => p.dataKey === k) || {}).value) || 0;
  const mostrarPct = !modo;
  let extra = null;
  if (modo === "saldo") {
    const s = val("Recebido") - val("Pago");
    extra = { rotulo: "Saldo do mês", texto: fmtBRL(s), cor: s < 0 ? C.coral : C.emerald };
  } else if (modo === "variacao") {
    const ant = val("Anterior"), atu = val("Atual");
    if (ant > 0) {
      const pct = ((atu - ant) / ant) * 100;
      // For costs, going up is the bad direction.
      extra = { rotulo: "Variação", texto: `${pct > 0 ? "+" : ""}${pct.toFixed(0)}%`, cor: pct > 0 ? C.coral : C.emerald };
    }
  }
  return (
    <div style={{ background: C.card, boxShadow: shadowMd, borderRadius: 10, padding: "10px 12px", fontSize: 12.5, minWidth: 150 }}>
      {label != null && label !== "" && (
        <div style={{ color: C.muted, fontWeight: 700, marginBottom: 6, fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.3 }}>{label}</div>
      )}
      {payload.map((p, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 7, marginTop: i > 0 ? 5 : 0 }}>
          <span style={{ width: 8, height: 8, borderRadius: 3, background: (colorMap && colorMap[p.dataKey]) || p.color || p.fill, flexShrink: 0 }} />
          <span style={{ color: C.ink, flex: 1, whiteSpace: "nowrap" }}>{p.name}</span>
          <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.ink, fontWeight: 700 }}>{fmtBRL(Number(p.value) || 0)}</span>
          {mostrarPct && sum > 0 && <span style={{ color: C.faint, fontSize: 10.5, minWidth: 30, textAlign: "right" }}>{((Number(p.value) / sum) * 100).toFixed(0)}%</span>}
        </div>
      ))}
      {extra && (
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, marginTop: 8, paddingTop: 7, borderTop: `1px solid ${C.border}` }}>
          <span style={{ color: C.muted, fontWeight: 600 }}>{extra.rotulo}</span>
          <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: extra.cor }}>{extra.texto}</span>
        </div>
      )}
    </div>
  );
}

function Field({ label, children, span }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 5, fontSize: 13, ...(span ? { gridColumn: `span ${span}` } : {}) }}>
      <span style={{ color: C.muted, fontWeight: 600, letterSpacing: 0.2 }}>{label}</span>
      {children}
    </label>
  );
}

const inputStyle = {
  border: "1px solid " + C.border, borderRadius: 11, padding: "12px 13px",
  fontSize: 14.5, fontFamily: "'Inter', sans-serif", background: C.card,
  color: C.ink, outline: "none",
};

// Native date/month inputs render taller than text inputs on iPad Safari, which
// threw the form rows and the filter labels out of line; pin them to the same height.
const dateInputFix = { textAlign: "left", height: 46, boxSizing: "border-box", WebkitAppearance: "none", appearance: "none", minWidth: 0 };
function TextInput(props) {
  const ehData = props.type === "date" || props.type === "month";
  return <input {...props} style={{ ...inputStyle, ...(ehData ? dateInputFix : {}), ...(props.style||{}) }} />;
}
function TextArea({ rows = 3, ...props }) {
  return (
    <textarea
      rows={rows}
      {...props}
      style={{ ...inputStyle, resize: "vertical", fontFamily: "'Inter', sans-serif", lineHeight: 1.45, ...(props.style || {}) }}
    />
  );
}
function SelectInput({ options, ...props }) {
  return (
    <select {...props} style={{ ...inputStyle, ...(props.style||{}) }}>
      {options.map(o => <option key={o} value={o}>{o}</option>)}
    </select>
  );
}

// Confirmação com o visual do app (no lugar do confirm nativo). Sem o ConfirmHost montado, cai no confirm nativo.
let _abrirConfirm = null;
const confirmar = (mensagem, opcoes = {}) => (_abrirConfirm ? _abrirConfirm({ mensagem, ...opcoes }) : Promise.resolve(window.confirm(mensagem)));
function ConfirmHost() {
  const [cfg, setCfg] = useState(null);
  const ref = useRef(null);
  useEffect(() => {
    _abrirConfirm = (c) => new Promise(res => {
      if (ref.current) ref.current.res(false);
      ref.current = { ...c, res };
      setCfg(ref.current);
    });
    return () => { _abrirConfirm = null; };
  }, []);
  const fechar = (v) => { const c = ref.current; ref.current = null; setCfg(null); if (c) c.res(v); };
  useEffect(() => {
    if (!cfg) return;
    const onKey = (e) => { if (e.key === "Escape") fechar(false); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [cfg]);
  if (!cfg) return null;
  return (
    <div onClick={() => fechar(false)} style={{ position: "fixed", inset: 0, zIndex: 600, background: "#0A0E1Cb8", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div onClick={e => e.stopPropagation()} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 20, maxWidth: 420, width: "100%", padding: "22px 22px 18px", boxShadow: "0 20px 50px #0008", animation: "slideUp .2s ease-out" }}>
        {cfg.titulo && <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19, color: C.ink, marginBottom: 10 }}>{cfg.titulo}</div>}
        <div style={{ whiteSpace: "pre-line", fontSize: 14, lineHeight: 1.55, color: C.ink, marginBottom: 20 }}>{cfg.mensagem}</div>
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button type="button" onClick={() => fechar(false)} style={{ border: `1px solid ${C.border}`, background: "transparent", color: C.muted, borderRadius: 12, padding: "10px 16px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>{cfg.cancelLabel || "Cancelar"}</button>
          <button type="button" autoFocus onClick={() => fechar(true)} style={{ border: "none", background: cfg.perigo ? C.coral : C.emerald, color: cfg.perigo ? "#fff" : "#0A0E1C", borderRadius: 12, padding: "10px 18px", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>{cfg.okLabel || "Confirmar"}</button>
        </div>
      </div>
    </div>
  );
}

// Se o app ficar aberto na virada do mês e a pessoa estiver vendo o mês que acabou, passa sozinho para o novo mês corrente.
function useMesCorrenteAuto(setMes) {
  const ref = useRef(currentMonth());
  useEffect(() => {
    const checar = () => {
      const atual = currentMonth();
      if (atual === ref.current) return;
      const antigo = ref.current;
      ref.current = atual;
      setMes(m => (m === antigo ? atual : m));
    };
    const t = setInterval(checar, 60000);
    document.addEventListener("visibilitychange", checar);
    return () => { clearInterval(t); document.removeEventListener("visibilitychange", checar); };
  }, []);
}

function Modal({ title, onClose, children, footer, headerExtra }) {
  // Enquanto há janela aberta, o "+" rápido some (a ponta dele aparecia na borda da janela).
  useEffect(() => {
    document.body.classList.add("modal-aberto");
    return () => document.body.classList.remove("modal-aberto");
  }, []);
  return (
    <div style={{
      position: "fixed", inset: 0, background: "#151A2E99", zIndex: 50,
      display: "flex", alignItems: "flex-end", justifyContent: "center",
      padding: "0 0 min(5vh, 36px)"
    }} onClick={onClose}>
      <div
        onClick={e => e.stopPropagation()}
        style={{
          background: C.card, width: "100%", maxWidth: "min(96vw, 1040px)", maxHeight: "90vh",
          borderRadius: 24, overflow: "hidden", display: "flex", flexDirection: "column",
          boxShadow: "0 -8px 30px #0004", animation: "slideUp .25s ease-out"
        }}
      >
        <div style={{ padding: "28px 28px 0", flexShrink: 0 }}>
          <div style={{ width: 36, height: 4, borderRadius: 4, background: C.border, margin: "0 auto 16px" }} />
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
            <h3 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontSize: 21, color: C.ink }}>{title}</h3>
            {headerExtra}
            <button onClick={onClose} style={{
              border: "none", background: C.border, borderRadius: "50%", width: 32, height: 32,
              display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: C.ink
            }}><X size={16} /></button>
          </div>
        </div>
        <div style={{ padding: "0 28px", overflowY: "auto", flex: 1 }}>
          {children}
        </div>
        {footer && (
          <div style={{
            flexShrink: 0, padding: "16px 28px calc(16px + env(safe-area-inset-bottom))",
            borderTop: `1px solid ${C.border}`, boxShadow: "0 -6px 16px rgba(0,0,0,0.25)"
          }}>
            {footer}
          </div>
        )}
        {!footer && <div style={{ height: 28, flexShrink: 0 }} />}
      </div>
    </div>
  );
}

function SummaryCard({ icon, label, value, tone = "ink", sub }) {
  const colors = { ink: C.ink, green: C.emerald, red: C.coral, gold: C.amber };
  const c = colors[tone];
  return (
    <div className="tap" style={{
      background: C.card, boxShadow: shadowSm, borderRadius: 16,
      padding: "16px 17px", flex: "1 1 140px", minWidth: 140
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, color: C.muted, fontSize: 11.5, fontWeight: 700, marginBottom: 9, textTransform: "uppercase", letterSpacing: 0.4 }}>
        {icon}{label}
      </div>
      <div style={{
        fontFamily: "'IBM Plex Mono', monospace", fontSize: 21, fontWeight: 700,
        color: c, fontVariantNumeric: "tabular-nums", letterSpacing: -0.3
      }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: C.faint, marginTop: 4 }}>{sub}</div>}
    </div>
  );
}

// One unified card with divided columns, instead of several separate boxes —
// this is the "single widget" pattern real fintech dashboards use for quick stats.
function StatRow({ stats }) {
  const colors = { ink: C.ink, green: C.emerald, red: C.coral, gold: C.amber };
  return (
    <div className="in tap" style={{
      background: C.card, boxShadow: shadowSm, borderRadius: 18,
      display: "flex", marginBottom: 20, overflow: "hidden"
    }}>
      {stats.map((s, i) => (
        <div key={i} style={{
          flex: 1, padding: "16px 14px", minWidth: 0,
          borderLeft: i > 0 ? `1px solid ${C.border}` : "none"
        }}>
          <div style={{
            display: "flex", alignItems: "center", gap: 5, color: C.muted, fontSize: 10.5, fontWeight: 700,
            marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.3, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis"
          }}>
            {s.icon}{s.label}
          </div>
          <div style={{
            fontFamily: "'IBM Plex Mono', monospace", fontSize: 17, fontWeight: 700,
            color: colors[s.tone || "ink"], fontVariantNumeric: "tabular-nums", letterSpacing: -0.3,
            whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis"
          }}>{s.value}</div>
          {s.trend && (
            <div style={{
              display: "flex", alignItems: "center", gap: 3, marginTop: 4, fontSize: 10.5, fontWeight: 700,
              color: s.trend.good ? C.emerald : C.coral
            }}>
              {s.trend.pct >= 0 ? <TrendingUp size={10} /> : <TrendingDown size={10} />}
              {Math.abs(s.trend.pct).toFixed(0)}% vs mês ant.
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

// Filters wrapped in one consistent card, instead of raw labels+selects floating
// directly on the page background.
function FilterBar({ children }) {
  return (
    <div style={{
      background: C.card, boxShadow: shadowSm, borderRadius: 16,
      padding: "14px 16px", marginBottom: 20, display: "flex", flexWrap: "wrap", gap: 14, alignItems: "flex-end"
    }}>
      {children}
    </div>
  );
}

const NAV_ITEMS = [
  { key: "dashboard", label: "Início", icon: TrendingUp },
  { key: "custos", label: "Custos", icon: Wallet },
  { key: "receitas", label: "Receitas", icon: PiggyBank },
  { key: "resumo", label: "Anual", icon: Calendar },
  { key: "demandas", label: "Tarefas", icon: ClipboardList },
];

// Global action, not a screen: reachable from any tab, opens a sheet on top
// of whatever the person is doing, and returns them there once confirmed.
function QuickAddFab({ onAddCusto, onAddReceita, onAddDemanda, demandas = [], onResolveDemanda, ocultarBotao = false }) {
  const [open, setOpen] = useState(false);
  // Atalho "Rápido" do Início abre o mesmo lançamento rápido por texto.
  useEffect(() => {
    const abrir = () => setOpen(true);
    window.addEventListener("pt-lancamento-rapido", abrir);
    return () => window.removeEventListener("pt-lancamento-rapido", abrir);
  }, []);
  const [step, setStep] = useState("input");
  const [text, setText] = useState("");
  const [draft, setDraft] = useState(null);
  const [saving, setSaving] = useState(false);

  const reset = () => { setOpen(false); setStep("input"); setText(""); setDraft(null); setSaving(false); };

  // draft is always an array — 1 item normally, 2+ when the sentence named a
  // split payment across different fontes (see parseQuickEntry). Each custo
  // entry also gets a matchedDemanda suggestion when its descrição clearly
  // corresponds to an open demanda (see matchDemanda) — pre-checked,
  // the person can untick it; ticked ones also mark that demanda Concluída on save.
  const analisar = async () => {
    if (!text.trim()) return;
    const parsed = parseQuickEntry(text).map(d => {
      if (d.tipo !== "custo") return d;
      const matched = matchDemanda(d.descricao, demandas);
      return matched ? { ...d, matchedDemandaId: matched.id, matchedDemandaNome: matched.nome, resolverDemanda: true } : d;
    });
    if (parsed.every(p => p.tipo !== "demanda" && !isValorValido(p.valor))) {
      const continuar = await confirmar(
        "Não encontrei nenhum valor no texto digitado. Quer continuar e preencher o valor manualmente na próxima tela?",
        { okLabel: "Continuar" }
      );
      if (!continuar) return;
    }
    // Amounts go into the editable field in Brazilian format ("2.675,00").
    setDraft(parsed.map(d => ({ ...d, valor: isValorValido(d.valor) ? fmtValorInput(d.valor) : "", prioridade: "Média" })));
    setStep("review");
  };

  const updateDraft = (i, patch) => setDraft(ds => ds.map((d, j) => (j === i ? { ...d, ...patch } : d)));
  // Removing the last remaining card leaves nothing to confirm, so close the
  // whole flow instead of stranding an empty review screen.
  const removeDraft = (i) => {
    const next = draft.filter((_, j) => j !== i);
    if (next.length === 0) reset(); else setDraft(next);
  };

  const salvar = async () => {
    if (saving || !draft || draft.length === 0) return;
    if (draft.some(d => !d.descricao || !d.descricao.trim())) { alert("Preencha a descrição de todos os lançamentos."); return; }
    if (draft.some(d => d.tipo !== "demanda" && !d.data)) { alert("Preencha a data de todos os lançamentos."); return; }
    const valores = draft.map(d => parseValorBR(d.valor));
    if (draft.some((d, i) => d.tipo !== "demanda" && !isValorValido(valores[i]))) {
      alert("Cada lançamento precisa de um valor maior que zero.");
      return;
    }
    if (draft[0].data && !(await confirmDateSanity(draft[0].data, draft[0].tipo === "custo" ? "data de vencimento" : draft[0].tipo === "demanda" ? "prazo" : "data de recebimento"))) return;
    setSaving(true);
    try {
      for (let i = 0; i < draft.length; i++) {
        const d = draft[i];
        const valorNum = valores[i];
        const mes = d.data.slice(0, 7);
        const anoNum = parseInt(d.data.slice(0, 4), 10);
        const status = d.status || "Pendente";
        if (d.tipo === "demanda") {
          await onAddDemanda({
            nome: d.descricao, contato: "",
            detalhes: d.hora ? `Horário: ${d.hora}` : "",
            prazo: d.data || "", prioridade: d.prioridade || "Média", status: "Pendente", obs: "", criadoEm: todayISO(), concluidaEm: ""
          });
        } else if (d.tipo === "custo") {
          await onAddCusto({
            mes, ano: anoNum, descricao: d.descricao,
            vencimento: d.data, dataPagamento: status === "Pago" ? d.data : "",
            valorPrevisto: valorNum, valorPago: status === "Pago" ? valorNum : null,
            fonte: d.fonte || FONTES[0], status, obs: ""
          });
          if (d.matchedDemandaId && d.resolverDemanda && onResolveDemanda) {
            await onResolveDemanda(d.matchedDemandaId);
          }
        } else {
          await onAddReceita({
            mes, ano: anoNum, descricao: d.descricao, dataRecebimento: d.data,
            valorPrevisto: valorNum, valorRecebido: status === "Recebido" ? valorNum : null,
            categoria: d.categoria || "", contaDestino: d.fonte || destinoPadraoReceita(d.descricao) || FONTES[1],
            status, obs: ""
          });
        }
      }
      reset();
    } finally {
      setSaving(false);
    }
  };

  // The button slides away while scrolling down (so it never covers a value or the
  // trash icon of the row underneath) and comes back as soon as the person scrolls up.
  const [escondido, setEscondido] = useState(false);
  useEffect(() => {
    let ultimo = window.scrollY || 0;
    const onScroll = () => {
      const y = window.scrollY || 0;
      if (y > ultimo + 8 && y > 80) setEscondido(true);
      else if (y < ultimo - 8 || y <= 80) setEscondido(false);
      ultimo = y;
    };
    window.addEventListener("scroll", onScroll, { passive: true });
    return () => window.removeEventListener("scroll", onScroll);
  }, []);

  return (
    <>
      {!ocultarBotao && <button className="tap fab-rapido" onClick={() => setOpen(true)} title="Lançamento rápido" aria-hidden={escondido} tabIndex={escondido ? -1 : 0} style={{
        position: "fixed", right: 8, bottom: "calc(88px + env(safe-area-inset-bottom))", zIndex: 42,
        width: 42, height: 42, borderRadius: 21, border: "none",
        background: C.emerald, color: "#0A0E1C", cursor: "pointer",
        display: "flex", alignItems: "center", justifyContent: "center",
        boxShadow: "0 6px 16px rgba(0,224,168,0.28)",
        transform: escondido ? "translateY(110px) scale(0.9)" : "translateY(0) scale(1)",
        opacity: escondido ? 0 : 0.82, pointerEvents: escondido ? "none" : "auto",
        transition: "transform .25s ease, opacity .2s ease"
      }}>
        <Plus size={21} strokeWidth={2.6} />
      </button>}

      {open && step === "input" && (
        <Modal title="Lançamento rápido" onClose={reset}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <span style={{ fontSize: 12.5, color: C.muted, lineHeight: 1.5 }}>
              Descreva em texto livre — ex: "aluguel 2675 dia 10" ou "recebi 800 freelance dia 15".
            </span>
            <TextArea rows={3} autoFocus value={text} onChange={e => setText(e.target.value)} placeholder="Digite aqui..." />
            <button className="tap" onClick={analisar} style={{ ...addBtnStyle, justifyContent: "center" }}>
              <Check size={16} /> Analisar
            </button>
          </div>
        </Modal>
      )}

      {open && step === "review" && draft && (
        <Modal
          title={draft.length > 1 ? `Confirmar ${draft.length} lançamentos` : "Confirmar lançamento"}
          onClose={reset}
          footer={
            <button type="button" onClick={salvar} disabled={saving} style={{ ...addBtnStyle, width: "100%", justifyContent: "center", opacity: saving ? 0.7 : 1, cursor: saving ? "default" : "pointer" }}>
              <Check size={16} /> {saving ? "Salvando..." : draft.length > 1 ? `Salvar os ${draft.length} lançamentos` : "Salvar"}
            </button>
          }
        >
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {draft.length > 1 && (
              <div style={{ fontSize: 12, color: C.muted, lineHeight: 1.5 }}>
                Encontrei {draft.length} valores com fontes diferentes nessa frase — conferi cada um separado, como um pagamento dividido. Confira antes de salvar.
              </div>
            )}
            {draft.map((d, i) => (
              <div key={i} style={{
                display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12,
                ...(draft.length > 1 ? { padding: 14, border: `1px solid ${C.border}`, borderRadius: 14 } : {})
              }}>
                {draft.length > 1 && (
                  <div style={{ gridColumn: "span 2", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontSize: 11.5, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: 0.3 }}>Lançamento {i + 1}</span>
                    <button type="button" onClick={() => removeDraft(i)} style={{ border: "none", background: "transparent", color: C.coral, cursor: "pointer", display: "flex", alignItems: "center", padding: 2 }} aria-label="Remover este lançamento">
                      <Trash2 size={15} />
                    </button>
                  </div>
                )}
                <Field label="Tipo">
                  <div style={{ display: "flex", gap: 8 }}>
                    {["custo", "receita", "demanda"].map(t => (
                      <button key={t} type="button" onClick={() => updateDraft(i, { tipo: t })} style={{
                        flex: 1, padding: "9px 0", borderRadius: 10, border: `1px solid ${d.tipo === t ? C.emerald : C.border}`,
                        background: d.tipo === t ? C.emeraldSoft : C.card, color: d.tipo === t ? C.emerald : C.muted,
                        fontWeight: 700, fontSize: 12.5, cursor: "pointer"
                      }}>{t === "custo" ? "Custo" : t === "receita" ? "Receita" : "Demanda"}</button>
                    ))}
                  </div>
                </Field>
                {d.tipo === "demanda"
                  ? <Field label="Horário (opcional)"><TextInput type="time" value={d.hora || ""} onChange={e => updateDraft(i, { hora: e.target.value })} /></Field>
                  : <Field label="Valor (R$)"><TextInput type="text" inputMode="decimal" placeholder="0,00" value={d.valor} onChange={e => updateDraft(i, { valor: e.target.value })} /></Field>}
                <Field label="Descrição" span={2}><TextInput value={d.descricao} onChange={e => updateDraft(i, { descricao: e.target.value })} /></Field>
                <Field label={d.tipo === "custo" ? "Vencimento" : d.tipo === "demanda" ? "Prazo / data" : "Recebimento"}>
                  <TextInput type="date" value={d.data} onChange={e => updateDraft(i, { data: e.target.value })} />
                </Field>
                {d.tipo === "demanda" ? (
                  <Field label="Prioridade" span={2}><SelectInput options={PRIORIDADES} value={d.prioridade || "Média"} onChange={e => updateDraft(i, { prioridade: e.target.value })} /></Field>
                ) : (<>
                <Field label={d.tipo === "custo" ? "Fonte de Pagamento" : "Conta de Destino"}>
                  <SelectInput options={opcoesFonte(d.fonte)} value={d.fonte || (d.tipo === "custo" ? FONTES[0] : (destinoPadraoReceita(d.descricao) || FONTES[1]))} onChange={e => updateDraft(i, { fonte: e.target.value })} />
                  {d.fonte === "Transferência" && <div style={{ fontSize: 11, color: C.amber, marginTop: 4 }}>Escolha de qual conta saiu: Transferência (Corrente) ou (Salário).</div>}
                </Field>
                <Field label="Status">
                  <SelectInput options={d.tipo === "custo" ? STATUS_CUSTO : STATUS_RECEITA} value={d.status || "Pendente"} onChange={e => updateDraft(i, { status: e.target.value })} />
                </Field>
                </>)}
                {d.matchedDemandaId && (
                  <label style={{
                    gridColumn: "span 2", display: "flex", alignItems: "center", gap: 8,
                    padding: "9px 12px", borderRadius: 10, background: C.emeraldSoft,
                    fontSize: 12.5, color: C.emerald, fontWeight: 600, cursor: "pointer"
                  }}>
                    <input type="checkbox" checked={!!d.resolverDemanda} onChange={e => updateDraft(i, { resolverDemanda: e.target.checked })} style={{ accentColor: C.emerald, width: 15, height: 15 }} />
                    Marcar a demanda "{d.matchedDemandaNome}" como Concluída
                  </label>
                )}
              </div>
            ))}
          </div>
        </Modal>
      )}
    </>
  );
}

function BottomNav({ tab, setTab }) {
  return (
    <nav style={{
      position: "fixed", bottom: 0, left: 0, right: 0, zIndex: 40,
      background: C.card, boxShadow: `0 -1px 0 ${C.border}, ${shadowNav}`,
      display: "flex", justifyContent: "space-around", padding: "10px 6px calc(10px + env(safe-area-inset-bottom))"
    }}>
      {NAV_ITEMS.map(item => {
        const Icon = item.icon;
        const active = tab === item.key;
        return (
          <button key={item.key} className="tap" onClick={() => setTab(item.key)} style={{
            display: "flex", flexDirection: "column", alignItems: "center", gap: 4,
            border: "none", background: "transparent", cursor: "pointer",
            padding: "4px 6px", flex: 1, maxWidth: 92
          }}>
            <div style={{
              display: "flex", alignItems: "center", justifyContent: "center",
              width: active ? 52 : 36, height: 30, borderRadius: 15,
              background: active ? C.emerald : "transparent",
              boxShadow: active ? `0 4px 12px ${C.emeraldGlow}` : "none",
              transition: "all .25s cubic-bezier(.34,1.56,.64,1)"
            }}>
              <Icon size={19} color={active ? "#fff" : C.faint} strokeWidth={active ? 2.5 : 2} />
            </div>
            <span style={{ fontSize: 10.5, fontWeight: active ? 800 : 600, color: active ? C.emerald : C.faint, transition: "color .2s" }}>
              {item.label}
            </span>
          </button>
        );
      })}
    </nav>
  );
}

function LoginScreen({ onLogin, error, loading }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");

  const submit = (e) => {
    e.preventDefault();
    onLogin(email, password);
  };

  return (
    <div style={{
      minHeight: "100vh", display: "flex", flexDirection: "column", justifyContent: "center",
      background: `linear-gradient(160deg, #0A0E1C 0%, #131A30 60%, #16203C 100%)`,
      fontFamily: "'Inter', sans-serif", padding: 24
    }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:wght@600;700&family=Inter:wght@400;500;600;700&display=swap');`}</style>
      <div style={{ marginBottom: 28, textAlign: "center" }}>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 14, color: C.lavender, marginBottom: 4 }}>Livro-Razão</div>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 28, fontWeight: 700, color: "#F3F5FA" }}>Pauta Tutti</div>
      </div>
      <form onSubmit={submit} style={{
        background: C.card, borderRadius: 20, padding: "26px 22px", width: "100%", maxWidth: 360,
        margin: "0 auto", display: "flex", flexDirection: "column", gap: 14, boxShadow: shadowMd
      }}>
        <Field label="E-mail">
          <TextInput type="email" autoFocus value={email} onChange={e => setEmail(e.target.value)} />
        </Field>
        <Field label="Senha">
          <TextInput type="password" value={password} onChange={e => setPassword(e.target.value)} />
        </Field>
        {error && <div style={{ fontSize: 12.5, color: C.coral, fontWeight: 600 }}>{error}</div>}
        <button type="submit" disabled={loading} style={{
          display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
          background: C.emerald, color: "#fff", border: "none", borderRadius: 12,
          padding: "13px 14px", fontSize: 14.5, fontWeight: 700, cursor: "pointer",
          opacity: loading ? 0.7 : 1, marginTop: 4
        }}>
          {loading ? "Entrando…" : "Entrar"}
        </button>
      </form>
    </div>
  );
}
export default function App() {
  useEffect(() => { document.title = "Pauta Tutti"; }, []);
  const [user, setUser] = useState(undefined); // undefined = checking, null = logged out
  const [authError, setAuthError] = useState("");
  const [authLoading, setAuthLoading] = useState(false);
  const [custos, setCustos] = useState([]);
  const [receitas, setReceitas] = useState([]);
  const [demandas, setDemandas] = useState([]);
  const [fixos, setFixos] = useState([]);
  const [fixosErro, setFixosErro] = useState(false);
  const [loadedFlags, setLoadedFlags] = useState({ custos: false, receitas: false, demandas: false });
  const [tab, setTab] = useState("dashboard");
  const [temaPref, setTemaPref] = useState(lerTemaPref);
  const temaAtual = temaEfetivo(temaPref);
  aplicarTema(temaAtual);   // before anything renders, so every C.* below already reads the right palette
  const escolherTema = (v) => { setTemaPref(v); salvarTemaPref(v); };
  const [inicioContagem, setInicioContagem] = useState(lerInicioContagem);
  const mudarInicioContagem = (v) => { setInicioContagem(v); salvarInicioContagem(v); };
  const [jumpTarget, setJumpTarget] = useState(null);
  const goToItem = (tab, id, extra = {}) => { setTab(tab); setJumpTarget({ tab, id, ...extra }); };
  const [saveError, setSaveError] = useState(false);
  const fileInputRef = useRef(null);
  const [menuOpen, setMenuOpen] = useState(false);

  const loading = !(loadedFlags.custos && loadedFlags.receitas && loadedFlags.demandas);
  const data = useMemo(() => ({ custos, receitas, demandas }), [custos, receitas, demandas]);
  // Dia de hoje: se o app passa a noite aberto (ou volta do segundo plano), tudo o que depende da data é recalculado.
  // Com uma janela aberta (formulário), a troca espera para não perder o que está sendo digitado.
  const [diaHoje, setDiaHoje] = useState(localISO);
  const [, setHoraHoje] = useState(() => new Date().getHours());
  useEffect(() => {
    const sync = () => {
      if (document.body && document.body.classList.contains("modal-aberto")) return;
      const agora = localISO();
      setDiaHoje(prev => (prev === agora ? prev : agora));
      setHoraHoje(new Date().getHours());
    };
    const t = setInterval(sync, 60000);
    document.addEventListener("visibilitychange", sync);
    window.addEventListener("focus", sync);
    window.addEventListener("pageshow", sync);
    return () => {
      clearInterval(t);
      document.removeEventListener("visibilitychange", sync);
      window.removeEventListener("focus", sync);
      window.removeEventListener("pageshow", sync);
    };
  }, []);
  const central = useCentralAvisos(custos, demandas, fixos, 14, diaHoje);
  const nAtrasadosApp = central.nUrgentes;
  // Central de avisos: abre por cima da aba atual (sino do topo) ou a partir do Início. null = fechada; { hidden } = valores ocultos?
  const [centralAberta, setCentralAberta] = useState(null);
  const receitasRenda = useMemo(() => receitas.filter(r => !ehDeposito(r)), [receitas]);
  const depositosRec = useMemo(() => receitas.filter(ehDeposito), [receitas]);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => setUser(u));
    return () => unsub();
  }, []);

  const handleLogin = async (email, password) => {
    setAuthError("");
    setAuthLoading(true);
    try {
      await signInWithEmailAndPassword(auth, email, password);
    } catch (e) {
      setAuthError("E-mail ou senha incorretos.");
    } finally {
      setAuthLoading(false);
    }
  };

  const handleLogout = () => signOut(auth);

  // Real-time listeners — each collection reflects exactly what's in
  // Firestore. Nothing here ever writes; reading never has side effects.
  useEffect(() => {
    if (!user) return;
    const toRow = (d) => ({ id: d.id, ...d.data() });
    const unsubC = onSnapshot(custosCol, snap => {
      setCustos(snap.docs.map(toRow));
      setLoadedFlags(f => ({ ...f, custos: true }));
    }, () => setSaveError(true));
    const unsubR = onSnapshot(receitasCol, snap => {
      setReceitas(snap.docs.map(toRow));
      setLoadedFlags(f => ({ ...f, receitas: true }));
    }, () => setSaveError(true));
    const unsubD = onSnapshot(demandasCol, snap => {
      setDemandas(snap.docs.map(toRow));
      setLoadedFlags(f => ({ ...f, demandas: true }));
    }, () => setSaveError(true));
    // Fixos list: its own listener; a permission error here must not raise the app-wide save error.
    const unsubF = onSnapshot(fixosCol, snap => { setFixos(snap.docs.map(toRow)); setFixosErro(false); }, () => setFixosErro(true));
    return () => { unsubC(); unsubR(); unsubD(); unsubF(); };
  }, [user]);

  // Every write below touches only the single document it concerns —
  // never the whole collection — so one action can never overwrite
  // or erase everything else.
  const addCusto = async (c) => {
    try { await addDoc(custosCol, c); setSaveError(false); } catch (e) { setSaveError(true); }
  };
  const removeCusto = async (id) => {
    try { await deleteDoc(doc(db, "custos", id)); setSaveError(false); } catch (e) { setSaveError(true); }
  };
  const updateCusto = async (id, c) => {
    try { await updateDoc(doc(db, "custos", id), c); setSaveError(false); } catch (e) { setSaveError(true); }
  };
  const avisoFixos = () => alert("Não consegui gravar a lista de fixos. Provavelmente falta liberar a coleção \"recorrentes\" nas regras do Firestore.");
  const addFixo = async (x) => { try { await addDoc(fixosCol, x); } catch (e) { avisoFixos(); } };
  const updateFixo = async (id, x) => { try { await updateDoc(doc(db, "recorrentes", id), x); } catch (e) { avisoFixos(); } };
  const removeFixo = async (id) => { try { await deleteDoc(doc(db, "recorrentes", id)); } catch (e) { avisoFixos(); } };
  const addReceita = async (r) => {
    try { await addDoc(receitasCol, r); setSaveError(false); } catch (e) { setSaveError(true); }
  };
  const removeReceita = async (id) => {
    try { await deleteDoc(doc(db, "receitas", id)); setSaveError(false); } catch (e) { setSaveError(true); }
  };
  const updateReceita = async (id, r) => {
    try { await updateDoc(doc(db, "receitas", id), r); setSaveError(false); } catch (e) { setSaveError(true); }
  };
  const addDemanda = async (d) => {
    try { await addDoc(demandasCol, d); setSaveError(false); } catch (e) { setSaveError(true); }
  };
  const removeDemanda = async (id) => {
    try { await deleteDoc(doc(db, "demandas", id)); setSaveError(false); } catch (e) { setSaveError(true); }
  };
  const updateDemandaStatus = async (id, status) => {
    // Remember the day a demanda was concluded (only when the person moves it there);
    // moving it back out clears the date.
    const extra = status === "Concluída" ? { concluidaEm: todayISO() } : { concluidaEm: "" };
    try { await updateDoc(doc(db, "demandas", id), { status, ...extra }); setSaveError(false); } catch (e) { setSaveError(true); }
  };
  const updateDemanda = async (id, d) => {
    try { await updateDoc(doc(db, "demandas", id), d); setSaveError(false); } catch (e) { setSaveError(true); }
  };

  const handleImportClick = () => fileInputRef.current?.click();
  const handleFileChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const parsed = await importFromExcel(file);
      const nArq = parsed.custos.length + parsed.receitas.length + parsed.demandas.length;
      if (nArq === 0) {
        alert("Este arquivo não tem custos, receitas nem demandas no formato do app. Nada foi alterado.");
        return;
      }
      const hasExisting = custos.length || receitas.length || demandas.length;
      if (hasExisting) {
        const menos = parsed.custos.length < custos.length || parsed.receitas.length < receitas.length || parsed.demandas.length < demandas.length;
        const resp = window.prompt(
          `ATENÇÃO: isso vai SUBSTITUIR todos os dados atuais pelos do arquivo.\n\n` +
          `Atual: ${custos.length} custos, ${receitas.length} receitas, ${demandas.length} demandas\n` +
          `Arquivo: ${parsed.custos.length} custos, ${parsed.receitas.length} receitas, ${parsed.demandas.length} demandas\n` +
          (menos ? `\nO arquivo tem MENOS registros do que o atual. Confira se é o arquivo certo.\n` : "") +
          `\nAntes de substituir, um Excel com os dados atuais (Backup_antes_de_importar) será baixado.\n\nPara continuar, digite SUBSTITUIR:`
        );
        if (!resp || resp.trim().toUpperCase() !== "SUBSTITUIR") return;
        // Safety copy of what is there now; if it cannot be produced, nothing is touched.
        exportToExcel(data, "Backup_antes_de_importar");
      }
      // New records are written first and the old ones removed afterwards, in small batches
      // (Firestore allows at most 500 operations per batch). If something fails midway, the
      // worst case is duplicates, never lost data.
      const gravar = [
        ...parsed.custos.map(c => { const { id, ...rest } = c; return (b) => b.set(doc(custosCol), rest); }),
        ...parsed.receitas.map(r => { const { id, ...rest } = r; return (b) => b.set(doc(receitasCol), rest); }),
        ...parsed.demandas.map(d => { const { id, ...rest } = d; return (b) => b.set(doc(demandasCol), rest); }),
      ];
      const apagar = [
        ...custos.map(c => (b) => b.delete(doc(db, "custos", c.id))),
        ...receitas.map(r => (b) => b.delete(doc(db, "receitas", r.id))),
        ...demandas.map(d => (b) => b.delete(doc(db, "demandas", d.id))),
      ];
      const emLotes = async (ops) => {
        for (let i = 0; i < ops.length; i += 400) {
          const batch = writeBatch(db);
          ops.slice(i, i + 400).forEach(op => op(batch));
          await batch.commit();
        }
      };
      let gravou = false;
      try {
        await emLotes(gravar);
        gravou = true;
        await emLotes(apagar);
      } catch (err2) {
        alert(gravou
          ? "Os dados novos foram gravados, mas não foi possível remover todos os antigos. Pode haver registros duplicados: confira antes de repetir. O backup foi baixado."
          : "A importação foi interrompida. Os dados antigos continuam intactos; alguns registros novos podem ter sido adicionados, então confira antes de repetir.");
        return;
      }
      setSaveError(false);
    } catch (err) {
      alert("Não foi possível importar este arquivo. Confirme se é um Excel exportado por este app.");
    } finally {
      e.target.value = "";
    }
  };

  if (user === undefined) {
    return (
      <div style={{ minHeight: "100vh", background: C.page }} />
    );
  }
  if (!user) {
    return (
      <LoginScreen onLogin={handleLogin} error={authError} loading={authLoading} />
    );
  }

  if (loading) {
    return (
      <div style={{
        minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
        background: C.page, fontFamily: "'Inter', sans-serif", color: C.muted
      }}>
        Carregando…
      </div>
    );
  }

  // Personal single-user app: the e-mail prefix ("robertoauad") read badly as a
  // greeting, so the name is set directly.
  const greeting = "Roberto";
  const hojeLabel = `${new Date().getDate()} de ${MESES[new Date().getMonth()].toLowerCase()}`;
  const horaAgora = new Date().getHours();
  const diaSemana = ["domingo","segunda-feira","terça-feira","quarta-feira","quinta-feira","sexta-feira","sábado"][new Date().getDay()];
  const diaSemanaCap = diaSemana.charAt(0).toUpperCase() + diaSemana.slice(1);
  const saudacao = horaAgora < 5 ? "Boa noite" : horaAgora < 12 ? "Bom dia" : horaAgora < 18 ? "Boa tarde" : "Boa noite";

  // Conteúdo do topo (saudação, título, tema e menu): no Início ele vira a parte de cima do cartão de Saldo.
  // Botões de tema e menu (iguais no herói escuro e na barra fina das outras abas).
  const botoesApp = (claro) => {
    const corBorda = claro ? C.border : "rgba(255,255,255,0.18)";
    const corFundo = claro ? C.card : "rgba(255,255,255,0.08)";
    const corIcone = claro ? C.inkSoft : "#D5DBF5";
    return (
          <div style={{ position: "relative", display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
            {claro && (
              <button className="tap" onClick={() => setCentralAberta({ hidden: false })} aria-label={nAtrasadosApp > 0 ? `Central de avisos: ${nAtrasadosApp} precisam de atenção` : "Central de avisos: tudo em dia"}
                style={{ position: "relative", width: 38, height: 38, borderRadius: 12, border: `1px solid ${corBorda}`, background: corFundo, color: nAtrasadosApp > 0 ? C.coral : corIcone, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Bell size={16} />
                {nAtrasadosApp > 0 && <span style={{ position: "absolute", top: -5, right: -5, minWidth: 17, height: 17, borderRadius: 9, background: C.coral, color: "#fff", fontSize: 10, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", padding: "0 4px", boxSizing: "border-box" }}>{nAtrasadosApp > 99 ? "99+" : nAtrasadosApp}</span>}
              </button>
            )}
            <button className="tap" onClick={() => window.dispatchEvent(new Event("pt-lancamento-rapido"))} aria-label="Lançamento rápido"
              style={{ height: 38, borderRadius: 12, border: `1px solid ${claro ? C.emerald : "rgba(0,224,168,0.45)"}`, background: claro ? C.emeraldSoft : "rgba(0,224,168,0.12)", color: claro ? C.emerald : "#4FE8BE", cursor: "pointer", display: "flex", alignItems: "center", gap: 6, padding: "0 13px", fontSize: 12.5, fontWeight: 700, fontFamily: "inherit" }}>
              <Plus size={16} strokeWidth={2.6} /> <span className="rapido-txt">Rápido</span>
            </button>
            <button className="tap" onClick={() => setMenuOpen(o => !o)} aria-label="Menu"
              style={{ width: 38, height: 38, borderRadius: 12, border: `1px solid ${corBorda}`, background: corFundo, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <span style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                <span style={{ width: 4, height: 4, borderRadius: 2, background: corIcone }} />
                <span style={{ width: 4, height: 4, borderRadius: 2, background: corIcone }} />
                <span style={{ width: 4, height: 4, borderRadius: 2, background: corIcone }} />
              </span>
            </button>
            {menuOpen && (
              <>
                <div onClick={() => setMenuOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 45 }} />
                <div style={{
                  position: "absolute", right: 0, top: 46, background: C.card, borderRadius: 14,
                  boxShadow: shadowMd, padding: 6, zIndex: 46, width: 210, animation: "fadeIn .15s ease-out"
                }}>
                  <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileChange} style={{ display: "none" }} />
                  <button onClick={() => { setMenuOpen(false); handleImportClick(); }} style={menuItemStyle}>
                    <Download size={15} style={{ transform: "rotate(180deg)" }} /> Importar Excel
                  </button>
                  <button onClick={() => { setMenuOpen(false); exportToExcel(data); }} style={menuItemStyle}>
                    <Download size={15} /> Exportar Excel
                  </button>
                  <div style={{ height: 1, background: C.border, margin: "4px 6px" }} />
                  {[["escuro", "Tema escuro", Moon], ["claro", "Tema claro", Sun], ["auto", "Tema automático", Sun]].map(([v, rot, Ic]) => (
                    <button key={v} onClick={() => { setMenuOpen(false); escolherTema(v); }} style={menuItemStyle}>
                      <Ic size={15} /> {rot}
                      {temaPref === v && <Check size={15} color={C.emerald} style={{ marginLeft: "auto" }} />}
                    </button>
                  ))}
                  <div style={{ height: 1, background: C.border, margin: "4px 6px" }} />
                  <button onClick={() => { setMenuOpen(false); handleLogout(); }} style={{ ...menuItemStyle, color: C.coral }}>
                    <X size={15} /> Sair
                  </button>
                </div>
              </>
            )}
          </div>
    );
  };

  const topoApp = (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 12.5, color: "#B8C2E8", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.6 }}>
              {saudacao}, {greeting} · {hojeLabel}
            </div>
            <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: "clamp(30px, 9vw, 38px)", letterSpacing: -1, color: "#F3F5FA", lineHeight: 1, marginTop: 6 }}>
              Pauta Tutti
            </div>
          </div>
          {botoesApp(false)}
        </div>
  );

  return (
    <div key={temaAtual} style={{ minHeight: "100vh", background: C.page, fontFamily: "'Inter', sans-serif", paddingBottom: 150 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:wght@400;500;600;700&family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600;700&display=swap');
        * { box-sizing: border-box; }
        body { margin: 0; background: ${C.page}; }
        :root { color-scheme: ${temaAtual === "claro" ? "light" : "dark"}; }
        body.modal-aberto .fab-rapido { opacity: 0 !important; pointer-events: none !important; }
        @keyframes slideUp { from { transform: translateY(24px); opacity: 0 } to { transform: translateY(0); opacity: 1 } }
        @keyframes fadeIn { from { opacity: 0 } to { opacity: 1 } }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 4px; }
        button:focus-visible, select:focus-visible, input:focus-visible { outline: 2px solid ${C.emerald}; outline-offset: 1px; }
        .tap { transition: transform .15s cubic-bezier(.34,1.56,.64,1), box-shadow .15s ease, filter .15s ease; }
        @media (hover: hover) and (pointer: fine) {
          .tap:hover { filter: brightness(1.08); transform: translateY(-1px); }
        }
        .tap:active { transform: scale(0.96); }
        @keyframes ptPulse { 0% { transform: scale(1); } 30% { transform: scale(1.14); } 60% { transform: scale(1); } 80% { transform: scale(1.08); } 100% { transform: scale(1); } }
        @keyframes fadeInUp { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        .in { animation: fadeInUp .4s cubic-bezier(.22,1,.36,1) backwards; }
        @media (max-width: 440px) { .rapido-txt { display: none; } }
        .grid2 { display: grid; grid-template-columns: 1fr; gap: 20px; }
        @media (min-width: 680px) { .grid2 { grid-template-columns: 1fr 1fr; } .grid2 > * { margin-bottom: 0 !important; } }
        @media (min-width: 1100px) { .grid2 { grid-template-columns: 1fr 1fr 1fr; } }
        input[type="date"]::-webkit-date-and-time-value, input[type="month"]::-webkit-date-and-time-value { text-align: left; }
        .grid2c { display: grid; grid-template-columns: 1fr; gap: 20px; }
        @media (min-width: 680px) { .grid2c { grid-template-columns: 1fr 1fr; } .grid2c > * { margin-bottom: 0 !important; } }
        @media (min-width: 760px) { .kanban { grid-template-columns: var(--kcols, 1fr 1fr 1fr) !important; } }
      `}</style>

      {tab !== "dashboard" && (
      <div style={{ padding: "14px 20px 0" }}>
        <header style={{ position: "relative", maxWidth: 1080, margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, padding: "13px 14px 13px 18px", borderRadius: 20, background: C.card, border: `1px solid ${C.border}`, boxShadow: shadowSm }}>
          <div aria-hidden="true" style={{ position: "absolute", inset: 0, overflow: "hidden", borderRadius: "inherit", pointerEvents: "none" }}>
            <svg width="100%" height="100%" style={{ position: "absolute", inset: 0 }}>
              {[0, 1, 2, 3, 4].map(i => <line key={i} x1="0" x2="100%" y1={`${22 + i * 14}%`} y2={`${22 + i * 14}%`} stroke={C.lavender} strokeOpacity="0.14" strokeWidth="1" />)}
              {[["46%", 3, 0], ["53%", 2, 1], ["60%", 4, 0], ["67%", 1, 1]].map(([x, linha, inv], k) => (
                <g key={k} fill={C.emerald} fillOpacity="0.22" stroke={C.emerald} strokeOpacity="0.22">
                  <ellipse cx={x} cy={`${22 + linha * 14}%`} rx="4.5" ry="3.3" />
                </g>
              ))}
            </svg>
            <div style={{ position: "absolute", top: -50, right: -40, width: 150, height: 150, borderRadius: "50%", background: `radial-gradient(circle, ${C.emeraldGlow} 0%, transparent 70%)`, opacity: 0.6 }} />
          </div>
          <div style={{ position: "relative", display: "flex", flexDirection: "column", gap: 3, minWidth: 0 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ width: 10, height: 10, borderRadius: 5, background: "linear-gradient(135deg, #00E0A8, #9AA5D1)", flexShrink: 0 }} />
              <span style={{ fontSize: 15, fontWeight: 800, letterSpacing: 1.3, whiteSpace: "nowrap", flexShrink: 0, textTransform: "uppercase", background: `linear-gradient(90deg, ${C.emerald}, ${C.lavender})`, WebkitBackgroundClip: "text", backgroundClip: "text", color: "transparent", WebkitTextFillColor: "transparent" }}>Pauta Tutti</span>
            </div>
            <span style={{ paddingLeft: 20, fontSize: 12, color: C.muted, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{diaSemanaCap} · {hojeLabel}</span>
          </div>
          <div style={{ position: "relative" }}>{botoesApp(true)}</div>
        </header>
      </div>
      )}

      {centralAberta && (
        <CentralAvisosModal central={central} hidden={centralAberta.hidden} onClose={() => setCentralAberta(null)}
          onNavigate={goToItem} onConcluirDemanda={(id) => updateDemandaStatus(id, "Concluída")} />
      )}

      <main style={{ maxWidth: 1080, margin: "0 auto", padding: "20px 20px 0" }}>
        {tab === "dashboard" && <DashboardTab key={diaHoje} custos={data.custos} receitas={receitasRenda} depositos={depositosRec} demandas={data.demandas} onNavigate={goToItem} onConcluirDemanda={(id) => updateDemandaStatus(id, "Concluída")} inicioContagem={inicioContagem} fixos={fixos} topo={topoApp} onAbrirCentral={(h) => setCentralAberta({ hidden: !!h })} central={central} />}
        {tab === "custos" && <CustosTab key={diaHoje} fixos={fixos} fixosErro={fixosErro} onAddFixo={addFixo} onUpdateFixo={updateFixo} onRemoveFixo={removeFixo} custos={data.custos} onAdd={addCusto} onRemove={removeCusto} onUpdate={updateCusto} jumpTarget={jumpTarget && jumpTarget.tab === "custos" ? jumpTarget : null} onConsumeJump={() => setJumpTarget(null)} />}
        {tab === "receitas" && <ReceitasTab key={diaHoje} receitas={data.receitas} onAdd={addReceita} onRemove={removeReceita} onUpdate={updateReceita} inicioContagem={inicioContagem} jumpTarget={jumpTarget && jumpTarget.tab === "receitas" ? jumpTarget : null} onConsumeJump={() => setJumpTarget(null)} />}
        {tab === "resumo" && <ResumoTab key={diaHoje} custos={data.custos} receitas={receitasRenda} inicioContagem={inicioContagem} onChangeInicio={mudarInicioContagem} />}
        {tab === "demandas" && <DemandasTab key={diaHoje} demandas={data.demandas} onAdd={addDemanda} onRemove={removeDemanda} onUpdate={updateDemanda} onStatusChange={updateDemandaStatus} jumpTarget={jumpTarget && jumpTarget.tab === "demandas" ? jumpTarget : null} onConsumeJump={() => setJumpTarget(null)} />}
      </main>

      <BottomNav tab={tab} setTab={setTab} />
      <ConfirmHost />
      <QuickAddFab onAddCusto={addCusto} onAddReceita={addReceita} onAddDemanda={addDemanda} demandas={data.demandas} onResolveDemanda={(id) => updateDemandaStatus(id, "Concluída")} ocultarBotao={true} />

      {saveError && (
        <div style={{
          position: "fixed", bottom: 84, left: "50%", transform: "translateX(-50%)",
          background: C.coral, color: "#fff", padding: "8px 16px", borderRadius: 10,
          fontSize: 13, display: "flex", alignItems: "center", gap: 6, boxShadow: shadowMd, zIndex: 41
        }}>
          <AlertCircle size={15} /> Não foi possível salvar. Tente novamente.
        </div>
      )}
    </div>
  );
}

const menuItemStyle = {
  display: "flex", alignItems: "center", gap: 9, width: "100%", textAlign: "left",
  border: "none", background: "transparent", cursor: "pointer", borderRadius: 9,
  padding: "10px 10px", fontSize: 13.5, fontWeight: 600, color: C.ink
};

/* ------------------------------------------------------------------ */
/*  Dashboard                                                          */
/* ------------------------------------------------------------------ */

const MESES_ABR = MESES.map(m => m.slice(0, 3));

// Straight (not angled) axis label that wraps onto up to 3 short lines, so long
// names like "ALUGUEL VIA IPIRANGA" stay readable and distinguishable.
function WrapTick({ x, y, payload, fmt, max = 11 }) {
  const texto = fmt ? fmt(String(payload.value)) : String(payload.value);
  const linhas = [];
  let atual = "";
  texto.split(/\s+/).filter(Boolean).forEach(w => {
    if (atual && (atual + " " + w).length > max) { linhas.push(atual); atual = w; }
    else atual = atual ? atual + " " + w : w;
  });
  if (atual) linhas.push(atual);
  const vis = linhas.slice(0, 3);
  if (linhas.length > 3) vis[2] = vis[2].slice(0, max - 1) + "…";
  return (
    <text x={x} y={y + 10} textAnchor="middle" fontSize={9.5} fill={C.faint}>
      {vis.map((l, i) => <tspan key={i} x={x} dy={i === 0 ? 0 : 11}>{l}</tspan>)}
    </text>
  );
}

// Seletor de mês discreto: folha pequena com os 12 meses, setas de ano e "Hoje".
function SeletorMes({ mes, onPick, onClose }) {
  const [ano, setAno] = useState(parseInt(mes.split("-")[0], 10));
  const hoje = currentMonth();
  const pick = (m) => { onPick(m); onClose(); };
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 550, background: "#0A0E1Cb0", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div onClick={e => e.stopPropagation()} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 20, width: "100%", maxWidth: 340, padding: "16px 16px 14px", boxShadow: "0 20px 50px #0008", animation: "slideUp .2s ease-out" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
          <button type="button" onClick={() => setAno(a => a - 1)} style={monthNavBtnStyle} aria-label="Ano anterior">‹</button>
          <span style={{ fontFamily: "'Fraunces', serif", fontSize: 18, fontWeight: 600, color: C.ink }}>{ano}</span>
          <button type="button" onClick={() => setAno(a => a + 1)} style={monthNavBtnStyle} aria-label="Próximo ano">›</button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
          {MESES_ABR.map((nome, i) => {
            const key = `${ano}-${String(i + 1).padStart(2, "0")}`;
            const sel = key === mes;
            return (
              <button key={key} type="button" onClick={() => pick(key)} style={{
                position: "relative", height: 42, borderRadius: 12, cursor: "pointer", fontSize: 13.5, fontWeight: sel ? 800 : 600,
                border: `1px solid ${sel ? C.emerald : C.border}`, background: sel ? C.emeraldSoft : "transparent", color: sel ? C.emerald : C.ink
              }}>
                {nome}
                {key === hoje && <span style={{ position: "absolute", bottom: 5, left: "50%", marginLeft: -2, width: 4, height: 4, borderRadius: 2, background: C.emerald }} />}
              </button>
            );
          })}
        </div>
        <button type="button" onClick={() => pick(hoje)} style={{ width: "100%", marginTop: 12, height: 40, borderRadius: 12, border: `1px solid ${C.border}`, background: "transparent", color: C.muted, fontSize: 13.5, fontWeight: 700, cursor: "pointer" }}>Hoje</button>
      </div>
    </div>
  );
}

function BalanceHero({ pix, saldo, trendAtual, trendAnterior, mesLabel, hidden, onToggle, saldosPorMes, porFonte, expanded, onToggleExpanded, inicioIdx = 0, mesIdxVisto = new Date().getMonth(), ehMesCorrente = true, onPickMes, onAbrirSeletor, semLancamentos = false, mesAnteriorLabel = "", onVerAnterior, contasLinha = [], topo = null, situacao = [], curva = null, atencao = null }) {
  // Scale only by the months that count (from "contar a partir de" on), so older
  // months with incomplete history can't flatten the ones that matter.
  const maxAbs = Math.max(1, ...saldosPorMes.filter((_, i) => i >= inicioIdx).map(s => Math.abs(s)));
  // Trend compares "recebido - pago até o mesmo dia do mês" on both sides (not
  // the full previous month — see the trend calc in DashboardTab for why) and,
  // within that, % only makes intuitive sense against a positive base: "466%"
  // over a negative previous saldo reads as nonsense, so that case falls back
  // to a plain R$ difference instead.
  const saldoDiffBruto = trendAnterior !== null ? trendAtual - trendAnterior : null;
  const saldoDiff = saldoDiffBruto !== null && Math.abs(saldoDiffBruto) < 0.005 ? null : saldoDiffBruto; // sem diferença, sem rótulo
  const trendPct = (trendAnterior !== null && trendAnterior > 0) ? (saldoDiff / trendAnterior) * 100 : null;
  const estiloLinkAtencao = { border: "none", background: "transparent", padding: 0, fontSize: 12.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.6, display: "inline-flex", alignItems: "center", gap: 7, fontFamily: "inherit" };
  const linkAtencao = !atencao ? null : atencao.n > 0 ? (
    <button type="button" className="tap" onClick={atencao.onOpen} style={{ ...estiloLinkAtencao, cursor: "pointer", color: "#FF8A76", animation: atencao.pulsar ? "ptPulse 1.2s ease-out 1" : "none" }}>
      <span style={{ minWidth: 22, height: 22, borderRadius: 11, background: "#FF5A68", color: "#fff", fontSize: 12, display: "inline-flex", alignItems: "center", justifyContent: "center", padding: "0 6px", letterSpacing: 0 }}>{atencao.n}</span>
      {atencao.n === 1 ? "precisa" : "precisam"} de atenção ›
    </button>
  ) : (
    atencao.parados > 0 ? (
      <button type="button" className="tap" onClick={atencao.onOpen} style={{ ...estiloLinkAtencao, cursor: "pointer", color: "#4FE8BE" }}>
        Tudo em dia ✓ <span style={{ color: "#B8C2E8", fontWeight: 600, textTransform: "none", letterSpacing: 0 }}>· {atencao.parados} {atencao.parados === 1 ? "tarefa parada" : "tarefas paradas"} ›</span>
      </button>
    ) : <span style={{ ...estiloLinkAtencao, color: "#4FE8BE" }}>Tudo em dia ✓</span>
  );
  return (
    <div className="in" style={{
      background: `linear-gradient(150deg, #0B0F1F 0%, #101632 55%, #0F3D33 100%)`,
      borderRadius: 28, padding: topo ? "24px 24px 26px" : "32px 26px", marginBottom: 24,
      boxShadow: "0 20px 48px rgba(11,15,31,0.35), 0 6px 16px rgba(11,15,31,0.2)",
      position: "relative",
      ...(topo ? { border: "5px solid transparent", background: "linear-gradient(150deg, #0B0F1F 0%, #101632 55%, #0F3D33 100%) padding-box, linear-gradient(135deg, #00E0A8, #9AA5D1) border-box" } : {})
    }}>
      <div style={{ position: "absolute", inset: 0, overflow: "hidden", borderRadius: "inherit", pointerEvents: "none" }}>
        <div style={{
          position: "absolute", top: -70, right: -60, width: 220, height: 220, borderRadius: "50%",
          background: `radial-gradient(circle, ${C.emeraldGlow} 0%, transparent 70%)`
        }} />
      </div>
      {topo && <div style={{ position: "relative", zIndex: 3, marginBottom: 22, paddingBottom: 20, borderBottom: "1px solid rgba(255,255,255,0.12)" }}>{topo}</div>}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", position: "relative" }}>
        <div className="tap" onClick={onToggleExpanded} style={{ cursor: "pointer", flex: 1, minWidth: 0 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
            <span style={{ fontSize: 12.5, color: C.lavender, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.6 }}>Saldo de {mesLabel}</span>
            {saldoDiff !== null && (
              <span style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 10.5, fontWeight: 700, color: saldoDiff >= 0 ? "#4FE8BE" : "#FF8A76" }}>
                {saldoDiff >= 0 ? <TrendingUp size={11} /> : <TrendingDown size={11} />}
                {hidden ? "•••" : fmtBRL(Math.abs(saldoDiff))} {saldoDiff >= 0 ? "acima" : "abaixo"} do mês ant.{ehMesCorrente ? " até hoje" : ""}
              </span>
            )}
          </div>
          <div style={{
            fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 48, letterSpacing: -1.2,
            color: saldo < 0 ? "#FF8A76" : "#4FE8BE", fontVariantNumeric: "tabular-nums", lineHeight: 1.02
          }}>
            {hidden ? "R$ •••••" : fmtBRL(saldo)}
          </div>
          {!hidden && contasLinha.length > 1 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 12 }}>
              {contasLinha.map(c => (
                <span key={c.fonte} style={{ fontSize: 11.5, fontWeight: 600, fontVariantNumeric: "tabular-nums", borderRadius: 999, padding: "4px 10px", background: "rgba(255,255,255,0.07)", border: "1px solid rgba(255,255,255,0.1)", color: "#B8BFDB" }}>
                  {c.fonte === "Conta Salário" ? "CS" : c.fonte === "Conta Corrente" ? "CC" : c.fonte}{" "}
                  <strong style={{ color: c.saldo < 0 ? "#FF8A76" : "#4FE8BE" }}>{c.saldo < 0 ? "−" : "+"}{fmtBRL(Math.abs(c.saldo)).replace("R$", "").trim()}</strong>
                </span>
              ))}
            </div>
          )}
          {!hidden && situacao.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: "4px 14px", marginTop: 12, fontSize: 12, fontWeight: 600, fontVariantNumeric: "tabular-nums" }}>
              {situacao.map((x, i) => x.onClick
                ? <button key={i} type="button" className="tap" onClick={x.onClick} style={{ border: "none", background: "transparent", padding: 0, font: "inherit", color: x.cor, cursor: "pointer" }}>{x.t}</button>
                : <span key={i} style={{ color: x.cor }}>{x.t}</span>)}
            </div>
          )}
        </div>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 8, flexShrink: 0 }}>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="tap" onClick={onToggleExpanded} aria-label="Ver saldo por fonte" style={{
            width: 38, height: 38, borderRadius: 12, border: "1px solid rgba(255,255,255,0.12)", background: "rgba(255,255,255,0.06)",
            color: C.lavender, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
          }}>
            <ChevronDown size={17} style={{ transform: expanded ? "rotate(180deg)" : "none", transition: "transform .2s ease" }} />
          </button>
          <button className="tap" onClick={onToggle} style={{
            width: 38, height: 38, borderRadius: 12, border: "1px solid rgba(255,255,255,0.12)", background: "rgba(255,255,255,0.06)",
            color: C.lavender, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center"
          }}>
            {hidden ? <EyeOff size={16} /> : <Eye size={16} />}
          </button>
        </div>
        </div>
      </div>
      <div style={{ position: "relative", marginTop: 22 }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
          {MESES_ABR.map((m, i) => {
            if (i < inicioIdx || i > Math.max(mesIdxVisto, new Date().getMonth())) return null;
            const sel = i === mesIdxVisto;
            return (
              <button key={i} type="button" className="tap" onClick={() => onPickMes && onPickMes(i)} title={MESES[i]}
                style={{ border: `1px solid ${sel ? "#4FE8BE" : "rgba(255,255,255,0.14)"}`, background: sel ? "rgba(79,232,190,0.16)" : "rgba(255,255,255,0.05)", color: sel ? "#4FE8BE" : "#B8BFDB", borderRadius: 999, padding: "6px 13px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{m}</button>
            );
          })}
          <button type="button" className="tap" onClick={onAbrirSeletor} aria-label="Escolher outro mês"
            style={{ border: `1px solid ${ehMesCorrente ? "rgba(255,255,255,0.14)" : "rgba(255,255,255,0.14)"}`, background: "transparent", color: "#8B93B5", borderRadius: 999, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>Outro mês</button>
        </div>
        {curva && curva.rows.length > 1 && (
          <>
            <div style={{ height: 96, marginTop: 16 }}>
              <ResponsiveContainer width="100%" height="100%">
                <AreaChart data={curva.rows} margin={{ top: 8, right: 8, left: 8, bottom: 8 }}>
                  <defs>
                    <linearGradient id="gCurva" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stopColor={saldo < 0 ? "#FF8A76" : "#4FE8BE"} stopOpacity={0.35} /><stop offset="100%" stopColor={saldo < 0 ? "#FF8A76" : "#4FE8BE"} stopOpacity={0.02} />
                    </linearGradient>
                  </defs>
                  <XAxis dataKey="dia" hide />
                  <YAxis hide domain={["dataMin", "dataMax"]} />
                  <ReferenceLine y={0} stroke="rgba(255,255,255,0.2)" strokeDasharray="3 3" />
                  <Tooltip cursor={{ stroke: "rgba(255,255,255,0.25)", strokeWidth: 1 }} content={({ active, payload, label }) => (active && payload && payload.length) ? (
                    <div style={{ background: "#0B0F1F", border: "1px solid rgba(255,255,255,0.16)", borderRadius: 10, padding: "7px 10px", fontSize: 11.5, color: "#F3F5FA" }}>
                      <div style={{ color: "#8B93B5", fontWeight: 700, marginBottom: 2 }}>Dia {label}</div>
                      {payload.filter(x => x.value != null).map(x => (
                        <div key={x.dataKey} style={{ display: "flex", gap: 10, justifyContent: "space-between" }}>
                          <span style={{ color: "#B8BFDB" }}>{x.dataKey === "atual" ? "Este mês" : "Mês anterior"}</span>
                          <strong style={{ fontVariantNumeric: "tabular-nums" }}>{hidden ? "•••" : fmtBRL(x.value)}</strong>
                        </div>
                      ))}
                    </div>
                  ) : null} />
                  <Area type="stepAfter" dataKey="anterior" stroke="#7C86AD" strokeWidth={1.5} strokeDasharray="4 4" fill="none" dot={false} isAnimationActive={false} />
                  <Area type="stepAfter" dataKey="atual" stroke={saldo < 0 ? "#FF8A76" : "#4FE8BE"} strokeWidth={2.5} fill="url(#gCurva)" connectNulls={false} isAnimationActive={false}
                    dot={(p) => (p.payload && p.payload.dia === curva.hojeDia && p.cx != null)
                      ? <circle key="hoje" cx={p.cx} cy={p.cy} r={5} fill={saldo < 0 ? "#FF8A76" : "#4FE8BE"} stroke="#fff" strokeWidth={2} />
                      : <g key={"g" + p.index} />} />
                </AreaChart>
              </ResponsiveContainer>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginTop: 8 }}>
              <div style={{ display: "flex", gap: 16, fontSize: 10.5, color: "#7C86AD", fontWeight: 600 }}>
                <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 14, height: 2.5, borderRadius: 2, background: saldo < 0 ? "#FF8A76" : "#4FE8BE" }} />saldo dia a dia</span>
                <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 14, borderTop: "2px dashed #7C86AD" }} />mês anterior</span>
              </div>
              {linkAtencao}
            </div>
          </>
        )}
        {!(curva && curva.rows.length > 1) && linkAtencao && <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 14 }}>{linkAtencao}</div>}
      </div>

      {semLancamentos && (
        <div className="tap" onClick={onVerAnterior} style={{ marginTop: 14, textAlign: "center", fontSize: 12, color: "#7C86AD", cursor: "pointer", position: "relative" }}>
          Mês começando, sem lançamentos ainda · <span style={{ color: "#4FE8BE", fontWeight: 700 }}>ver {mesAnteriorLabel} ›</span>
        </div>
      )}

      {expanded && (
        <div className="in" style={{ marginTop: 22, paddingTop: 18, borderTop: "1px solid rgba(255,255,255,0.1)", position: "relative" }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: C.lavender, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 12 }}>
            Movimento do mês, por fonte
          </div>
          {porFonte.length === 0 ? (
            <div style={{ fontSize: 12.5, color: "#7C86AD" }}>Nenhuma fonte com movimento neste mês ainda.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              {porFonte.map(f => (
                <div key={f.fonte}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0" }}>
                    <span style={{ width: 9, height: 9, borderRadius: 3, background: fonteColor(f.fonte), flexShrink: 0 }} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: "#F3F5FA" }}>{f.fonte}</div>
                      <div style={{ fontSize: 10.5, color: "#7C86AD", marginTop: 1 }}>
                        Recebido {hidden ? "•••" : fmtBRL(f.recebido)} · {f.pendente > 0 ? "a pagar" : "pago"} {hidden ? "•••" : fmtBRL(f.pendente > 0 ? f.pendente : f.pago)}
                        {!hidden && f.meios && Object.keys(f.meios).length > 0 && ` · inclui ${Object.entries(f.meios).map(([m, v]) => `${m} ${fmtBRL(v)}`).join(" e ")}`}
                        {!hidden && f.depositoSaida > 0 && ` · depositado na conta ${fmtBRL(f.depositoSaida)}`}
                        {!hidden && f.depositoEntrada > 0 && ` · depósitos recebidos ${fmtBRL(f.depositoEntrada)}`}
                      </div>
                    </div>
                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 13.5, color: f.saldo >= 0 ? "#4FE8BE" : "#FF8A76", flexShrink: 0 }}>
                      {hidden ? "•••" : `${f.saldo >= 0 ? "+" : "−"}${fmtBRL(Math.abs(f.saldo))}`}
                    </div>
                  </div>
                  {f.vaiNegativar && (
                    <div style={{
                      background: "#2A151C", border: "1px solid #4A2430", borderRadius: 12, padding: "9px 11px",
                      display: "flex", alignItems: "flex-start", gap: 7, fontSize: 11.5, color: "#FFB4BE", marginBottom: 4
                    }}>
                      <AlertCircle size={13} style={{ flexShrink: 0, marginTop: 1 }} />
                      <span>
                        {f.fonte} fica negativa com {f.itensPendentes.length > 1 ? `os ${f.itensPendentes.length} custos pendentes` : "o custo pendente"} desta fonte
                        {" "}({f.itensPendentes.slice(0, 2).map(it => it.descricao).join(", ")}{f.itensPendentes.length > 2 ? "…" : ""}) — considere transferir de outra fonte antes do vencimento.
                      </span>
                    </div>
                  )}
                </div>
              ))}
              {pix && (pix.n > 0 || pix.nPend > 0) && (
                <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 0 2px", borderTop: "1px dashed rgba(255,255,255,0.12)", marginTop: 4 }}>
                  <span style={{ width: 9, height: 9, borderRadius: 3, border: "1.5px solid #7C86AD", flexShrink: 0 }} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 13.5, fontWeight: 700, color: "#B8BFDB" }}>PIX <span style={{ fontWeight: 500, fontSize: 11.5, color: "#7C86AD" }}>· meio de pagamento</span></div>
                    <div style={{ fontSize: 11.5, color: "#7C86AD", marginTop: 1 }}>
                      {pix.n} {pix.n === 1 ? "lançamento pago" : "lançamentos pagos"} · já somado na Conta Corrente
                      {pix.nPend > 0 && !hidden ? ` · ${pix.nPend} a pagar ${fmtBRL(pix.totalPend)}` : ""}
                    </div>
                  </div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 13.5, color: "#B8BFDB", flexShrink: 0 }}>
                    {hidden ? "•••" : fmtBRL(pix.total)}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// Overdue or due-within-7-days custos/demandas, surfaced at the very top
// of the dashboard so nothing urgent needs to be hunted for.
function useAtencao(custos, demandas, dias = 7, diaHoje = null) {
  return useMemo(() => {
    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    const limite = new Date(hoje); limite.setDate(limite.getDate() + dias);
    const toDateObj = (s) => {
      if (!s) return null;
      const [y, m, d] = s.split("-").map(Number);
      if (!y || !m || !d) return null;
      return new Date(y, m - 1, d);
    };
    const itensCusto = custos
      .filter(c => c.status !== "Pago" && c.vencimento)
      .map(c => ({ id: `c-${c.id}`, rawId: c.id, tab: "custos", label: c.descricao, valor: Number(c.valorPrevisto || 0), data: c.vencimento, dateObj: toDateObj(c.vencimento) }))
      .filter(i => i.dateObj && i.dateObj <= limite);
    const itensDemanda = demandas
      .filter(d => d.status !== "Concluída" && d.prazo)
      .map(d => ({ id: `d-${d.id}`, rawId: d.id, tab: "demandas", label: d.nome, data: d.prazo, dateObj: toDateObj(d.prazo), prio: d.prioridade, hora: ((d.detalhes || "").match(/\b(\d{1,2}):(\d{2})\b/) || [""])[0] }))
      .filter(i => i.dateObj && i.dateObj <= limite);
    return [...itensCusto, ...itensDemanda]
      .map(i => ({ ...i, atrasado: i.dateObj < hoje, diasAtraso: Math.max(0, Math.round((hoje - i.dateObj) / 86400000)), diasFalta: Math.max(0, Math.round((i.dateObj - hoje) / 86400000)) }))
      .sort((a, b) => {
        // Mesmo dia: prioridade (Alta primeiro) e depois nome — igual à ordem da aba Tarefas.
        const ord = { "Alta": 0, "Média": 1, "Baixa": 2 };
        return (a.dateObj - b.dateObj) || ((ord[a.prio] ?? 1) - (ord[b.prio] ?? 1)) || String(a.label || "").localeCompare(String(b.label || ""));
      });
  }, [custos, demandas, dias, diaHoje]);
}

// "Atrasado há 20 dias" stays in days; very old items read better in months.
// Center label for donut charts (total inside the hole).
function DonutTotal({ total, hidden }) {
  const txt = hidden ? "•••" : Number(total || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return (
    <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
      <span style={{ fontSize: 9, color: C.faint, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5 }}>total</span>
      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fontWeight: 700, color: C.ink }}>{txt}</span>
    </div>
  );
}
// Phone numbers inside free text become tappable (tel:) links.
const renderTel = (texto) => {
  const partes = String(texto || "").split(/(\+?\(?\d[\d\s\-()]{6,}\d)/g);
  return partes.map((p, i) => i % 2 === 1
    ? <a key={i} href={`tel:${p.replace(/[^\d+]/g, "")}`} onClick={(e) => e.stopPropagation()} style={{ color: C.lavender, textDecoration: "underline", whiteSpace: "nowrap" }}>{p}</a>
    : p);
};
// Saldos within this distance of zero are shown in a neutral tone (a month that closed even).
const SALDO_NEUTRO = 100;
const pctDe = (v, lista) => { const t = lista.reduce((s, x) => s + x.value, 0); return t > 0 ? `${Math.round((v / t) * 100)}%` : ""; };
const saldoCor = (v) => Math.abs(v) < SALDO_NEUTRO ? C.lavender : v < 0 ? C.coral : C.emerald;
const fmtAtraso = (dias) => {
  if (!(dias > 0)) return "Atrasado";
  if (dias < 30) return `Atrasado há ${dias} ${dias === 1 ? "dia" : "dias"}`;
  const meses = Math.floor(dias / 30);
  return `Atrasado há ${meses} ${meses === 1 ? "mês" : "meses"}`;
};

// Days from today to a YYYY-MM-DD date (negative = already past); null when there is no valid date.
const diasAteVenc = (iso) => {
  if (!iso) return null;
  const [y, m, d] = String(iso).split("-").map(Number);
  if (!y || !m || !d) return null;
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  return Math.round((new Date(y, m - 1, d) - hoje) / 86400000);
};

// Agenda: "qua 19/11 · 10:00 · em 20 dias" (hoje/amanhã por extenso).
// Central de avisos: atrasados + o que vence nos próximos 14 dias (custos, tarefas e fixos ainda não lançados). Só leitura.
// Tarefa atrasada há 30 dias ou mais sai da contagem de urgentes e vai para o grupo "Paradas" (continua visível e acionável).
const PARADA_DIAS = 30;
function useCentralAvisos(custos, demandas, fixos, dias = 14, diaHoje = null) {
  const base = useAtencao(custos, demandas, dias, diaHoje);
  return useMemo(() => {
    const mes = currentMonth();
    const lista = base.map(i => ({
      id: i.id, tipo: i.tab === "demandas" ? "Tarefa" : "Custo", label: i.label, data: i.data, tab: i.tab, rawId: i.rawId,
      valor: i.tab === "custos" ? i.valor : null, atrasado: i.atrasado,
      sub: i.atrasado ? fmtAtraso(i.diasAtraso) : rotuloAgenda(i)
    }));
    (fixos && fixos.length ? fixosPendentes(fixos, custos, mes).faltam : []).forEach(f => {
      const data = vencimentoDoFixo(f, mes);
      const d = diasAteVenc(data);
      if (d === null || d > dias) return;
      const v = Number(f.valorPago || f.valor || 0);
      lista.push({
        id: `f-${f.id || f.descricao}`, tipo: "Fixo", label: f.descricao, data, fixo: true, valor: v > 0 ? v : null, atrasado: d < 0,
        sub: d < 0 ? `a lançar · venceu dia ${data.slice(8, 10)}` : d === 0 ? "a lançar · vence hoje" : d === 1 ? "a lançar · vence amanhã" : `a lançar · vence em ${d} dias`
      });
    });
    lista.sort((a, b) => String(a.data).localeCompare(String(b.data)) || String(a.label).localeCompare(String(b.label)));
    const ehParada = (i) => i.tipo === "Tarefa" && (diasAteVenc(i.data) ?? 0) <= -PARADA_DIAS;
    const itens = lista.filter(i => !ehParada(i));
    const parados = lista.filter(ehParada);
    return { itens, parados, nUrgentes: itens.filter(i => diasAteVenc(i.data) <= 0).length };
  }, [base, custos, fixos, dias, diaHoje]);
}
const rotuloAgenda = (item) => {
  const [y, m, d] = String(item.data).split("-").map(Number);
  const dow = ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"][new Date(y, m - 1, d).getDay()];
  const quando = item.diasFalta === 0 ? "hoje" : item.diasFalta === 1 ? "amanhã" : `em ${item.diasFalta} dias`;
  return `${dow} ${fmtDate(item.data).slice(0, 5)}${item.hora ? " · " + item.hora : ""} · ${quando}`;
};
function AttentionCard({ itens, onNavigate, onConcluirDemanda, title = "Precisa de atenção", limit = 6, emptyText, okWhenEmpty = false, neutro = false, agenda = false, flush = false, recolhivel = false, onPagar = null, semCabecalho = false }) {
  const [aberto, setAberto] = useState(false);
  // Recolhível: começa fechado (a escolha fica guardada neste aparelho) e abre sozinho se aparecer um atraso novo.
  const [recolhido, setRecolhido] = useState(() => {
    if (!recolhivel) return false;
    try { return localStorage.getItem("pt_atencao_aberta") !== "1"; } catch (e) { return true; }
  });
  const qtdAnterior = useRef(itens.length);
  useEffect(() => {
    if (recolhivel && itens.length > qtdAnterior.current) setRecolhido(false);
    qtdAnterior.current = itens.length;
  }, [itens.length, recolhivel]);
  const alternarRecolhido = () => setRecolhido(r => {
    const n = !r;
    try { localStorage.setItem("pt_atencao_aberta", n ? "0" : "1"); } catch (e) { /* sem armazenamento: só não lembra */ }
    return n;
  });
  const fechado = recolhivel && recolhido;
  if (itens.length === 0) {
    // Cards with an emptyText stay on screen with a calm message; the others simply disappear.
    if (!emptyText) return null;
    return (
      <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 18, padding: "16px 18px", marginBottom: flush ? 0 : 20 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
          {okWhenEmpty ? <Check size={16} color={C.emerald} strokeWidth={2.6} /> : <Clock size={16} color={C.faint} />}
          <span style={{ fontWeight: 700, fontSize: 13, color: C.ink, textTransform: "uppercase", letterSpacing: 0.4 }}>{title}</span>
        </div>
        <div style={{ fontSize: 12.5, color: okWhenEmpty ? C.emerald : C.muted, fontWeight: okWhenEmpty ? 600 : 400 }}>{emptyText}</div>
      </div>
    );
  }
  const visiveis = aberto ? itens : itens.slice(0, limit);
  const escondidos = itens.length - limit;
  return (
    <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 18, padding: "16px 18px", marginBottom: flush ? 0 : 20 }}>
      <div onClick={recolhivel ? alternarRecolhido : undefined} role={recolhivel ? "button" : undefined} aria-expanded={recolhivel ? !recolhido : undefined}
        style={{ display: semCabecalho ? "none" : "flex", alignItems: "center", gap: 8, marginBottom: fechado ? 0 : 12, cursor: recolhivel ? "pointer" : "default" }}>
        {neutro ? <Calendar size={16} color={C.lavender} /> : <AlertCircle size={16} color={C.coral} />}
        <span style={{ fontWeight: 700, fontSize: 13, color: C.ink, textTransform: "uppercase", letterSpacing: 0.4 }}>{title}</span>
        <span style={{ fontSize: 11, fontWeight: 700, color: neutro ? C.muted : C.coral, background: neutro ? C.border : C.coralSoft, borderRadius: 10, padding: "1px 8px" }}>
          {itens.length}
        </span>
        {fechado && itens[0] && (
          <span style={{ flex: 1, minWidth: 0, fontSize: 12, color: C.muted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
            {itens[0].label} · {itens[0].atrasado ? fmtAtraso(itens[0].diasAtraso) : fmtDate(itens[0].data)}
          </span>
        )}
        {recolhivel && (
          <span style={{ marginLeft: "auto", flexShrink: 0, width: 30, height: 30, borderRadius: 10, border: `1px solid ${C.border}`, color: C.inkSoft, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Plus size={16} style={{ transform: fechado ? "none" : "rotate(45deg)", transition: "transform .2s ease" }} />
          </span>
        )}
      </div>
      {!fechado && <div style={{ display: "flex", flexDirection: "column" }}>
        {visiveis.map((item, i) => (
          <div
            key={item.id}
            className="tap"
            onClick={() => onNavigate && onNavigate(item.tab, item.rawId)}
            style={{
              display: "flex", alignItems: "center", gap: 10, cursor: onNavigate ? "pointer" : "default",
              padding: "9px 2px", borderBottom: i < visiveis.length - 1 ? `1px solid ${C.border}` : "none"
            }}
          >
            <Clock size={14} color={item.atrasado ? C.coral : C.amber} style={{ flexShrink: 0 }} />
            {agenda && !item.atrasado ? (
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, color: C.ink, fontWeight: 600, lineHeight: 1.3 }}>{item.label}</div>
                <div style={{ fontSize: 11.5, fontWeight: 700, color: C.amber, marginTop: 2 }}>{rotuloAgenda(item)}</div>
              </div>
            ) : (<>
            <span style={{ flex: 1, fontSize: 13, color: C.ink, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {item.label}
            </span>
            <span style={{ fontSize: 11.5, fontWeight: 700, color: item.atrasado ? C.coral : C.amber, flexShrink: 0 }}>
              {item.atrasado ? fmtAtraso(item.diasAtraso) : fmtDate(item.data)}
            </span>
            {item.tab === "custos" && item.valor > 0 && <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, color: C.muted, flexShrink: 0, minWidth: 78, textAlign: "right" }}>{fmtBRL(item.valor)}</span>}</>)}
            {item.tab === "custos" && onPagar && (
              <button type="button" title="Pagar" onClick={(e) => { e.stopPropagation(); onPagar(item.rawId); }}
                style={{ flexShrink: 0, height: 26, borderRadius: 13, border: `1px solid ${C.emerald}`, background: C.emeraldSoft, color: C.emerald, cursor: "pointer", fontSize: 11.5, fontWeight: 700, padding: "0 10px" }}>
                Pagar
              </button>
            )}
            {item.tab === "demandas" && onConcluirDemanda && (
              <button
                type="button"
                title="Marcar como concluída"
                aria-label="Marcar demanda como concluída"
                onClick={async (e) => {
                  e.stopPropagation();
                  if (await confirmar(`Marcar "${item.label}" como Concluída?`, { okLabel: "Concluir" })) onConcluirDemanda(item.rawId);
                }}
                style={{
                  flexShrink: 0, width: 26, height: 26, borderRadius: 13, border: `1px solid ${C.border}`,
                  background: "transparent", color: C.emerald, cursor: "pointer",
                  display: "flex", alignItems: "center", justifyContent: "center", padding: 0
                }}
              >
                <Check size={14} />
              </button>
            )}
          </div>
        ))}
      </div>}
      {!fechado && escondidos > 0 && (
        <button type="button" onClick={() => setAberto(a => !a)} style={{
          marginTop: 8, border: "none", background: "transparent", color: C.muted, cursor: "pointer",
          fontSize: 12, fontWeight: 700, padding: "4px 2px"
        }}>
          {aberto ? "Ver menos" : `+${escondidos} mais`}
        </button>
      )}
    </div>
  );
}

function ProjBar({ label, done, pending, color, max, hidden }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
      <span style={{ width: 62, fontSize: 11.5, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: 0.3 }}>{label}</span>
      <div style={{ flex: 1, height: 14, borderRadius: 7, background: C.border, overflow: "hidden", display: "flex" }}>
        <div style={{ width: `${(done / max) * 100}%`, background: color }} />
        <div style={{ width: `${(pending / max) * 100}%`, background: `repeating-linear-gradient(135deg, ${color}88 0 6px, ${color}33 6px 12px)` }} />
      </div>
      <span style={{ minWidth: 130, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: C.ink }}>
        {hidden ? "•••" : fmtBRL(done + pending)}
      </span>
    </div>
  );
}

// "Ritmo do mês": compares what's been paid so far this month against what was
// typically already paid by this same day of the month, averaged over the last
// few months of REAL per-day payment history (dataPagamento), not a uniform/
// linear daily distribution — months front-load several fixed costs (Pólo,
// Pensão, Unimed...), so a naive average would trigger false alarms early on.
// Degrades gracefully to "indisponível" when there's no paid history yet.
function useRitmoMes(custos, mesAtual, nMeses = 3) {
  return useMemo(() => {
    const [anoAtual, mesNumAtual] = mesAtual.split("-").map(Number);
    const hoje = new Date();
    const ehMesCorrente = hoje.getFullYear() === anoAtual && (hoje.getMonth() + 1) === mesNumAtual;
    const diaRef = ehMesCorrente ? hoje.getDate() : new Date(anoAtual, mesNumAtual, 0).getDate();

    const historico = [];
    for (let i = 1; i <= nMeses; i++) {
      const d = new Date(anoAtual, mesNumAtual - 1 - i, 1);
      const mk = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const pagosDoMes = custos.filter(c => c.mes === mk && c.status === "Pago" && c.dataPagamento);
      if (pagosDoMes.length === 0) continue;
      const ateODia = pagosDoMes
        .filter(c => Number(c.dataPagamento.split("-")[2]) <= diaRef)
        .reduce((s, c) => s + Number(c.valorPago || 0), 0);
      const totalMes = pagosDoMes.reduce((s, c) => s + Number(c.valorPago || 0), 0);
      historico.push({ mk, ateODia, totalMes });
    }

    if (historico.length === 0) return { disponivel: false, diaRef };

    const mediaAteODia = historico.reduce((s, m) => s + m.ateODia, 0) / historico.length;
    // If nothing was ever paid this early in the month across the historical
    // months, there's no usable reference point yet — showing a fabricated
    // "999%" would be as misleading as the metric this replaced.
    if (mediaAteODia <= 0) return { disponivel: false, diaRef };

    const mediaMensal = historico.reduce((s, m) => s + m.totalMes, 0) / historico.length;
    const pagoAteAgora = custos
      .filter(c => c.mes === mesAtual && c.status === "Pago" && c.dataPagamento && Number(c.dataPagamento.split("-")[2]) <= diaRef)
      .reduce((s, c) => s + Number(c.valorPago || 0), 0);
    const pct = (pagoAteAgora / mediaAteODia) * 100;

    return { disponivel: true, diaRef, mediaAteODia, mediaMensal, pagoAteAgora, pct, nMeses: historico.length };
  }, [custos, mesAtual, nMeses]);
}

// Cumulative "pago" day by day for one month, plus what is still to pay placed on its
// due date (overdue items land on today), against the average cumulative of the
// previous 3 months that had payments (same method as the Ritmo do mês).
function serieSaidas(custos, mesKey, fonte, minMes) {
  const [ano, mes] = mesKey.split("-").map(Number);
  const nDias = new Date(ano, mes, 0).getDate();
  const hoje = new Date();
  const corrente = hoje.getFullYear() === ano && hoje.getMonth() + 1 === mes;
  const futuro = mesKey > currentMonth();
  const diaRef = corrente ? hoje.getDate() : futuro ? 0 : nDias;
  const doFiltro = (c) => fonteCombina(c, fonte);
  const porDia = (mk) => {
    const arr = new Array(32).fill(0);
    let n = 0;
    custos.forEach(c => {
      if (c.mes !== mk || c.status !== "Pago" || !c.dataPagamento || !doFiltro(c)) return;
      const d = Number(c.dataPagamento.split("-")[2]);
      if (d >= 1 && d <= 31) { arr[d] += Number(c.valorPago || 0); n++; }
    });
    return { arr, n };
  };
  const atual = porDia(mesKey);
  const hist = [];
  for (let i = 1; i <= 3; i++) {
    const mk = shiftMonth(mesKey, -i);
    if (minMes && mk < minMes) continue; // months before "contar a partir de" have incomplete launches
    const h = porDia(mk);
    if (h.n > 0) hist.push(h.arr);
  }
  // Still to pay: placed on the due date (clamped to the month); overdue ones on today.
  const abertos = custos.filter(c => c.mes === mesKey && c.status !== "Pago" && doFiltro(c));
  const faltam = abertos.reduce((t, c) => t + Number(c.valorPrevisto || 0), 0);
  const pendPorDia = new Array(nDias + 2).fill(0);
  let nComData = 0;
  abertos.forEach(c => {
    if (!c.vencimento) return;
    const [vy, vm, vd] = c.vencimento.split("-").map(Number);
    if (!vy || !vm || !vd) return;
    let dia = (vy === ano && vm === mes) ? vd : (vy * 12 + vm < ano * 12 + mes ? 1 : nDias);
    dia = Math.min(nDias, Math.max(1, dia));
    if (corrente) dia = Math.max(dia, diaRef);
    pendPorDia[dia] += Number(c.valorPrevisto || 0);
    nComData++;
  });
  const temProj = (corrente || futuro) && nComData > 0;
  const rows = [];
  let acA = 0;
  const acH = hist.map(() => 0);
  for (let d = 1; d <= nDias; d++) {
    acA += atual.arr[d];
    hist.forEach((arr, i) => { acH[i] += arr[d]; });
    rows.push({
      dia: d,
      atual: d <= diaRef ? acA : null,
      normal: hist.length ? acH.reduce((t, v) => t + v, 0) / hist.length : null,
      prev: null
    });
  }
  if (temProj) {
    // dashed projection: real cumulative at the reference day + pending items due up to each day
    let base = 0;
    for (let d = 1; d <= diaRef; d++) base += atual.arr[d];
    let acc = 0;
    for (let d = 1; d <= nDias; d++) {
      acc += pendPorDia[d];
      if (d >= Math.max(1, diaRef)) rows[d - 1].prev = base + acc;
    }
  }
  const ref = rows[Math.max(0, diaRef - 1)];
  return {
    rows, nDias, diaRef, corrente, futuro, nBase: hist.length, temProj, faltam, nFaltam: abertos.length,
    pagoAteHoje: diaRef > 0 && ref ? ref.atual || 0 : 0, normalAteHoje: ref && ref.normal !== null && diaRef > 0 ? ref.normal : 0,
    temDados: atual.n > 0 || ((corrente || futuro) && (hist.length > 0 || abertos.length > 0))
  };
}

function SaidasMesCard({ serie, mesLabel, entrada = false }) {
  const { rows, nDias, diaRef, corrente, nBase, temProj, faltam, nFaltam, pagoAteHoje, normalAteHoje } = serie;
  const ticks = [1, 5, 10, 15, 20, 25, nDias].filter((v, i, a) => v <= nDias && a.indexOf(v) === i);
  const cor = entrada ? "#00E0A8" : "#FF5A68";
  const T = entrada
    ? { titulo: "Entradas ao longo do mês", feito: "Recebido", falta: "A receber", feitoMin: "recebido", sem: "sem receber", un: ["receita", "receitas"], gid: "gEntradasMes", serie: "Com o que falta receber" }
    : { titulo: "Saídas ao longo do mês", feito: "Pago", falta: "A pagar", feitoMin: "pago", sem: "sem pagar", un: ["conta", "contas"], gid: "gSaidasMes", serie: "Com o que falta pagar" };
  const dado = (t) => <strong style={{ color: C.ink }}>{t}</strong>;
  return (
    <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 16, display: "flex", flexDirection: "column" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 8, flexWrap: "wrap", marginBottom: 4 }}>
        <span style={{ fontWeight: 700, fontSize: 12.5, color: C.ink, textTransform: "uppercase", letterSpacing: 0.4 }}>{T.titulo}</span>
        <span style={{ display: "flex", gap: 10, fontSize: 11, color: C.muted, fontWeight: 600, flexWrap: "wrap" }}>
          <span style={{ display: "flex", alignItems: "center", gap: 4 }}><span style={{ width: 10, height: 3, borderRadius: 2, background: cor }} />{T.feito}</span>
          {temProj && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><span style={{ width: 10, height: 0, borderTop: `2px dotted ${cor}` }} />{T.falta}</span>}
          {nBase > 0 && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><span style={{ width: 10, height: 0, borderTop: `2px dashed ${C.muted}` }} />Normal</span>}
        </span>
      </div>
      <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 8, lineHeight: 1.5 }}>
        <div>
          {corrente ? <>Até o dia {diaRef}: {T.feitoMin} {dado(fmtBRL(pagoAteHoje))}</> : serie.futuro ? <>Mês ainda não começou</> : <>Mês fechado: {T.feitoMin} {dado(fmtBRL(pagoAteHoje))}</>}
          {nFaltam > 0 && <> · {corrente || serie.futuro ? "faltam" : T.sem} {dado(fmtBRL(faltam))} em {nFaltam} {nFaltam === 1 ? T.un[0] : T.un[1]}</>}
        </div>
        {nBase > 0 && !serie.futuro && <div>Normal até o dia {diaRef}: {dado(fmtBRL(normalAteHoje))}</div>}
        {nBase === 0 && !temProj && <div>Ainda sem meses anteriores com {entrada ? "recebimentos" : "pagamentos"} para comparar.</div>}
      </div>
      <div style={{ width: "100%", flex: 1, minHeight: 150 }}>
        <ResponsiveContainer width="100%" height="100%" minHeight={150}>
          <AreaChart data={rows} margin={{ top: 6, right: 6, left: 6, bottom: 0 }}>
            <defs>
              <linearGradient id={T.gid} x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor={cor} stopOpacity={0.35} /><stop offset="100%" stopColor={cor} stopOpacity={0.02} />
              </linearGradient>
            </defs>
            <XAxis dataKey="dia" ticks={ticks} tick={{ fontSize: 10.5, fill: C.faint }} axisLine={{ stroke: C.border }} tickLine={false} />
            <YAxis hide domain={[0, "dataMax"]} />
            {corrente && <ReferenceLine x={diaRef} stroke={C.border} strokeDasharray="3 3" />}
            <Tooltip
              labelFormatter={(l) => `Dia ${l}`}
              formatter={(v, n) => [fmtBRL(v), n]}
              contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12, color: C.ink }}
              itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
            {nBase > 0 && <Area type="monotone" dataKey="normal" name="Normal" stroke={C.muted} strokeWidth={1.6} strokeDasharray="5 4" fill="none" dot={false} isAnimationActive={false} />}
            {temProj && <Area type="stepAfter" dataKey="prev" name={T.serie} stroke={cor} strokeOpacity={0.75} strokeWidth={2} strokeDasharray="2 4" fill="none" dot={false} isAnimationActive={false} />}
            <Area type="monotone" dataKey="atual" name={T.feito} stroke={cor} strokeWidth={2.4} fill={`url(#${T.gid})`} dot={false} connectNulls={false} />
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

// Receitas seen as custo-shaped rows, so the same month series and "next due" led work for both.
const receitasComoCustos = (rs) => rs.map(r => {
  const rec = r.status === "Recebido";
  return { ...r, orig: r, status: rec ? "Pago" : "Pendente", dataPagamento: rec ? (r.dataRecebimento || "") : "",
    valorPago: r.valorRecebido, vencimento: rec ? "" : (r.dataRecebimento || ""), fonte: r.contaDestino };
});

// "Led" with the next cost to be paid: red if overdue, amber within 7 days, green beyond.
function proximoVencimento(filtrados) {
  const abertos = filtrados.filter(c => c.status !== "Pago");
  if (abertos.length === 0) return null;
  const comData = abertos.filter(c => c.vencimento).sort((a, b) => a.vencimento.localeCompare(b.vencimento));
  const item = comData[0] || abertos[0];
  const dias = item.vencimento ? diasAteVenc(item.vencimento) : null;
  const faltaPagar = abertos.reduce((t, c) => t + Number(c.valorPrevisto || 0), 0);
  let texto, cor;
  if (dias === null) { texto = "Sem data de vencimento"; cor = C.amber; }
  else if (dias < 0) { texto = fmtAtraso(-dias); cor = C.coral; }
  else if (dias === 0) { texto = "Vence hoje"; cor = C.amber; }
  else { texto = `Vence em ${dias} ${dias === 1 ? "dia" : "dias"}`; cor = dias <= 7 ? C.amber : C.emerald; }
  return { item, texto, cor, nAbertos: abertos.length, faltaPagar };
}

/* ------------------------------------------------------------------ */
/*  Fixos do mês — recurring costs list. It never writes by itself:    */
/*  "Lançar" only opens the normal new-custo form, pre-filled.         */
/* ------------------------------------------------------------------ */
const FIXOS_INICIAIS = [
  { descricao: "Aluguel", dia: 10, valor: 2675, fonte: "Conta Corrente" },
  { descricao: "Prestação Pólo", dia: 3, valor: 1773.86, fonte: "Conta Salário" },
  { descricao: "Prestação Gibson", dia: 30, valor: null, fonte: "Dinheiro" },
  { descricao: "Facul Roberto", dia: 10, valor: 963.16, valorPago: 679.33, fonte: "" },
  { descricao: "Facul Rubem", dia: 10, valor: 1282.85, valorPago: 862.90, fonte: "" },
  { descricao: "Condomínio Via", dia: 10, valor: 813.47, fonte: "" },
  { descricao: "Condomínio Altos", dia: 10, valor: 688.57, fonte: "" },
  { descricao: "Unimed", dia: 10, valor: null, fonte: "" },
  { descricao: "Pensão Rubem", dia: 5, valor: 1000, fonte: "" },
  { descricao: "Pensão Zannie", dia: 5, valor: 700, fonte: "" },
];
const normDesc = (t) => String(t || "").trim().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
// Vencimento of a fixo in month "AAAA-MM", clamped to the month's last day.
function vencimentoDoFixo(f, mes) {
  const [a, mm] = mes.split("-").map(Number);
  const ultimo = new Date(a, mm, 0).getDate();
  const d = Math.min(Math.max(parseInt(f.dia, 10) || 1, 1), ultimo);
  return `${mes}-${String(d).padStart(2, "0")}`;
}
// Ordem dos fixos: dia do vencimento; no mesmo dia, maior valor primeiro (valor varia por último); depois nome.
const cmpFixos = (a, b) => (a.dia || 0) - (b.dia || 0) || (Number(b.valor || 0) - Number(a.valor || 0)) || String(a.descricao || "").localeCompare(String(b.descricao || ""));
function fixosPendentes(fixos, custos, mes) {
  const lancados = new Set(custos.filter(c => c.mes === mes).map(c => normDesc(c.descricao)));
  const ativos = fixos.filter(f => f.ativo !== false);
  return { ativos, faltam: ativos.filter(f => !lancados.has(normDesc(f.descricao))).sort(cmpFixos) };
}

function FixosAVencerCard({ itens, onOpen }) {
  if (!itens || itens.length === 0) return null;
  const nomes = itens.slice(0, 3).map(f => `${f.descricao} (dia ${f.dia})`).join(" · ");
  const resto = itens.length - 3;
  const um = itens.length === 1;
  return (
    <div className="tap in" onClick={onOpen} style={{ display: "flex", alignItems: "center", gap: 12, background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: "13px 16px", marginBottom: 16, cursor: "pointer" }}>
      <span style={{ width: 11, height: 11, borderRadius: 6, background: C.amber, boxShadow: `0 0 10px ${C.amber}`, flexShrink: 0 }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 10.5, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: 0.5 }}>Fixos a lançar</div>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.ink }}>{itens.length} {um ? "fixo vence" : "fixos vencem"} em até 7 dias e ainda {um ? "não foi lançado" : "não foram lançados"}</div>
        <div style={{ fontSize: 11.5, color: C.faint, marginTop: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{nomes}{resto > 0 ? ` · +${resto}` : ""}</div>
      </div>
      <ChevronDown size={16} color={C.faint} style={{ transform: "rotate(-90deg)", flexShrink: 0 }} />
    </div>
  );
}

function FixosLed({ fixos, custos, mes, onOpen, erro }) {
  const { ativos, faltam } = useMemo(() => fixosPendentes(fixos, custos, mes), [fixos, custos, mes]);
  const base = { display: "flex", alignItems: "center", gap: 12, borderRadius: 16, padding: "13px 16px", marginBottom: 16, cursor: "pointer" };
  if (erro) {
    return (
      <div className="in" style={{ marginBottom: 16, padding: "11px 14px", borderRadius: 14, background: C.amberSoft, color: C.amber, fontSize: 12.5 }}>
        Não consegui ler a lista de fixos (provavelmente falta liberar a coleção "recorrentes" nas regras do Firestore). O resto do app segue normal.
      </div>
    );
  }
  if (ativos.length === 0) {
    return (
      <div className="tap in" onClick={onOpen} style={{ ...base, border: `1px dashed ${C.border}`, background: C.card }}>
        <span style={{ flex: 1, fontSize: 13, color: C.muted }}>Fixos do mês: a lista inicial já está pronta ({FIXOS_INICIAIS.length} fixos). Toque para criar.</span>
        <span style={{ fontSize: 12, fontWeight: 700, color: C.lavender }}>Abrir</span>
      </div>
    );
  }
  const tudo = faltam.length === 0;
  const cor = tudo ? C.emerald : C.amber;
  return (
    <div className="tap in" onClick={onOpen} style={{ ...base, background: tudo ? C.emeraldSoft : C.card, boxShadow: tudo ? "none" : shadowSm }}>
      <span style={{ width: 11, height: 11, borderRadius: 6, background: cor, boxShadow: `0 0 10px ${cor}`, flexShrink: 0 }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 10.5, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: 0.5 }}>Fixos do mês</div>
        <div style={{ fontSize: 14, fontWeight: 700, color: tudo ? C.emerald : C.ink }}>
          {tudo ? `Os ${ativos.length} fixos já estão lançados` : `Faltam lançar ${faltam.length} de ${ativos.length}`}
        </div>
      </div>
      <ChevronDown size={16} color={C.faint} style={{ transform: "rotate(-90deg)", flexShrink: 0 }} />
    </div>
  );
}

function FixosModal({ fixos, custos, mes, onLancar, onAdd, onUpdate, onRemove, onClose }) {
  const { ativos, faltam } = useMemo(() => fixosPendentes(fixos, custos, mes), [fixos, custos, mes]);
  const [aba, setAba] = useState("mes");
  const vazio = { id: null, descricao: "", dia: "", valor: "", valorPago: "", fonte: "Varia", ativo: true };
  const [f, setF] = useState(vazio);
  const [busy, setBusy] = useState(false);
  const lista = useMemo(() => [...fixos].sort(cmpFixos), [fixos]);
  const salvar = async (e) => {
    e.preventDefault();
    if (busy) return;
    const dia = parseInt(f.dia, 10);
    if (!f.descricao.trim() || !(dia >= 1 && dia <= 31)) { alert("Informe a descrição e o dia do mês (1 a 31)."); return; }
    let valor = null;
    if (String(f.valor).trim() !== "") {
      valor = parseValorBR(f.valor);
      if (!isValorValido(valor)) { alert("O valor precisa ser maior que zero (ou deixe em branco se varia)."); return; }
    }
    let valorPago = null;
    if (String(f.valorPago).trim() !== "") {
      valorPago = parseValorBR(f.valorPago);
      if (!isValorValido(valorPago)) { alert("O valor pago esperado precisa ser maior que zero (ou deixe em branco)."); return; }
    }
    const payload = { descricao: f.descricao.trim(), dia, valor, valorPago, fonte: f.fonte === "Varia" ? "" : f.fonte, ativo: f.ativo !== false };
    setBusy(true);
    try { if (f.id) await onUpdate(f.id, payload); else await onAdd(payload); setF(vazio); } finally { setBusy(false); }
  };
  const iniciais = async () => {
    if (busy) return;
    if (!(await confirmar(`Criar a lista inicial com ${FIXOS_INICIAIS.length} fixos? Você pode editar tudo depois.`, { okLabel: "Criar" }))) return;
    setBusy(true);
    try { for (const x of FIXOS_INICIAIS) await onAdd({ ...x, ativo: true }); } finally { setBusy(false); }
  };
  const mini = { border: `1px solid ${C.border}`, background: C.page, color: C.muted, borderRadius: 9, padding: "6px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" };
  return (
    <Modal title="Fixos do mês" onClose={onClose}>
      <div style={{ padding: "4px 0 24px" }}>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          {[["mes", `Para lançar (${faltam.length})`], ["lista", "Gerenciar lista"]].map(([k, t]) => (
            <button key={k} type="button" onClick={() => setAba(k)} style={{ border: `1px solid ${aba === k ? C.emerald : C.border}`, background: aba === k ? C.emeraldSoft : C.page, color: aba === k ? C.emerald : C.muted, borderRadius: 10, padding: "8px 14px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>{t}</button>
          ))}
        </div>
        {aba === "mes" && (
          <div>
            {ativos.length === 0 ? (
              <div style={{ border: `1px dashed ${C.border}`, borderRadius: 14, padding: 18, color: C.muted, fontSize: 13.5 }}>
                <div style={{ marginBottom: 8 }}>A lista inicial já está pronta, é só criar com um toque:</div>
                <div style={{ fontSize: 12.5, color: C.faint, lineHeight: 1.6, marginBottom: 12 }}>{FIXOS_INICIAIS.map(x => x.descricao).join(" · ")}</div>
                <button type="button" onClick={iniciais} disabled={busy} style={{ ...addBtnStyle, opacity: busy ? 0.7 : 1 }}>Criar lista inicial</button>
              </div>
            ) : faltam.length === 0 ? (
              <div style={{ display: "flex", alignItems: "center", gap: 8, color: C.emerald, fontSize: 13.5, fontWeight: 600, padding: "12px 0" }}><Check size={16} strokeWidth={2.6} /> Todos os {ativos.length} fixos de {monthLabel(mes)} já estão lançados.</div>
            ) : faltam.map((f, i) => (
              <div key={f.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0", borderTop: i === 0 ? "none" : `1px solid ${C.border}` }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 14, fontWeight: 700, color: C.ink }}>{f.descricao}</div>
                  <div style={{ fontSize: 11.5, color: C.faint }}>dia {f.dia} · {f.fonte || "fonte varia"}</div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 700, color: f.valor ? C.ink : C.faint }}>{f.valor ? fmtBRL(f.valor) : "valor varia"}</div>
                  {f.valorPago ? <div style={{ fontSize: 11, color: C.emerald }}>paga {fmtBRL(f.valorPago)}</div> : null}
                </div>
                <button type="button" onClick={() => onLancar(f, false)} style={{ border: `1px solid ${C.border}`, background: C.page, color: C.inkSoft, borderRadius: 10, padding: "8px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>Lançar</button>
                <button type="button" onClick={() => onLancar(f, true)} style={{ border: `1px solid ${C.emerald}`, background: C.emeraldSoft, color: C.emerald, borderRadius: 10, padding: "8px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>✓ Pago hoje</button>
              </div>
            ))}
          </div>
        )}
        {aba === "lista" && <div>
        {lista.length === 0 && (
          <div style={{ border: `1px dashed ${C.border}`, borderRadius: 14, padding: 18, textAlign: "center", color: C.muted, fontSize: 13.5, marginBottom: 16 }}>
            Nenhum fixo ainda.
            <div style={{ marginTop: 10 }}><button type="button" onClick={iniciais} disabled={busy} style={{ ...addBtnStyle, margin: "0 auto" }}>Criar lista inicial</button></div>
          </div>
        )}
        {lista.map(x => (
          <div key={x.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0", borderBottom: `1px solid ${C.border}`, opacity: x.ativo === false ? 0.5 : 1 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.ink }}>{x.descricao}{x.ativo === false ? " · desativado" : ""}</div>
              <div style={{ fontSize: 11.5, color: C.faint }}>dia {x.dia} · {x.valor ? fmtBRL(x.valor) : "valor varia"}{x.valorPago ? ` (paga ${fmtBRL(x.valorPago)})` : ""} · {x.fonte || "fonte varia"}</div>
            </div>
            <button type="button" style={mini} onClick={() => setF({ id: x.id, descricao: x.descricao, dia: String(x.dia), valor: x.valor ? fmtValorInput(x.valor) : "", valorPago: x.valorPago ? fmtValorInput(x.valorPago) : "", fonte: x.fonte || "Varia", ativo: x.ativo !== false })}>Editar</button>
            <button type="button" style={mini} onClick={() => onUpdate(x.id, { ativo: x.ativo === false })}>{x.ativo === false ? "Ativar" : "Desativar"}</button>
            <button type="button" style={{ ...mini, color: C.coral }} onClick={async () => { if (await confirmar(`Excluir "${x.descricao}" da lista de fixos? Os custos já lançados não mudam.`, { perigo: true, okLabel: "Excluir" })) onRemove(x.id); }}>Excluir</button>
          </div>
        ))}
        <form onSubmit={salvar} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginTop: 20 }}>
          <div style={{ gridColumn: "span 2", fontSize: 12, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: 0.5 }}>{f.id ? "Editar fixo" : "Novo fixo"}</div>
          <Field label="Descrição" span={2}><TextInput value={f.descricao} onChange={e => setF({ ...f, descricao: e.target.value })} /></Field>
          <Field label="Dia do vencimento"><TextInput type="text" inputMode="numeric" placeholder="1 a 31" value={f.dia} onChange={e => setF({ ...f, dia: e.target.value.replace(/\D/g, "").slice(0, 2) })} /></Field>
          <Field label="Valor previsto (vazio = varia)"><TextInput type="text" inputMode="decimal" placeholder="0,00" value={f.valor} onChange={e => setF({ ...f, valor: e.target.value })} /></Field>
          <Field label="Valor pago esperado (ex.: com desconto)"><TextInput type="text" inputMode="decimal" placeholder="0,00" value={f.valorPago} onChange={e => setF({ ...f, valorPago: e.target.value })} /></Field>
          <Field label="Fonte"><SelectInput options={["Varia", ...FONTES_FORM]} value={f.fonte} onChange={e => setF({ ...f, fonte: e.target.value })} /></Field>
          <div style={{ display: "flex", gap: 10, alignItems: "flex-end" }}>
            <button type="submit" disabled={busy} style={{ ...addBtnStyle, opacity: busy ? 0.7 : 1 }}><Check size={16} /> {f.id ? "Salvar" : "Adicionar"}</button>
            {f.id && <button type="button" style={mini} onClick={() => setF(vazio)}>Cancelar</button>}
          </div>
        </form>
        </div>}
      </div>
    </Modal>
  );
}

const rotuloK = (v) => v >= 1000 ? (v / 1000).toFixed(1).replace(".", ",") + "k" : String(Math.round(v));
function ProximoVencLed({ info, onOpen, entrada = false }) {
  const { item, texto, cor, nAbertos, faltaPagar } = info;
  return (
    <div className="tap in" onClick={() => onOpen(item.orig || item)} style={{
      display: "flex", alignItems: "center", gap: 12, background: C.card, boxShadow: shadowSm, borderRadius: 16,
      padding: "13px 16px", marginBottom: 16, cursor: "pointer"
    }}>
      <span style={{ width: 11, height: 11, borderRadius: 6, background: cor, boxShadow: `0 0 10px ${cor}`, flexShrink: 0 }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 10.5, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: 0.5 }}>{entrada ? "Próximo a receber" : "Próximo vencimento"}</div>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          {item.descricao}{item.vencimento ? ` · ${fmtDate(item.vencimento)}` : ""}
        </div>
        <div style={{ fontSize: 11, color: C.faint, marginTop: 1 }}>
          {nAbertos > 1 ? `${nAbertos} ${entrada ? "a receber" : "a pagar"} neste mês` : `Último ${entrada ? "a receber" : "a pagar"} neste mês`}
        </div>
      </div>
      <div style={{ textAlign: "right", flexShrink: 0 }}>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 14, color: C.ink }}>{fmtBRL(item.valorPrevisto)}</div>
        <div style={{ fontSize: 11.5, fontWeight: 700, color: cor }}>{entrada ? texto.replace("Atrasado", "Atrasada") : texto}</div>
      </div>
    </div>
  );
}

function PixMesCard({ pix, mesLabel, hidden }) {
  const { n, total, nPend, totalPend, itens } = pix;
  return (
    <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 8, marginBottom: 4 }}>
        <span style={{ fontWeight: 700, fontSize: 12.5, color: C.ink, textTransform: "uppercase", letterSpacing: 0.4 }}>PIX do mês</span>
        <span style={{ fontSize: 11, color: C.faint }}>{mesLabel}</span>
      </div>
      {n === 0 && nPend === 0 ? (
        <div style={{ fontSize: 13, color: C.muted, padding: "10px 0 4px" }}>Nenhum PIX registrado neste mês.</div>
      ) : (
        <>
          <div style={{ display: "flex", alignItems: "baseline", gap: 10, flexWrap: "wrap", margin: "6px 0 2px" }}>
            <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 22, color: C.ink }}>{hidden ? "•••" : fmtBRL(total)}</span>
            <span style={{ fontSize: 12.5, color: C.muted, fontWeight: 600 }}>{n} {n === 1 ? "PIX pago" : "PIX pagos"}</span>
          </div>
          <div style={{ fontSize: 11, color: C.faint, marginBottom: itens.length ? 10 : 0 }}>
            Só os lançados com fonte PIX; já somados na Conta Corrente.
            {nPend > 0 && !hidden ? ` Mais ${nPend} a pagar (${fmtBRL(totalPend)}).` : ""}
          </div>
          {itens.slice(0, 5).map(c => (
            <div key={c.id} style={{ display: "flex", justifyContent: "space-between", gap: 10, fontSize: 12.5, padding: "5px 0", borderTop: `1px solid ${C.border}` }}>
              <span style={{ color: C.ink, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                {c.descricao}<span style={{ color: C.faint }}> · {fmtDate(c.dataPagamento)}</span>
              </span>
              <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted, flexShrink: 0 }}>{hidden ? "•••" : fmtBRL(c.valorPago)}</span>
            </div>
          ))}
          {itens.length > 5 && <div style={{ fontSize: 11, color: C.faint, paddingTop: 6 }}>+{itens.length - 5} mais</div>}
        </>
      )}
    </div>
  );
}

// "ALUGUEL APTO 1" → "Aluguel Apto 1" (siglas curtas como "BC" ficam como estão).
const tituloBonito = (v) => (typeof v === "string" && v === v.toUpperCase() && /[A-ZÀ-Ú]{3}/.test(v)) ? v.toLowerCase().replace(/(^|\s)\S/g, m => m.toUpperCase()) : v;

function TooltipPT({ active, payload, label }) {
  if (!active || !payload || payload.length === 0) return null;
  const curto = (n) => String(n).replace("Conta Salário", "CS").replace("Conta Corrente", "CC");
  return (
    <div style={{ background: C.card, boxShadow: shadowMd, border: `1px solid ${C.border}`, borderRadius: 12, padding: "9px 12px", fontSize: 12.5, color: C.ink }}>
      <div style={{ fontSize: 11, color: C.muted, fontWeight: 700, marginBottom: 4 }}>{label}</div>
      {payload.filter(p => p.value != null).map(p => (
        <div key={p.dataKey} style={{ display: "flex", alignItems: "center", gap: 7, padding: "1px 0" }}>
          <span style={{ width: 8, height: 8, borderRadius: 2, background: p.stroke || p.color || p.fill }} />
          <span style={{ color: C.muted }}>{curto(p.name || p.dataKey)}</span>
          <span style={{ marginLeft: "auto", paddingLeft: 12, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700 }}>{fmtBRL(p.value)}</span>
        </div>
      ))}
    </div>
  );
}

// Números-chave no topo de cada aba das Análises.
function TilesAnalise({ itens, hidden }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginBottom: 14 }}>
      {itens.map(t => (
        <div key={t.rotulo} style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: "14px 16px", minWidth: 0 }}>
          <div style={{ fontSize: 10.5, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: 0.5 }}>{t.rotulo}</div>
          <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 22, letterSpacing: -0.4, color: t.cor || C.ink, marginTop: 6, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontVariantNumeric: "tabular-nums" }}>{hidden && t.mask ? "•••" : tituloBonito(t.valor)}</div>
          {t.sub && <div style={{ fontSize: 11, color: C.faint, marginTop: 3, lineHeight: 1.4, whiteSpace: "pre-line", overflowWrap: "anywhere" }}>{hidden && t.mask ? "•••" : t.sub}</div>}
        </div>
      ))}
    </div>
  );
}

// Cartão de abertura das abas (Custos, Receitas, Anual, Tarefas): título, número principal, ação e navegação de mês.
function TabHero({ titulo, mesLabel, onLabelClick, destaque, pct = null, pctLabel = null, chips = null, faixa = null, rodapeExtra = null, rodapeTexto = null, rodapeCor = "#B8C2E8", acao, onPrev, onNext, prevOff = false, nextOff = false }) {
  const navBtn = (off) => ({ width: 34, height: 34, borderRadius: 11, border: "1px solid rgba(255,255,255,0.14)", background: "rgba(255,255,255,0.06)", color: off ? "#5C6485" : "#D5DBF5", cursor: off ? "default" : "pointer", fontSize: 18, lineHeight: 1, display: "flex", alignItems: "center", justifyContent: "center" });
  const chipSt = (cor, clicavel) => ({ fontSize: 12, fontWeight: 700, padding: "4px 11px", borderRadius: 999, background: "rgba(255,255,255,0.07)", border: "1px solid rgba(255,255,255,0.12)", color: cor || "#D5DBF5", cursor: clicavel ? "pointer" : "default", fontFamily: "inherit" });
  return (
    <div className="in" style={{
      position: "relative", display: "flex", flexDirection: "column", minHeight: 356, borderRadius: 24, padding: "20px 22px", marginBottom: 20, border: "3px solid transparent",
      background: "linear-gradient(150deg, #0B0F1F 0%, #101632 60%, #0F3D33 100%) padding-box, linear-gradient(135deg, #00E0A8, #9AA5D1) border-box",
      boxShadow: "0 14px 34px rgba(11,15,31,0.28), 0 4px 12px rgba(11,15,31,0.18)"
    }}>
      <div style={{ position: "absolute", inset: 0, overflow: "hidden", borderRadius: "inherit", pointerEvents: "none" }}>
        <div style={{ position: "absolute", top: -70, right: -60, width: 200, height: 200, borderRadius: "50%", background: `radial-gradient(circle, ${C.emeraldGlow} 0%, transparent 70%)` }} />
      </div>
      <div style={{ position: "relative", display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontSize: 12.5, color: "#B8C2E8", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.6 }}>{titulo}</div>
          {destaque && (
            <>
              <div style={{ fontSize: 11.5, color: "#8B93B5", fontWeight: 600, marginTop: 12 }}>{destaque.rotulo}</div>
              <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: "clamp(30px, 8vw, 40px)", letterSpacing: -1, lineHeight: 1.05, color: destaque.cor || "#F3F5FA", fontVariantNumeric: "tabular-nums", marginTop: 2 }}>{destaque.valor}</div>
              {destaque.sub && <div style={{ fontSize: 12, color: "#8B93B5", marginTop: 6 }}>{destaque.sub}</div>}
            </>
          )}
        </div>
        {acao && (
          <button className="tap" onClick={acao.onClick} style={{ flexShrink: 0, height: 38, borderRadius: 12, border: "none", background: "#00E0A8", color: "#04201A", fontWeight: 700, fontSize: 13, padding: "0 14px", display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
            <Plus size={16} strokeWidth={2.6} /> {acao.label}
          </button>
        )}
      </div>
      {chips && chips.length > 0 && (
        <div style={{ position: "relative", display: "flex", flexWrap: "wrap", gap: 8, marginTop: 14 }}>
          {chips.map((c, i) => c.onClick
            ? <button key={i} type="button" className="tap" onClick={c.onClick} style={chipSt(c.cor, true)}>{c.t}</button>
            : <span key={i} style={chipSt(c.cor, false)}>{c.t}</span>)}
        </div>
      )}
      {pct !== null && (
        <div style={{ position: "relative", marginTop: 16 }}>
          <div style={{ height: 6, borderRadius: 3, background: "rgba(255,255,255,0.1)", overflow: "hidden" }}>
            <div style={{ width: `${Math.max(0, Math.min(100, pct))}%`, height: "100%", borderRadius: 3, background: "linear-gradient(90deg, #00E0A8, #22D3EE)" }} />
          </div>
          {pctLabel && <div style={{ fontSize: 11, color: "#8B93B5", fontWeight: 600, marginTop: 6, textAlign: "right" }}>{pctLabel}</div>}
        </div>
      )}
      {faixa && <div style={{ position: "relative", marginTop: 16 }}>{faixa}</div>}
      {(onPrev || rodapeTexto || rodapeExtra) && (
        <div style={{ position: "relative", marginTop: "auto", paddingTop: 16, display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
          {onPrev ? (
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <button type="button" onClick={prevOff ? undefined : onPrev} style={navBtn(prevOff)} aria-label="Anterior">‹</button>
              {onLabelClick
                ? <button type="button" className="tap" onClick={onLabelClick} aria-label="Escolher mês" style={{ border: "none", background: "transparent", cursor: "pointer", fontFamily: "inherit", fontSize: 13.5, fontWeight: 700, color: "#F3F5FA", minWidth: 118, display: "flex", alignItems: "center", justifyContent: "center", gap: 5 }}>{mesLabel}<ChevronDown size={14} color="#8B93B5" /></button>
                : <span style={{ fontSize: 13.5, fontWeight: 700, color: "#F3F5FA", minWidth: 118, textAlign: "center" }}>{mesLabel}</span>}
              <button type="button" onClick={nextOff ? undefined : onNext} style={navBtn(nextOff)} aria-label="Próximo">›</button>
            </div>
          ) : (
            <div style={{ fontSize: 12.5, color: rodapeCor, fontWeight: 600, minWidth: 0, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{rodapeTexto}</div>
          )}
          {rodapeExtra}
        </div>
      )}
    </div>
  );
}

function BarraEmpilhada({ dados, corDe, hidden, nota, totalLabel = "Total" }) {
  const total = dados.reduce((t, d) => t + d.value, 0) || 1;
  return (
    <div>
      <div style={{ display: "flex", height: 16, borderRadius: 8, overflow: "hidden", gap: 2, marginBottom: 14 }}>
        {dados.map(d => (
          <div key={d.name} title={d.name} style={{ width: `${(d.value / total) * 100}%`, minWidth: 3, background: corDe(d.name), position: "relative", overflow: "hidden" }}>
            {d.pend > 0 && d.value > 0 && <div style={{ position: "absolute", right: 0, top: 0, bottom: 0, width: `${Math.min(100, (d.pend / d.value) * 100)}%`, background: "repeating-linear-gradient(135deg, rgba(10,14,28,0.55) 0 4px, transparent 4px 8px)" }} />}
          </div>
        ))}
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 8 }}>
        <span style={{ fontSize: 11, fontWeight: 700, color: C.faint, textTransform: "uppercase", letterSpacing: 0.5 }}>{totalLabel}</span>
        <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 14, fontWeight: 700, color: C.ink }}>{hidden ? "•••" : fmtBRL(total)}</span>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {dados.map(d => (
          <div key={d.name} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>
            <span style={{ width: 9, height: 9, borderRadius: 3, background: corDe(d.name), flexShrink: 0 }} />
            <span style={{ color: C.ink, flex: 1, minWidth: 0, lineHeight: 1.25, overflow: "hidden", textOverflow: "ellipsis" }}>{tituloBonito(d.name)}</span>
            <span style={{ fontSize: 10.5, color: C.faint, minWidth: 28, textAlign: "right" }}>{pctDe(d.value, dados)}</span>
            <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted, fontSize: 12 }}>{hidden ? "•••" : fmtBRL(d.value)}</span>
          </div>
        ))}
      </div>
      {dados.some(d => d.pend > 0) && <div style={{ fontSize: 10.5, color: C.faint, marginTop: 8 }}>Listrado = ainda a pagar</div>}
      {nota && <div style={{ fontSize: 10.5, color: C.faint, marginTop: 8 }}>{nota}</div>}
    </div>
  );
}

// Régua dos próximos 14 dias (hero de Tarefas): um ponto em cada dia que tem tarefa em aberto com prazo.
function ReguaDias({ dias }) {
  return (
    <div style={{ display: "flex", gap: 4 }} aria-label="Próximos 14 dias com prazos de tarefas">
      {dias.map(d => (
        <div key={d.iso} style={{ flex: 1, minWidth: 0, textAlign: "center", padding: "5px 0 6px", borderRadius: 9, background: d.hoje ? "rgba(0,224,168,0.14)" : "rgba(255,255,255,0.04)", border: d.hoje ? "1px solid rgba(0,224,168,0.5)" : "1px solid transparent" }}>
          <div style={{ fontSize: 9, fontWeight: 700, color: "#8B93B5", letterSpacing: 0.3 }}>{d.dow}</div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, fontWeight: 700, color: d.hoje ? "#4FE8BE" : "#D5DBF5", lineHeight: 1.3 }}>{d.dia}</div>
          <div style={{ height: 8, display: "flex", alignItems: "center", justifyContent: "center" }}>
            {d.n > 0 && <span style={{ width: d.n > 1 ? 12 : 7, height: 7, borderRadius: 4, background: "#FFC53D" }} />}
          </div>
        </div>
      ))}
    </div>
  );
}

function ProximosDiasCard({ itens, onOpenItem, onAcao, hidden, central = false, grupoUnico = null, vazioTexto = null }) {
  const [aberto, setAberto] = useState(false);
  const limite = central ? 999 : 3;
  const vis = aberto ? itens : itens.slice(0, limite);
  const cor = (t) => t === "Fixo" ? C.amber : t === "Tarefa" ? C.emerald : C.lavender;
  const grupoDe = (it) => { if (grupoUnico) return grupoUnico; const d = diasAteVenc(it.data); return d < 0 ? (central ? "Atrasados" : "Já venceu") : d === 0 ? "Hoje" : d === 1 ? "Amanhã" : d <= 7 ? "Esta semana" : (central ? "Semana seguinte" : "Depois"); };
  return (
    <div className={central ? undefined : "in"} style={central ? { padding: 0 } : { background: C.card, boxShadow: shadowSm, borderRadius: 18, padding: "16px 18px", marginBottom: 20 }}>
      {!central && <div className={itens.length > limite ? "tap" : undefined} onClick={itens.length > limite ? () => setAberto(a => !a) : undefined} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: itens.length ? 10 : 6, cursor: itens.length > limite ? "pointer" : "default" }}>
        <Calendar size={16} color={C.lavender} />
        <span style={{ fontWeight: 700, fontSize: 13, color: C.ink, textTransform: "uppercase", letterSpacing: 0.4 }}>Próximos dias</span>
        {itens.length > 0 && <span style={{ fontSize: 11, fontWeight: 700, color: C.muted, background: C.border, borderRadius: 10, padding: "1px 8px" }}>{itens.length}</span>}
        {itens.length > limite && <ChevronDown size={18} color={C.muted} style={{ marginLeft: "auto", transform: aberto ? "rotate(180deg)" : "none", transition: "transform .2s ease" }} />}
      </div>}
      {itens.length === 0 ? (
        <div style={{ fontSize: 12.5, color: C.muted }}>{central ? (vazioTexto || "Nada atrasado nem agendado pelos próximos 14 dias. Tudo em dia ✓") : "Nada a lançar nem agendado pelos próximos dias."}</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column" }}>
          {vis.map((it, i) => (
            <React.Fragment key={it.id}>
            {(i === 0 || grupoDe(vis[i - 1]) !== grupoDe(it)) && (
              <div style={{ fontSize: 10.5, fontWeight: 700, color: grupoDe(it) === "Atrasados" && !grupoUnico ? C.coral : C.faint, textTransform: "uppercase", letterSpacing: 0.6, padding: i === 0 ? "0 0 4px" : "12px 0 4px" }}>{grupoDe(it)}</div>
            )}
            <div className="tap" onClick={() => onOpenItem && onOpenItem(it)}
              style={{ display: "flex", alignItems: "center", gap: 12, padding: "9px 0", cursor: "pointer", borderBottom: i < vis.length - 1 ? `1px solid ${C.border}` : "none" }}>
              <div style={{ width: 46, flexShrink: 0, textAlign: "center", borderRadius: 10, background: C.border, padding: "5px 0" }}>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 15, fontWeight: 700, color: C.ink, lineHeight: 1.1 }}>{String(it.data).slice(8, 10)}</div>
                <div style={{ fontSize: 9.5, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: 0.4 }}>{MESES_ABR[Number(String(it.data).slice(5, 7)) - 1]}</div>
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{it.label}</div>
                <div style={{ fontSize: 11.5, color: it.atrasado && !grupoUnico ? C.coral : C.faint, marginTop: 2 }}>
                  <span style={{ color: cor(it.tipo), fontWeight: 700 }}>{it.tipo}</span> · {it.sub}
                </div>
              </div>
              {it.valor ? <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: C.muted, flexShrink: 0 }}>{hidden ? "•••" : fmtBRL(it.valor)}</span> : null}
              {onAcao && (
                <button type="button" onClick={(e) => { e.stopPropagation(); onAcao(it); }}
                  style={{ flexShrink: 0, height: 28, borderRadius: 14, border: `1px solid ${C.emerald}`, background: C.emeraldSoft, color: C.emerald, cursor: "pointer", fontSize: 11.5, fontWeight: 700, padding: "0 11px" }}>
                  {it.tipo === "Fixo" ? "Lançar" : it.tipo === "Custo" ? "Pagar" : "Concluir"}
                </button>
              )}
            </div>
            </React.Fragment>
          ))}
        </div>
      )}
      {itens.length > limite && (
        <button type="button" onClick={() => setAberto(a => !a)} style={{ marginTop: 8, border: "none", background: "transparent", color: C.muted, cursor: "pointer", fontSize: 12, fontWeight: 700, padding: "4px 2px" }}>
          {aberto ? "Ver menos" : `+${itens.length - limite} mais`}
        </button>
      )}
    </div>
  );
}

// Central de avisos: atrasados e o que vence nos próximos 14 dias (+ tarefas paradas há muito tempo, recolhidas).
// Fica no nível do App para abrir por cima de qualquer aba, sem trocar de aba.
function CentralAvisosModal({ central, hidden = false, onClose, onNavigate, onConcluirDemanda }) {
  const [paradosAberto, setParadosAberto] = useState(false);
  return (
    <Modal title="Central de avisos" onClose={onClose}>
      <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 10 }}>Atrasados e tudo o que vence nos próximos 14 dias.</div>
      <ProximosDiasCard central itens={central.itens} hidden={hidden}
        vazioTexto={central.parados && central.parados.length > 0 ? "Nada urgente nem agendado pelos próximos 14 dias ✓" : null}
        onAcao={async (it) => {
          if (it.tipo === "Fixo") { onClose(); onNavigate && onNavigate("custos", null, { lancarFixo: it.label, mes: currentMonth() }); }
          else if (it.tipo === "Custo") { onClose(); onNavigate && onNavigate("custos", it.rawId, { pagar: true }); }
          else if (onConcluirDemanda && await confirmar(`Marcar "${it.label}" como Concluída?`, { okLabel: "Concluir" })) onConcluirDemanda(it.rawId);
        }}
        onOpenItem={(it) => { onClose(); it.fixo ? (onNavigate && onNavigate("custos", null, { abrirFixos: true })) : (onNavigate && onNavigate(it.tab, it.rawId)); }} />
      {central.parados && central.parados.length > 0 && (
        <div style={{ marginTop: central.itens.length ? 18 : 12, paddingTop: 12, borderTop: central.itens.length ? `1px solid ${C.border}` : "none" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <button type="button" className="tap" onClick={() => setParadosAberto(a => !a)} aria-expanded={paradosAberto}
              style={{ flex: 1, minWidth: 0, display: "flex", alignItems: "center", gap: 8, border: "none", background: "transparent", color: C.ink, cursor: "pointer", fontFamily: "inherit", padding: "4px 0", textAlign: "left" }}>
              <span style={{ fontWeight: 700, fontSize: 12.5, textTransform: "uppercase", letterSpacing: 0.4 }}>Paradas há mais de {PARADA_DIAS} dias</span>
              <span style={{ fontSize: 11, fontWeight: 700, color: C.muted, background: C.border, borderRadius: 10, padding: "1px 8px" }}>{central.parados.length}</span>
              <ChevronDown size={18} color={C.muted} style={{ marginLeft: "auto", transform: paradosAberto ? "rotate(180deg)" : "none", transition: "transform .2s ease" }} />
            </button>
          </div>
          <div style={{ fontSize: 11.5, color: C.faint, margin: "2px 0 6px" }}>Não entram na contagem de avisos. Conclua, adie ou apague quando puder.</div>
          {paradosAberto && <>
            <ProximosDiasCard central grupoUnico="Paradas" itens={central.parados} hidden={hidden}
              onAcao={async (it) => { if (onConcluirDemanda && await confirmar(`Marcar "${it.label}" como Concluída?`, { okLabel: "Concluir" })) onConcluirDemanda(it.rawId); }}
              onOpenItem={(it) => { onClose(); onNavigate && onNavigate(it.tab, it.rawId); }} />
            <button type="button" className="tap" onClick={() => { onClose(); onNavigate && onNavigate("demandas", null, { adiarParadas: true }); }}
              style={{ marginTop: 10, height: 34, borderRadius: 17, border: `1px solid ${C.border}`, background: "transparent", color: C.inkSoft, cursor: "pointer", fontSize: 12, fontWeight: 700, padding: "0 14px", fontFamily: "inherit" }}>
              Adiar todas de uma vez…
            </button>
          </>}
        </div>
      )}
    </Modal>
  );
}

// Blocos do modal "Observações" da Projeção de fechamento.
function ObsBloco({ titulo, total = null, hidden = false, nota = null, acao = null, children }) {
  return (
    <div style={{ marginBottom: 18 }}>
      <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 10, marginBottom: 2 }}>
        <span style={{ fontWeight: 700, fontSize: 12.5, color: C.ink, textTransform: "uppercase", letterSpacing: 0.4 }}>{titulo}</span>
        {total !== null && <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 13, color: C.ink }}>{hidden ? "R$ •••••" : fmtBRL(total)}</span>}
      </div>
      {nota && <div style={{ fontSize: 11.5, color: C.faint, lineHeight: 1.45, marginBottom: 6 }}>{nota}</div>}
      <div>{children}</div>
      {acao}
    </div>
  );
}
function ObsLinha({ nome, detalhe = null, valor = null, hidden = false }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "8px 0", borderTop: `1px solid ${C.border}` }}>
      <div style={{ minWidth: 0 }}>
        <div style={{ fontSize: 13.5, fontWeight: 600, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{nome}</div>
        {detalhe && <div style={{ fontSize: 11.5, color: C.faint }}>{detalhe}</div>}
      </div>
      {valor !== null && <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: C.inkSoft, flexShrink: 0 }}>{hidden ? "R$ •••••" : fmtBRL(valor)}</span>}
    </div>
  );
}

function DashboardTab({ custos, depositos = [], receitas, demandas, onNavigate, onConcluirDemanda, inicioContagem, fixos = [], topo = null, onAbrirCentral, central = { itens: [], parados: [], nUrgentes: 0 } }) {
  const [hidden, setHidden] = useState(false);
  const [saldoAberto, setSaldoAberto] = useState(false);
  const [analise, setAnalise] = useState("evolucao");
  const [obsAberto, setObsAberto] = useState(false);
  const [pulsarAtencao, setPulsarAtencao] = useState(false);
  const qtdAtencaoAnt = useRef(null);
  // Mês exibido nos cartões do Início (padrão: o mês corrente). "Precisa de atenção" e "Agenda" continuam sendo de hoje.
  // null = acompanha o mês corrente (se o app ficar aberto na virada do mês, ele muda sozinho).
  const [mesEscolhido, setMesEscolhido] = useState(null);
  const mesAtual = mesEscolhido || currentMonth();
  const ehMesCorrente = !mesEscolhido;
  const escolherMes = (m) => setMesEscolhido(m === currentMonth() ? null : m);
  const [showSeletorMes, setShowSeletorMes] = useState(false);
  // Fixos que vencem em até 7 dias (ou já venceram) e ainda não foram lançados neste mês. Só exibição: nada é gravado.
  const fixosAVencer = useMemo(() => {
    if (!ehMesCorrente || fixos.length === 0) return [];
    const limite = new Date().getDate() + 7;
    return fixosPendentes(fixos, custos, mesAtual).faltam.filter(f => (f.dia || 0) <= limite);
  }, [fixos, custos, mesAtual, ehMesCorrente]);
  const semLancamentos = ehMesCorrente && !custos.some(c => c.mes === mesAtual) && !receitas.some(r => r.mes === mesAtual);
  const ano = parseInt(mesAtual.split("-")[0], 10);
  const mesNumVisto = parseInt(mesAtual.split("-")[1], 10);
  const mesAnterior = shiftMonth(mesAtual, -1);
  const proximos30Todos = useAtencao(custos, demandas, 30);
  // "Precisa de atenção" = só o que já passou do prazo. "Agenda" = tudo que vem pela frente
  // (consultas, demandas e custos a vencer) em ordem de data, sem repetir itens entre os dois cartões.
  const atencao = useMemo(() => proximos30Todos.filter(i => i.atrasado), [proximos30Todos]);
  const proximos30 = useMemo(() => proximos30Todos.filter(i => !i.atrasado), [proximos30Todos]);
  const ritmo = useRitmoMes(custos, mesAtual);
  useEffect(() => {
    if (qtdAtencaoAnt.current !== null && central.nUrgentes > qtdAtencaoAnt.current) { setPulsarAtencao(true); const t = setTimeout(() => setPulsarAtencao(false), 1600); qtdAtencaoAnt.current = central.nUrgentes; return () => clearTimeout(t); }
    qtdAtencaoAnt.current = central.nUrgentes;
  }, [central.nUrgentes]);

  const receitasComparativo = useMemo(() => {
    // Grouped by título (like the donut next to it): two "BC" records become one "BC" bar.
    const map = new Map();
    receitas.filter(r => r.mes === mesAtual).forEach(r => {
      const nome = (r.descricao || "Outras").toUpperCase();
      const g = map.get(nome) || { nome, Previsto: 0, Recebido: 0, Pendente: false };
      g.Previsto += Number(r.valorPrevisto || 0);
      g.Recebido += Number(r.valorRecebido || 0);
      if (r.status !== "Recebido") g.Pendente = true;
      map.set(nome, g);
    });
    return [...map.values()].sort((a, b) => b.Previsto - a.Previsto).slice(0, 6);
  }, [receitas, mesAtual]);

  const saldosPorMes = useMemo(() => MESES.map((_, idx) => {
    const mesKey = `${ano}-${String(idx + 1).padStart(2, "0")}`;
    const pago = custos.filter(c => c.mes === mesKey).reduce((s, c) => s + Number(c.valorPago || 0), 0);
    const recebido = receitas.filter(r => r.mes === mesKey).reduce((s, r) => s + Number(r.valorRecebido || 0), 0);
    return recebido - pago;
  }), [custos, receitas, ano]);

  const kpis = useMemo(() => {
    const cs = custos.filter(c => c.mes === mesAtual);
    const rs = receitas.filter(r => r.mes === mesAtual);
    const totalPrevisto = cs.reduce((s, c) => s + Number(c.valorPrevisto || 0), 0);
    const totalPago = cs.reduce((s, c) => s + Number(c.valorPago || 0), 0);
    const totalRecebido = rs.reduce((s, r) => s + Number(r.valorRecebido || 0), 0);
    return { totalPrevisto, totalPago, totalRecebido, saldo: totalRecebido - totalPago, pctUsado: totalPrevisto > 0 ? (totalPago / totalPrevisto) * 100 : 0 };
  }, [custos, receitas, mesAtual]);

  // Trend comparisons use "até o mesmo dia do mês" on both sides — comparing
  // this month's running total (partial, since the month isn't over) against
  // a FULL previous month always makes the current month look artificially
  // better early on and worse near the end. Same fix as Ritmo do mês, applied
  // to the Pago/Recebido/Saldo trend indicators.
  const diaCorte = ritmo.diaRef;
  const diaCorteAnterior = Math.min(diaCorte, daysInMonthOf(mesAnterior));
  const pagoAteDiaAtual = useMemo(() => pagoAteDia(custos, mesAtual, diaCorte), [custos, mesAtual, diaCorte]);
  const pagoAteDiaAnterior = useMemo(() => pagoAteDia(custos, mesAnterior, diaCorteAnterior), [custos, mesAnterior, diaCorteAnterior]);
  const recebidoAteDiaAtual = useMemo(() => recebidoAteDia(receitas, mesAtual, diaCorte), [receitas, mesAtual, diaCorte]);
  const recebidoAteDiaAnterior = useMemo(() => recebidoAteDia(receitas, mesAnterior, diaCorteAnterior), [receitas, mesAnterior, diaCorteAnterior]);
  const pctChange = (atual, anterior) => (anterior > 0 ? ((atual - anterior) / anterior) * 100 : null);
  const pagoTrendPct = pctChange(pagoAteDiaAtual, pagoAteDiaAnterior);
  const recebidoTrendPct = pctChange(recebidoAteDiaAtual, recebidoAteDiaAnterior);
  const saldoAteDiaAtual = recebidoAteDiaAtual - pagoAteDiaAtual;
  const saldoAteDiaAnterior = recebidoAteDiaAnterior - pagoAteDiaAnterior;

  // Saldo por fonte: current month's net movement per payment source (custos'
  // "fonte" field, receitas' "contaDestino" field), only fontes that actually
  // had movement this month. The warning sums ALL pending custos tied to a
  // fonte (not just the next one), so a fonte with several upcoming bills
  // doesn't under-warn.
  const pixMes = useMemo(() => resumoPix(custos, mesAtual), [custos, mesAtual]);
  const saldoPorFonte = useMemo(() => {
    const map = {};
    const get = (f) => (map[f] || (map[f] = { recebido: 0, pago: 0, pendente: 0, itensPendentes: [], meios: {}, depositoSaida: 0, depositoEntrada: 0 }));
    custos.filter(c => c.mes === mesAtual).forEach(c => {
      const f = get(contaDaFonte(c.fonte || "—"));
      if (c.status === "Pago") {
        f.pago += Number(c.valorPago || 0);
        if (c.fonte !== contaDaFonte(c.fonte)) f.meios[c.fonte] = (f.meios[c.fonte] || 0) + Number(c.valorPago || 0);
      } else {
        const v = Number(c.valorPrevisto || 0);
        f.pendente += v;
        f.itensPendentes.push({ descricao: c.descricao, valor: v, vencimento: c.vencimento });
      }
    });
    receitas.filter(r => r.mes === mesAtual).forEach(r => {
      get(contaDaFonte(r.contaDestino || "—")).recebido += Number(r.valorRecebido || 0);
    });
    // Received "Depósito …": leaves Dinheiro and enters the receita's conta (net zero overall).
    depositos.filter(r => r.mes === mesAtual && r.status === "Recebido").forEach(r => {
      const v = Number(r.valorRecebido || 0);
      const dest = get(contaDaFonte(r.contaDestino || "—"));
      dest.recebido += v; dest.depositoEntrada += v;
      const din = get("Dinheiro");
      din.pago += v; din.depositoSaida += v;
    });
    return Object.entries(map)
      .map(([fonte, v]) => {
        const saldo = v.recebido - v.pago;
        const saldoProjetado = saldo - v.pendente;
        return { fonte, ...v, saldo, saldoProjetado, vaiNegativar: v.pendente > 0 && saldoProjetado < 0 };
      })
      .filter(f => f.recebido > 0 || f.pago > 0 || f.pendente > 0)
      .sort((a, b) => (a.vaiNegativar !== b.vaiNegativar ? (a.vaiNegativar ? -1 : 1) : b.saldo - a.saldo));
  }, [custos, depositos, receitas, mesAtual]);

  const porFonte = useMemo(() => {
    const map = {};
    custos.filter(c => c.mes === mesAtual).forEach(c => {
      const v = Number(c.valorPago ?? c.valorPrevisto ?? 0);
      const k = contaDaFonte(c.fonte);
      const m = map[k] || (map[k] = { value: 0, pend: 0 });
      m.value += v;
      if (c.valorPago == null) m.pend += v;
    });
    return Object.entries(map).map(([name, m]) => ({ name, value: m.value, pend: m.pend })).sort((a, b) => b.value - a.value);
  }, [custos, mesAtual]);
  const meiosMes = useMemo(() => notasMeios(custos.filter(c => c.mes === mesAtual)), [custos, mesAtual]);

  const porCategoria = useMemo(() => {
    const map = {};
    receitas.filter(r => r.mes === mesAtual).forEach(r => {
      const v = Number(r.valorRecebido ?? r.valorPrevisto ?? 0);
      const key = (r.descricao || "Outras").toUpperCase();
      map[key] = (map[key] || 0) + v;
    });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [receitas, mesAtual]);

  const catColor = useMemo(() => {
    const titulos = [...new Set(receitas.filter(r => r.mes === mesAtual).map(r => (r.descricao || "Outras").toUpperCase()))].sort();
    return (name) => PALETTE_IN[Math.max(0, titulos.indexOf(name)) % PALETTE_IN.length];
  }, [receitas, mesAtual]);

  const tendencia = useMemo(() => {
    const out = [];
    for (let i = 5; i >= 0; i--) {
      const d = new Date(ano, mesNumVisto - 1 - i, 1);
      const mesKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const pago = custos.filter(c => c.mes === mesKey).reduce((s, c) => s + Number(c.valorPago || 0), 0);
      const recebido = receitas.filter(r => r.mes === mesKey).reduce((s, r) => s + Number(r.valorRecebido || 0), 0);
      out.push({ mes: MESES_ABR[d.getMonth()], mesKey, sel: mesKey === mesAtual, Pago: pago, Recebido: recebido, fora: !!inicioContagem && mesKey < inicioContagem });
    }
    // Meses anteriores ao início da contagem não entram quando já há 3 ou mais meses válidos (sem faixas vazias).
    const dentro = out.filter(d => !d.fora);
    return dentro.length >= 3 ? dentro : out;
  }, [custos, receitas, inicioContagem, ano, mesNumVisto]);

  // Ranked by descrição, not by record: a cost split across fontes (one record
  // per fonte, same descrição) counts as ONE cost with its combined value.
  const topCustos = useMemo(() => {
    const grupos = new Map();
    for (const c of custos.filter(c => c.mes === mesAtual)) {
      const key = (c.descricao || "").trim().toLowerCase();
      const v = Number(c.valorPago ?? c.valorPrevisto ?? 0);
      const g = grupos.get(key) || { id: key, descricao: c.descricao, valor: 0, fontes: [] };
      g.valor += v;
      const fc = contaDaFonte(c.fonte);
      if (fc && !g.fontes.includes(fc)) g.fontes.push(fc);
      grupos.set(key, g);
    }
    return [...grupos.values()].sort((a, b) => b.valor - a.valor).slice(0, 7);
  }, [custos, mesAtual]);
  const topMax = Math.max(1, ...topCustos.map(c => c.valor));

  const projecao = useMemo(() => {
    const cs = custos.filter(c => c.mes === mesAtual);
    const rs = receitas.filter(r => r.mes === mesAtual);
    const pago = cs.reduce((t, c) => t + Number(c.valorPago || 0), 0);
    const aPagar = cs.filter(c => c.status !== "Pago").reduce((t, c) => t + Number(c.valorPrevisto || 0), 0);
    const recebido = rs.reduce((t, r) => t + Number(r.valorRecebido || 0), 0);
    const aReceber = rs.filter(r => r.status !== "Recebido").reduce((t, r) => t + Math.max(0, Number(r.valorPrevisto || 0) - Number(r.valorRecebido || 0)), 0);
    const previstoRec = rs.reduce((t, r) => t + Number(r.valorPrevisto || 0), 0);
    return { pago, aPagar, recebido, aReceber, previstoRec, saldoProjetado: (recebido + aReceber) - (pago + aPagar) };
  }, [custos, receitas, mesAtual]);
  // Fixos ainda não lançados neste mês entram na projeção como ESTIMATIVA (usa o valor pago esperado, se houver).
  // Fixos sem valor (ex.: Unimed, Gibson) ficam de fora e são avisados.
  const fixosFaltam = useMemo(() => ehMesCorrente ? fixosPendentes(fixos, custos, mesAtual).faltam : [], [fixos, custos, mesAtual, ehMesCorrente]);
  // Sem valor cadastrado, a estimativa usa o lançamento do mês anterior com o mesmo nome (é o mesmo valor que o "Lançar" sugere).
  const valorAnteriorFixo = (f) => { const a = custos.find(c => c.mes === mesAnterior && normDesc(c.descricao) === normDesc(f.descricao)); return a ? Number(a.valorPrevisto || a.valorPago || 0) : 0; };
  const valorFixoCad = (f) => Number(f.valorPago || f.valor || 0);
  const valorFixo = (f) => valorFixoCad(f) > 0 ? valorFixoCad(f) : valorAnteriorFixo(f);
  const fixosEst = fixosFaltam.reduce((t, f) => t + valorFixo(f), 0);
  const fixosSemValor = fixosFaltam.filter(f => !(valorFixo(f) > 0));
  const fixosPeloAnterior = fixosFaltam.filter(f => !(valorFixoCad(f) > 0) && valorAnteriorFixo(f) > 0);
  // Receitas habituais (mesmo título do mês anterior) ainda não lançadas neste mês entram como ESTIMATIVA de entrada.
  const receitasEst = useMemo(() => {
    if (!ehMesCorrente) return { total: 0, itens: [] };
    const tit = (r) => (r.descricao || "Outras").toUpperCase();
    const jaTem = new Set(receitas.filter(r => r.mes === mesAtual).map(tit));
    const vistos = new Set();
    const itens = [];
    receitas.filter(r => r.mes === mesAnterior).forEach(r => {
      const t = tit(r);
      if (jaTem.has(t) || vistos.has(t)) return;
      vistos.add(t);
      const v = receitas.filter(x => x.mes === mesAnterior && tit(x) === t).reduce((a, x) => a + Number(x.valorRecebido ?? x.valorPrevisto ?? 0), 0);
      if (v > 0) itens.push({ titulo: t, valor: v });
    });
    return { total: itens.reduce((a, i) => a + i.valor, 0), itens };
  }, [receitas, mesAtual, mesAnterior, ehMesCorrente]);
  const receitasEstTotal = receitasEst.total;
  const saldoEstimado = projecao.saldoProjetado + receitasEstTotal - fixosEst;
  const projMax = Math.max(1, projecao.recebido + projecao.aReceber + receitasEstTotal, projecao.pago + projecao.aPagar + fixosEst);
  const projQuitado = projecao.aPagar === 0 && fixosEst === 0 && receitasEstTotal === 0 && projecao.aReceber === 0 && (projecao.pago + projecao.recebido) > 0;
  const projEstimada = fixosEst > 0 || receitasEstTotal > 0;
  // Notas da projeção (abrem no modal "Observações"): receitas esperadas, fixos estimados e fixos sem valor.
  const nObs = [receitasEstTotal > 0, fixosEst > 0, fixosSemValor.length > 0].filter(Boolean).length;
  const projTemDados = (projecao.pago + projecao.recebido + projecao.aPagar + projecao.aReceber + fixosEst + receitasEstTotal) > 0;

  // "Próximos dias": fixos a lançar + custos e tarefas que vêm pela frente, tudo em ordem de data.
  const proximosDias = useMemo(() => {
    const lista = proximos30.map(i => ({
      id: i.id, tipo: i.tab === "demandas" ? "Tarefa" : "Custo", label: i.label, data: i.data,
      sub: rotuloAgenda(i), tab: i.tab, rawId: i.rawId
    }));
    fixosFaltam.forEach(f => {
      const data = vencimentoDoFixo(f, mesAtual);
      const d = diasAteVenc(data);
      lista.push({
        id: `f-${f.id || f.descricao}`, tipo: "Fixo", label: f.descricao, data, fixo: true, valor: valorFixo(f) > 0 ? valorFixo(f) : null,
        sub: d < 0 ? `a lançar · venceu dia ${data.slice(8, 10)}` : d === 0 ? "a lançar · vence hoje" : d === 1 ? "a lançar · vence amanhã" : `a lançar · vence em ${d} dias`
      });
    });
    return lista.sort((a, b) => String(a.data).localeCompare(String(b.data)) || String(a.label).localeCompare(String(b.label)));
  }, [proximos30, fixosFaltam, mesAtual]);

  const comparativo = useMemo(() => {
    const soma = (mes, fonte) => custos.filter(c => c.mes === mes && contaDaFonte(c.fonte) === fonte).reduce((t, c) => t + Number(c.valorPago ?? c.valorPrevisto ?? 0), 0);
    return FONTES.map(f => ({ fonte: f, Atual: soma(mesAtual, f), Anterior: soma(mesAnterior, f) })).filter(r => r.Atual > 0 || r.Anterior > 0);
  }, [custos, mesAtual, mesAnterior]);
  // Ranked by the change in R$ (a % on a tiny base, like R$ 299, is just noise).
  const variacoes = comparativo
    .map(r => ({ fonte: r.fonte, delta: r.Atual - r.Anterior, pct: r.Anterior > 0 ? ((r.Atual - r.Anterior) / r.Anterior) * 100 : null }))
    .filter(r => Math.abs(r.delta) > 0.005)
    .sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));

  // Months before the chosen "contar a partir de" month don't enter the sum
  // (and aren't drawn), so months with incomplete history can't distort it.
  const inicioIdx = inicioIdxDoAno(inicioContagem, ano);
  const acumulado = useMemo(() => {
    let cum = 0;
    const mesIdx = mesNumVisto - 1;
    return MESES_ABR.map((m, i) => {
      if (i < inicioIdx) return { mes: m, Acumulado: null };
      cum += saldosPorMes[i];
      return { mes: m, Acumulado: i <= mesIdx ? cum : null };
    });
  }, [saldosPorMes, inicioIdx, mesNumVisto]);
  const acumAtual = [...acumulado].reverse().find(d => d.Acumulado !== null)?.Acumulado ?? 0;
  const acumColor = acumAtual < 0 ? C.coral : C.emerald;

  // Curva do saldo dia a dia (recebido − pago até cada dia): este mês contra o mês anterior.
  const curva = useMemo(() => {
    const nA = daysInMonthOf(mesAtual), nB = daysInMonthOf(mesAnterior);
    const hojeDia = ehMesCorrente ? Math.min(new Date().getDate(), nA) : nA;
    const rows = [];
    for (let d = 1; d <= Math.max(nA, nB); d++) {
      rows.push({
        dia: d,
        atual: d <= hojeDia ? recebidoAteDia(receitas, mesAtual, d) - pagoAteDia(custos, mesAtual, d) : null,
        anterior: d <= nB ? recebidoAteDia(receitas, mesAnterior, d) - pagoAteDia(custos, mesAnterior, d) : null
      });
    }
    return { rows, hojeDia };
  }, [custos, receitas, mesAtual, mesAnterior, ehMesCorrente]);

  // Movimento acumulado por conta (entradas − saídas pagas, mês a mês desde "contar a partir de").
  // É movimento, não saldo bancário: não há saldo inicial cadastrado.
  const movConta = useMemo(() => {
    const cum = { "Conta Salário": 0, "Conta Corrente": 0, "Dinheiro": 0 };
    const usadas = new Set();
    const rows = [];
    for (let i = inicioIdx; i < mesNumVisto; i++) {
      const mk = `${ano}-${String(i + 1).padStart(2, "0")}`;
      custos.filter(c => c.mes === mk && c.status === "Pago").forEach(c => { const k = contaDaFonte(c.fonte); if (k in cum) { cum[k] -= Number(c.valorPago || 0); usadas.add(k); } });
      receitas.filter(r => r.mes === mk).forEach(r => { const k = contaDaFonte(r.contaDestino); if (k in cum) { cum[k] += Number(r.valorRecebido || 0); usadas.add(k); } });
      depositos.filter(r => r.mes === mk && r.status === "Recebido").forEach(r => {
        const v = Number(r.valorRecebido || 0); const k = contaDaFonte(r.contaDestino);
        if (k in cum) { cum[k] += v; usadas.add(k); }
        cum["Dinheiro"] -= v; usadas.add("Dinheiro");
      });
      rows.push({ mes: MESES_ABR[i], ...cum });
    }
    return { rows, contas: ["Conta Salário", "Conta Corrente", "Dinheiro"].filter(k => usadas.has(k)) };
  }, [custos, receitas, depositos, inicioIdx, mesNumVisto, ano]);

  const mesLabel = monthLabel(mesAtual);

  // Números-chave de cada aba das Análises.
  const tilesAnalise = (() => {
    if (analise === "evolucao") {
      let melhor = null;
      for (let i = inicioIdx; i < mesNumVisto; i++) if (melhor === null || saldosPorMes[i] > saldosPorMes[melhor]) melhor = i;
      return [
        { rotulo: "Saldo acumulado", valor: fmtBRL(acumAtual), sub: `desde ${MESES[inicioIdx].toLowerCase()}`, cor: acumAtual < 0 ? C.coral : C.emerald, mask: true },
        { rotulo: "Média paga por mês", valor: ritmo.disponivel ? fmtBRL(ritmo.mediaMensal) : "—", sub: ritmo.disponivel ? `últimos ${ritmo.nMeses} meses` : "sem histórico", mask: true },
        ...(melhor !== null ? [{ rotulo: "Melhor mês", valor: MESES[melhor], sub: `saldo ${fmtBRL(saldosPorMes[melhor])}`, mask: true }] : [])
      ];
    }
    if (analise === "custos") {
      const top = topCustos[0];
      return [
        { rotulo: "Falta pagar", valor: fmtBRL(projecao.aPagar), sub: `${projecao.aPagar > 0 ? "ainda previsto" : "tudo pago"}${fixosFaltam.length > 0 ? ` · ${fixosFaltam.length} ${fixosFaltam.length === 1 ? "fixo" : "fixos"} a lançar` : ""}`, cor: projecao.aPagar > 0 ? C.amber : C.emerald, mask: true },
        ...(top ? [{ rotulo: "Maior custo", valor: top.descricao, sub: `${fmtBRL(top.valor)}${kpis.totalPago > 0 ? ` · ${((top.valor / kpis.totalPago) * 100).toFixed(0)}% do pago` : ""}`, mask: true }] : []),
        ...(porFonte[0] ? [{ rotulo: "Fonte principal", valor: porFonte[0].name, sub: `${pctDe(porFonte[0].value, porFonte)} dos custos` }] : [])
      ];
    }
    const abaixo = Math.max(0, projecao.previstoRec - projecao.recebido - projecao.aReceber);
    return [
      { rotulo: "Previsto no mês", valor: fmtBRL(projecao.previstoRec), sub: `${fmtBRL(kpis.totalRecebido)} já recebidos`, mask: true },
      { rotulo: "Falta entrar", valor: fmtBRL(projecao.aReceber), sub: `${projecao.aReceber > 0 ? "ainda previsto" : abaixo > 0.005 ? `recebido ${fmtBRL(abaixo)} abaixo do previsto` : "tudo recebido"}${receitasEstTotal > 0 ? `\n+ ${fmtBRL(receitasEstTotal)} esperadas a lançar` : ""}`, mask: true },
      ...(porCategoria[0] ? [{ rotulo: "Maior receita", valor: porCategoria[0].name, sub: `${fmtBRL(porCategoria[0].value)} · ${pctDe(porCategoria[0].value, porCategoria)}`, mask: true }] : [])
    ];
  })();

  return (
    <div>
      {obsAberto && (
        <Modal title="Observações da projeção" onClose={() => setObsAberto(false)}>
          <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 14, lineHeight: 1.45 }}>O que entrou como estimativa no saldo de {mesLabel}, além do que já está lançado. Ao lançar cada item, ele deixa de ser estimativa e passa a contar como previsto ou pago.</div>
          {receitasEstTotal > 0 && (
            <ObsBloco titulo="Receitas esperadas" total={receitasEstTotal} hidden={hidden} nota="Receitas do mês anterior que ainda não foram lançadas neste mês, contadas como entrada.">
              {receitasEst.itens.map(i => (
                <ObsLinha key={i.titulo} hidden={hidden} nome={i.titulo.toLowerCase().replace(/(^|\s)(\S)/g, (m, sp, ch) => sp + ch.toUpperCase())} detalhe="valor do mês anterior" valor={i.valor} />
              ))}
            </ObsBloco>
          )}
          {fixosEst > 0 && (
            <ObsBloco titulo="Fixos ainda não lançados" total={fixosEst} hidden={hidden} nota="Entram como saída pelo valor cadastrado no fixo ou, sem ele, pelo lançamento do mês anterior.">
              {fixosFaltam.filter(f => valorFixo(f) > 0).map(f => (
                <ObsLinha key={f.id || f.descricao} hidden={hidden} nome={f.descricao} detalhe={`${f.dia ? `dia ${f.dia} · ` : ""}${valorFixoCad(f) > 0 ? "valor cadastrado" : "pelo valor do mês anterior"}`} valor={valorFixo(f)} />
              ))}
            </ObsBloco>
          )}
          {fixosSemValor.length > 0 && (
            <ObsBloco titulo="Fora da estimativa" nota="Sem valor cadastrado e sem lançamento no mês anterior, esses fixos não pesam nos números. Cadastre o valor no fixo para eles entrarem."
              acao={(
                <button type="button" className="tap" onClick={() => { setObsAberto(false); onNavigate && onNavigate("custos", null, { abrirFixos: true }); }}
                  style={{ marginTop: 10, height: 34, borderRadius: 17, border: `1px solid ${C.border}`, background: "transparent", color: C.inkSoft, cursor: "pointer", fontSize: 12, fontWeight: 700, padding: "0 14px", fontFamily: "inherit" }}>
                  Cadastrar valor nos Fixos ›
                </button>
              )}>
              {fixosSemValor.map(f => (
                <ObsLinha key={f.id || f.descricao} nome={f.descricao} detalhe={`${f.dia ? `dia ${f.dia} · ` : ""}sem valor cadastrado`} />
              ))}
            </ObsBloco>
          )}
          <div style={{ height: 8 }} />
        </Modal>
      )}

      <BalanceHero
        pix={pixMes}
        saldo={kpis.saldo}
        trendAtual={saldoAteDiaAtual}
        trendAnterior={semLancamentos ? null : saldoAteDiaAnterior}
        mesLabel={mesLabel}
        hidden={hidden}
        onToggle={() => setHidden(h => !h)}
        saldosPorMes={saldosPorMes}
        inicioIdx={inicioIdx}
        mesIdxVisto={mesNumVisto - 1}
        ehMesCorrente={ehMesCorrente}
        onPickMes={(i) => escolherMes(`${ano}-${String(i + 1).padStart(2, "0")}`)}
        onAbrirSeletor={() => setShowSeletorMes(true)}
        semLancamentos={semLancamentos}
        mesAnteriorLabel={monthLabel(mesAnterior)}
        onVerAnterior={() => escolherMes(mesAnterior)}
        topo={topo}
        curva={curva}
        atencao={{ n: central.nUrgentes, parados: (central.parados || []).length, onOpen: () => onAbrirCentral && onAbrirCentral(hidden), pulsar: pulsarAtencao }}
        situacao={!ehMesCorrente ? [] : [
          ...(ritmo.disponivel ? [{ t: `Ritmo ${ritmo.pct.toFixed(0)}% do normal`, cor: Math.round(ritmo.pct) > 150 ? "#FF8A76" : Math.round(ritmo.pct) > 100 ? "#FFC53D" : "#4FE8BE" }] : []),
          ...(fixosEst > 0 ? [{ t: `Faltam ${fmtBRL(fixosEst)} em fixos ›`, cor: "#B8BFDB", onClick: () => onNavigate && onNavigate("custos", null, { abrirFixos: true }) }] : []),
          ...(projQuitado || !projTemDados ? [] : [{ t: `Fecha em ${fmtBRL(saldoEstimado)} (estimativa)`, cor: saldoEstimado < 0 ? "#FF8A76" : "#4FE8BE" }])
        ]}
        porFonte={saldoPorFonte}
        contasLinha={saldoPorFonte.filter(f => f.fonte !== "—").slice().sort((a, b) => FONTES.indexOf(a.fonte) - FONTES.indexOf(b.fonte))}
        expanded={saldoAberto}
        onToggleExpanded={() => setSaldoAberto(a => !a)}
      />
      {showSeletorMes && <SeletorMes mes={mesAtual} onPick={escolherMes} onClose={() => setShowSeletorMes(false)} />}

      <StatRow stats={[
        { icon: <Wallet size={12} />, label: "Previsto", value: hidden ? "•••" : fmtBRL(kpis.totalPrevisto) },
        {
          icon: <Check size={12} />, label: "Pago", value: hidden ? "•••" : fmtBRL(kpis.totalPago),
          trend: !semLancamentos && pagoTrendPct !== null ? { pct: pagoTrendPct, good: pagoTrendPct <= 0 } : null
        },
        {
          icon: <PiggyBank size={12} />, label: "Recebido", value: hidden ? "•••" : fmtBRL(kpis.totalRecebido),
          trend: !semLancamentos && recebidoTrendPct !== null ? { pct: recebidoTrendPct, good: recebidoTrendPct >= 0 } : null
        },
      ]} />

      {ehMesCorrente && <ProximosDiasCard itens={proximosDias} hidden={hidden}
        onAcao={async (it) => {
          if (it.tipo === "Fixo") onNavigate && onNavigate("custos", null, { lancarFixo: it.label, mes: mesAtual });
          else if (it.tipo === "Custo") onNavigate && onNavigate("custos", it.rawId, { pagar: true });
          else if (onConcluirDemanda && await confirmar(`Marcar "${it.label}" como Concluída?`, { okLabel: "Concluir" })) onConcluirDemanda(it.rawId);
        }}
        onOpenItem={(it) => it.fixo ? (onNavigate && onNavigate("custos", null, { abrirFixos: true })) : (onNavigate && onNavigate(it.tab, it.rawId))} />}

      <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 18, padding: "20px 22px", marginBottom: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 16 }}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink }}>Projeção de fechamento — {mesLabel}</div>
            <div style={{ fontSize: 11.5, color: C.faint, marginTop: 3 }}>{projQuitado ? "Tudo o que foi lançado já foi pago e recebido" : projEstimada ? "Estimativa: o previsto mais fixos e receitas habituais ainda não lançados" : "Se tudo o que está previsto for pago e recebido"}</div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 24, fontWeight: 700, color: saldoEstimado < 0 ? C.coral : C.emerald }}>
              {hidden ? "R$ •••••" : fmtBRL(saldoEstimado)}
            </div>
            <div style={{ fontSize: 11, color: C.faint }}>{projQuitado ? "saldo do mês" : projEstimada ? "saldo estimado no fim do mês" : "saldo previsto no fim do mês"}</div>
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <ProjBar label="Entradas" done={projecao.recebido} pending={projecao.aReceber + receitasEstTotal} color={C.emerald} max={projMax} hidden={hidden} />
          <ProjBar label="Saídas" done={projecao.pago} pending={projecao.aPagar + fixosEst} color={C.coral} max={projMax} hidden={hidden} />
        </div>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginTop: 12 }}>
          {!projQuitado ? (
            <div style={{ display: "flex", gap: 16, fontSize: 11, color: C.faint, flexWrap: "wrap" }}>
              <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 10, height: 10, borderRadius: 3, background: C.muted }} />já realizado</span>
              <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 10, height: 10, borderRadius: 3, background: `repeating-linear-gradient(135deg, ${C.muted}88 0 3px, ${C.muted}33 3px 6px)` }} />ainda previsto</span>
            </div>
          ) : <span />}
          {nObs > 0 && (
            <button type="button" className="tap" onClick={() => setObsAberto(true)} aria-label={`Observações sobre a projeção (${nObs})`}
              style={{ marginLeft: "auto", border: "none", background: "transparent", color: C.inkSoft, cursor: "pointer", fontFamily: "inherit", fontSize: 11, fontWeight: 800, letterSpacing: 0.6, textTransform: "uppercase", padding: "4px 0" }}>
              Observações ({nObs}) ›
            </button>
          )}
        </div>
        <div style={{ marginTop: 18 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
            <span style={{ fontWeight: 700, fontSize: 12.5, color: C.ink }}>Ritmo do mês</span>
            {ritmo.disponivel && (
              <span style={{ fontSize: 12.5, fontWeight: 700, color: Math.round(ritmo.pct) > 150 ? C.coral : Math.round(ritmo.pct) > 100 ? C.amber : C.emerald }}>
                {ritmo.pct.toFixed(0)}% do ritmo normal
              </span>
            )}
          </div>
          {ritmo.disponivel ? (
            <>
              <div style={{ fontSize: 11, color: C.faint, marginBottom: 8 }}>
                Comparado ao que você normalmente já tinha pago até o dia {ritmo.diaRef}, nos últimos {ritmo.nMeses === 1 ? "mês" : `${ritmo.nMeses} meses`} — não à média do mês inteiro
              </div>
              <div style={{ position: "relative", height: 10, borderRadius: 5, background: C.border, margin: "6px 0 18px" }}>
                <div style={{ height: "100%", width: `${Math.min(ritmo.pct, 200) / 2}%`, borderRadius: 5, background: Math.round(ritmo.pct) > 150 ? "linear-gradient(90deg,#FFB929,#FF5A68)" : Math.round(ritmo.pct) > 100 ? "linear-gradient(90deg,#FFC53D,#FFB929)" : "linear-gradient(90deg,#00E0A8,#22D3EE)" }} />
                <div style={{ position: "absolute", left: "50%", top: -4, bottom: -4, width: 2, borderRadius: 1, background: C.ink, opacity: 0.55 }} />
                <span style={{ position: "absolute", left: "50%", top: 14, transform: "translateX(-50%)", fontSize: 10, color: C.faint }}>normal</span>
              </div>
              <div style={{ fontSize: 11, color: C.faint, marginTop: 8 }}>
                {ehMesCorrente ? "Pago até agora" : "Pago no mês"}: <strong style={{ color: C.ink }}>{hidden ? "•••" : fmtBRL(ritmo.pagoAteAgora)}</strong>
                {" · "}Normal até o dia {ritmo.diaRef}: <strong style={{ color: C.ink }}>{hidden ? "•••" : fmtBRL(ritmo.mediaAteODia)}</strong>
                {Math.abs(Number(ritmo.mediaMensal) - Number(ritmo.mediaAteODia)) > 0.005 && <>{" · "}Média mensal: <strong style={{ color: C.ink }}>{hidden ? "•••" : fmtBRL(ritmo.mediaMensal)}</strong></>}
              </div>
            </>
          ) : (
            <div style={{ fontSize: 12, color: C.muted, padding: "6px 0" }}>
              Ainda não há meses anteriores com custos pagos suficientes para comparar o ritmo deste mês.
            </div>
          )}
        </div>
      </div>

      <div role="tablist" aria-label="Análises" style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap", marginBottom: 14 }}>
        <span style={{ fontWeight: 700, fontSize: 13, color: C.ink, textTransform: "uppercase", letterSpacing: 0.4, marginRight: 6 }}>Análises</span>
        {[["evolucao", "Evolução"], ["custos", "Custos"], ["receitas", "Receitas"]].map(([k, t]) => (
          <button key={k} type="button" role="tab" aria-selected={analise === k} onClick={() => setAnalise(k)}
            style={{ border: `1px solid ${analise === k ? C.emerald : C.border}`, background: analise === k ? C.emeraldSoft : "transparent", color: analise === k ? C.emerald : C.muted, borderRadius: 999, padding: "7px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>{t}</button>
        ))}
      </div>

      <TilesAnalise itens={tilesAnalise} hidden={hidden} />

      <div className="grid2" style={{ marginBottom: 20, alignItems: "start" }}>
        {/* Orçamento agora vive dentro do card de projeção */}

                {analise === "custos" && porFonte.length > 1 && (
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "40ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 12 }}>Custos por fonte — {mesLabel}</div>
          <BarraEmpilhada dados={porFonte} corDe={fonteColor} hidden={hidden} nota={meiosMes.join(" · ")} totalLabel="Pago + a pagar" />
        </div>
        )}

                {analise === "receitas" && porCategoria.length > 1 && (
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "120ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 12 }}>Receitas por título — {mesLabel}</div>
          <BarraEmpilhada dados={porCategoria} corDe={catColor} hidden={hidden} />
        </div>
        )}

        {analise === "evolucao" && (
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "160ms", display: "flex", flexDirection: "column" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
            <span style={{ fontWeight: 700, fontSize: 13.5, color: C.ink }}>Pago x Recebido <span style={{ fontWeight: 500, fontSize: 11.5, color: C.faint }}>· toque num mês para abri-lo</span></span>
            <span style={{ display: "flex", gap: 10, fontSize: 11, color: C.muted, fontWeight: 600 }}>
              <span style={{ display: "flex", alignItems: "center", gap: 4 }}><span style={{ width: 8, height: 8, borderRadius: 2, background: "#E8434F" }} />Pago</span>
              <span style={{ display: "flex", alignItems: "center", gap: 4 }}><span style={{ width: 8, height: 8, borderRadius: 2, background: "#00B98A" }} />Recebido</span>
            </span>
          </div>
          {tendencia.some(d => d.fora) && (
            <div style={{ fontSize: 11, color: C.faint, marginBottom: 6 }}>Meses esmaecidos são anteriores ao início da contagem (lançamentos incompletos).</div>
          )}
          <div style={{ width: "100%", flex: 1, minHeight: 170 }}>
            <ResponsiveContainer width="100%" height="100%" minHeight={170}>
              <BarChart data={tendencia} barGap={4} barCategoryGap="24%" margin={{ top: 18 }}>
                <defs>
                  <linearGradient id="gPago" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={C.gradOut[0]} /><stop offset="100%" stopColor={C.gradOut[1]} />
                  </linearGradient>
                  <linearGradient id="gRecebido" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={C.gradIn[0]} /><stop offset="100%" stopColor={C.gradIn[1]} />
                  </linearGradient>
                </defs>
                <CartesianGrid vertical={false} stroke={C.border} strokeDasharray="3 4" />
                <XAxis dataKey="mes" tick={{ fontSize: 11, fill: C.faint }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis hide />
                <Tooltip content={<RichTooltip modo="saldo" colorMap={{ Pago: "#E8434F", Recebido: "#00B98A" }} />} />
                <Bar dataKey="Pago" fill="url(#gPago)" radius={[4, 4, 0, 0]} onClick={(d) => { const k = d && (d.mesKey || (d.payload && d.payload.mesKey)); if (k) escolherMes(k); }} cursor="pointer">
                  {tendencia.map((d, i) => <Cell key={i} fillOpacity={d.fora ? 0.3 : d.sel ? 1 : 0.72} />)}
                  <LabelList dataKey="Pago" position="top" formatter={rotuloK} style={{ fontSize: 10, fill: C.muted }} />
                </Bar>
                <Bar dataKey="Recebido" fill="url(#gRecebido)" radius={[4, 4, 0, 0]} onClick={(d) => { const k = d && (d.mesKey || (d.payload && d.payload.mesKey)); if (k) escolherMes(k); }} cursor="pointer">
                  {tendencia.map((d, i) => <Cell key={i} fillOpacity={d.fora ? 0.3 : d.sel ? 1 : 0.72} />)}
                  <LabelList dataKey="Recebido" position="top" formatter={rotuloK} style={{ fontSize: 10, fill: C.muted }} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        )}

        {analise === "custos" && (
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "200ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 10 }}>Maiores custos — {mesLabel}</div>
          {topCustos.length === 0 ? (
            <div style={{ color: C.muted, fontSize: 13, padding: "16px 0", textAlign: "center" }}>Sem custos lançados neste mês.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {topCustos.map((c, i) => (
                <div key={c.id}>
                  <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginBottom: 5 }}>
                    <span style={{ width: 16, fontSize: 11, fontWeight: 700, color: C.faint }}>{i + 1}</span>
                    <span style={{ flex: 1, minWidth: 0, fontSize: 13, fontWeight: 600, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      {c.descricao}
                      {c.fontes.length > 1 && <span style={{ color: C.faint, fontWeight: 500, fontSize: 11 }}> · {c.fontes.length} fontes</span>}
                    </span>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: i === 0 ? C.coral : C.muted, fontWeight: i === 0 ? 700 : 500 }}>{hidden ? "•••" : fmtBRL(c.valor)}</span>
                  </div>
                  <div style={{ height: 5, borderRadius: 3, background: C.border, marginLeft: 24, overflow: "hidden" }}>
                    <div style={{ width: `${(c.valor / topMax) * 100}%`, height: "100%", borderRadius: 3, background: c.fontes.length === 1 ? fonteColor(c.fontes[0]) : C.lavender, opacity: i === 0 ? 1 : 0.8 }} />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
        )}

        {analise === "custos" && (
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "240ms", display: "flex", flexDirection: "column" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 4 }}>Este mês x mês anterior — por fonte</div>
          <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 10 }}>{monthLabel(mesAnterior)} (cinza) x {mesLabel}</div>
          {comparativo.length === 0 ? (
            <div style={{ color: C.muted, fontSize: 13, padding: "16px 0", textAlign: "center" }}>Sem custos nos dois meses para comparar.</div>
          ) : (
            <>
              <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
                {comparativo.map(r => {
                  const mx = Math.max(1, ...comparativo.map(x => Math.max(x.Atual, x.Anterior)));
                  const d = r.Atual - r.Anterior;
                  return (
                    <div key={r.fonte}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 8, marginBottom: 6 }}>
                        <span style={{ fontSize: 13, fontWeight: 600, color: C.ink }}>{r.fonte}</span>
                        {Math.abs(d) > 0.005 && <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, fontWeight: 700, color: d > 0 ? C.coral : C.emerald }}>{d > 0 ? "+" : "−"}{hidden ? "•••" : fmtBRL(Math.abs(d))}</span>}
                      </div>
                      <div style={{ height: 5, borderRadius: 3, background: C.border, overflow: "hidden", marginBottom: 4 }}>
                        <div style={{ width: `${(r.Anterior / mx) * 100}%`, height: "100%", borderRadius: 3, background: C.barCinza }} />
                      </div>
                      <div style={{ height: 8, borderRadius: 4, background: C.border, overflow: "hidden" }}>
                        <div style={{ width: `${(r.Atual / mx) * 100}%`, height: "100%", borderRadius: 4, background: fonteColor(r.fonte) }} />
                      </div>
                      <div style={{ display: "flex", justifyContent: "space-between", fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: C.faint, marginTop: 3 }}>
                        <span>{hidden ? "•••" : fmtBRL(r.Anterior)}</span><span style={{ color: C.muted }}>{hidden ? "•••" : fmtBRL(r.Atual)}</span>
                      </div>
                    </div>
                  );
                })}
              </div>
              {variacoes[0] && ehMesCorrente && new Date().getDate() < 15 && (
                <div style={{ fontSize: 12, color: C.faint, marginTop: 8 }}>
                  Mês em andamento: a comparação ainda é parcial (até o dia {new Date().getDate()}).
                </div>
              )}
              {variacoes[0] && !(ehMesCorrente && new Date().getDate() < 15) && (
                <div style={{ fontSize: 12, color: C.muted, marginTop: 8 }}>
                  Maior variação: <strong style={{ color: C.ink }}>{variacoes[0].fonte}</strong>{" "}
                  <strong style={{ color: variacoes[0].delta > 0 ? C.coral : C.emerald }}>
                    {variacoes[0].delta > 0 ? "+" : "−"}{hidden ? "•••" : fmtBRL(Math.abs(variacoes[0].delta))}
                  </strong>
                  {variacoes[0].pct !== null && <span style={{ color: C.faint }}> ({variacoes[0].pct > 0 ? "+" : ""}{variacoes[0].pct.toFixed(0)}%)</span>}
                </div>
              )}
            </>
          )}
        </div>
        )}

        {analise === "evolucao" && movConta.rows.length >= 2 && movConta.contas.length > 0 && (
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "280ms", display: "flex", flexDirection: "column" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 4 }}>Movimento acumulado por conta</div>
          <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 10 }}>Entradas menos saídas de cada conta desde {MESES[inicioIdx].toLowerCase()} (movimento, sem saldo inicial)</div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", fontSize: 11.5, fontWeight: 600, color: C.muted, marginBottom: 6 }}>
            {movConta.contas.map(k => {
              const ult = movConta.rows[movConta.rows.length - 1][k];
              return (
                <span key={k} style={{ display: "flex", alignItems: "center", gap: 5 }}>
                  <span style={{ width: 9, height: 9, borderRadius: 3, background: fonteColor(k) }} />
                  {k === "Conta Salário" ? "CS" : k === "Conta Corrente" ? "CC" : k}{" "}
                  <strong style={{ color: ult < 0 ? C.coral : C.ink, fontVariantNumeric: "tabular-nums" }}>{hidden ? "•••" : fmtBRL(ult)}</strong>
                </span>
              );
            })}
          </div>
          <div style={{ width: "100%", flex: 1, minHeight: 180 }}>
            <ResponsiveContainer width="100%" height="100%" minHeight={180}>
              <AreaChart data={movConta.rows} margin={{ top: 8, right: 10, left: 10 }}>
                <defs>
                  {movConta.contas.map((k, mi) => (
                    <linearGradient key={k} id={`gMov${mi}`} x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stopColor={fonteColor(k)} stopOpacity={0.2} /><stop offset="100%" stopColor={fonteColor(k)} stopOpacity={0.02} />
                    </linearGradient>
                  ))}
                </defs>
                <CartesianGrid vertical={false} stroke={C.border} strokeDasharray="3 4" />
                <XAxis dataKey="mes" tick={{ fontSize: 11, fill: C.faint }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis hide />
                <ReferenceLine y={0} stroke={C.border} strokeDasharray="3 3" />
                <Tooltip content={<TooltipPT />} cursor={{ stroke: C.border, strokeWidth: 1 }} />
                {movConta.contas.map((k, mi) => (
                  <Area key={k} type="monotone" dataKey={k} stroke={fonteColor(k)} strokeWidth={2.5} fill={`url(#gMov${mi})`} isAnimationActive={false}
                    dot={(p) => p.index === movConta.rows.length - 1 && p.cx != null
                      ? <circle key={"d" + p.index} cx={p.cx} cy={p.cy} r={4.5} fill={fonteColor(k)} stroke={C.card} strokeWidth={2} />
                      : <g key={"d" + p.index} />}
                    activeDot={{ r: 5 }} />
                ))}
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
        )}
        {analise === "receitas" && (
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "320ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 4 }}>Receitas previstas x recebidas — {mesLabel}</div>
          <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 10 }}>O que já entrou, sobre o que estava previsto</div>
          {receitasComparativo.length === 0 ? (
            <div style={{ color: C.muted, fontSize: 13, padding: "16px 0", textAlign: "center" }}>Sem receitas lançadas neste mês.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              {receitasComparativo.map(r => {
                const pct = r.Previsto > 0 ? Math.min(100, (r.Recebido / r.Previsto) * 100) : (r.Recebido > 0 ? 100 : 0);
                const falta = r.Previsto - r.Recebido;
                return (
                  <div key={r.nome}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 8, marginBottom: 6 }}>
                      <span style={{ fontSize: 13, fontWeight: 600, color: C.ink, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{tituloBonito(r.nome)}</span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: C.muted, flexShrink: 0 }}>
                        <strong style={{ color: r.Recebido > 0 ? C.emerald : C.faint }}>{hidden ? "•••" : fmtBRL(r.Recebido)}</strong>
                        <span style={{ color: C.faint }}> de {hidden ? "•••" : fmtBRL(r.Previsto)}</span>
                      </span>
                    </div>
                    <div style={{ height: 7, borderRadius: 4, background: C.border, overflow: "hidden" }}>
                      <div style={{ width: `${pct}%`, height: "100%", borderRadius: 4, background: "linear-gradient(90deg, #00E0A8, #22D3EE)" }} />
                    </div>
                    {falta > 0.005 && <div style={{ fontSize: 11, color: C.faint, marginTop: 4 }}>{r.Pendente ? "falta " : ""}{hidden ? "•••" : fmtBRL(falta)}{r.Pendente ? "" : " abaixo do previsto"}</div>}
                  </div>
                );
              })}
            </div>
          )}
        </div>
        )}

        {/* alignSelf: start — a short list must not be stretched to the height of its neighbour */}
        {analise === "custos" && (
        <div style={{ alignSelf: "start", display: "flex", flexDirection: "column", gap: 20 }}>
          <PixMesCard pix={pixMes} mesLabel={mesLabel} hidden={hidden} />
        </div>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Custos Mensais                                                     */
/* ------------------------------------------------------------------ */

function CardList({ rows, onRemove, onEdit, empty, render }) {
  if (rows.length === 0) {
    return (
      <div style={{
        border: `1px dashed ${C.border}`, borderRadius: 14, padding: "32px 16px",
        textAlign: "center", color: C.muted, fontSize: 13.5, background: C.card
      }}>{empty}</div>
    );
  }
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {rows.map((row, i) => (
        <div key={row.id} className="tap in" onClick={() => onEdit && onEdit(row)} style={{
          display: "flex", alignItems: "center", gap: 10, background: C.card,
          boxShadow: shadowSm, borderRadius: 16, padding: "14px 14px",
          cursor: onEdit ? "pointer" : "default", position: "relative",
          animationDelay: `${Math.min(i, 8) * 40}ms`
        }}>
          <div style={{ position: "absolute", left: 0, top: 12, bottom: 12, width: 3, borderRadius: 3, background: STATUS_ACCENT[row.status] || C.border }} />
          {render(row)}
          <button onClick={async (e) => { e.stopPropagation(); if (await confirmar("Remover este item?", { perigo: true, okLabel: "Remover" })) onRemove(row.id); }} title="Remover" style={{
            border: "none", background: "transparent", color: C.faint, cursor: "pointer",
            padding: 4, display: "flex", alignItems: "center", flexShrink: 0, marginRight: 30
          }}>
            <Trash2 size={16} />
          </button>
        </div>
      ))}
    </div>
  );
}

// Costs paid across several fontes are stored as one record per fonte with the same
// descrição. Inside a status group they are gathered into one block (first-appearance
// order) so the list can draw an outline around them. Single costs stay as they are.
function agruparDivididos(items) {
  const chave = (c) => `${(c.descricao || "").trim().toLowerCase()}|${c.mes}`;
  const mapa = new Map();
  for (const c of items) {
    const k = chave(c);
    if (!mapa.has(k)) mapa.set(k, []);
    mapa.get(k).push(c);
  }
  return [...mapa.entries()].map(([k, lista]) => ({
    key: k, itens: lista, dividido: lista.length > 1,
    total: lista.reduce((s, c) => s + Number(c.valorPago ?? c.valorPrevisto ?? 0), 0)
  }));
}

function CustosTab({ custos, onAdd, onRemove, onUpdate, jumpTarget, onConsumeJump, fixos = [], fixosErro = false, onAddFixo, onUpdateFixo, onRemoveFixo }) {
  const [showFixos, setShowFixos] = useState(false);
  const [mesFiltro, setMesFiltro] = useState(currentMonth());
  const [showSeletorMesC, setShowSeletorMesC] = useState(false);
  const [showFontes, setShowFontes] = useState(false);
  useMesCorrenteAuto(setMesFiltro);
  const [fonteFiltro, setFonteFiltro] = useState("Todas");
  // Old records may still carry a fonte that is no longer offered (BB, plain "Transferência"); keep them filterable.
  const fontesLegadas = useMemo(() => [...new Set(custos.map(c => c.fonte).filter(f => f && !FONTES_FORM.includes(f)))], [custos]);
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [saving, setSaving] = useState(false);

  const filtrados = useMemo(() => {
    const list = custos.filter(c => c.mes === mesFiltro && fonteCombina(c, fonteFiltro));
    return [...list].sort((a, b) => {
      // Sem data de pagamento (a pagar / atrasados): por vencimento; no mesmo dia, maior valor primeiro, depois nome.
      if (!a.dataPagamento && !b.dataPagamento) {
        return String(a.vencimento || "9999").localeCompare(String(b.vencimento || "9999"))
          || (Number(b.valorPrevisto || 0) - Number(a.valorPrevisto || 0))
          || String(a.descricao || "").localeCompare(String(b.descricao || ""));
      }
      if (!a.dataPagamento) return 1;
      if (!b.dataPagamento) return -1;
      return a.dataPagamento.localeCompare(b.dataPagamento);
    });
  }, [custos, mesFiltro, fonteFiltro]);

  const totals = useMemo(() => {
    const totalPrevisto = filtrados.reduce((s,c) => s + Number(c.valorPrevisto||0), 0);
    const totalPago = filtrados.reduce((s,c) => s + Number(c.valorPago||0), 0);
    const pagos = filtrados.filter(c => c.status === "Pago").length;
    const pendentes = filtrados.filter(c => statusEfetivo(c) === "Pendente").length;
    const atrasados = filtrados.filter(c => statusEfetivo(c) === "Atrasado").length;
    const porFonte = {};
    filtrados.forEach(c => { const k = contaDaFonte(c.fonte); porFonte[k] = (porFonte[k]||0) + Number(c.valorPago||0); });
    // Largest fonte first — same ordering as the dashboard's "Custos por fonte".
    const porFonteOrdenado = Object.fromEntries(Object.entries(porFonte).sort((a, b) => b[1] - a[1]));
    const faltaPagar = filtrados.filter(c => c.status !== "Pago").reduce((s,c) => s + Number(c.valorPrevisto||0), 0);
    return { totalPrevisto, totalPago, diferenca: totalPago-totalPrevisto, faltaPagar, pagos, pendentes, atrasados, porFonte: porFonteOrdenado };
  }, [filtrados]);
  const serie = useMemo(() => serieSaidas(custos, mesFiltro, fonteFiltro), [custos, mesFiltro, fonteFiltro]);
  const meiosCustos = useMemo(() => notasMeios(filtrados), [filtrados]);
  const proximo = useMemo(() => proximoVencimento(filtrados), [filtrados]);
  // Maior custo do mês (fontes divididas somam pela descrição) — aparece na faixa do herói quando não há vencimento a mostrar.
  const maiorCusto = useMemo(() => {
    const m = new Map();
    filtrados.forEach(c => { const k = c.descricao || "—"; m.set(k, (m.get(k) || 0) + Number(c.valorPago ?? c.valorPrevisto ?? 0)); });
    let best = null;
    m.forEach((v, k) => { if (!best || v > best.valor) best = { descricao: k, valor: v }; });
    const total = [...m.values()].reduce((a, b) => a + b, 0);
    return best && best.valor > 0 ? { ...best, pct: total > 0 ? Math.round((best.valor / total) * 100) : 0 } : null;
  }, [filtrados]);

  function blankCusto(mes) {
    return { mes, ano: parseInt(mes.split("-")[0],10), descricao: "", vencimento: "", dataPagamento: "",
      valorPrevisto: "", valorPago: "", fonte: FONTES[0], status: "Pendente", obs: "" };
  }
  const [form, setForm] = useState(blankCusto(mesFiltro));

  const openNew = () => { setEditingId(null); setForm(blankCusto(mesFiltro)); setShowForm(true); };
  const formDeCusto = (c) => ({
    mes: c.mes, ano: c.ano, descricao: c.descricao, vencimento: c.vencimento || "",
    dataPagamento: c.dataPagamento || "", valorPrevisto: fmtValorInput(c.valorPrevisto),
    valorPago: fmtValorInput(c.valorPago), fonte: c.fonte, status: c.status, obs: c.obs || ""
  });
  const openEdit = (c) => {
    setEditingId(c.id);
    setForm(formDeCusto(c));
    setShowForm(true);
  };

  // "Lançar" on a fixo: opens the normal form pre-filled (nothing is saved until the person saves).
  const lancarFixo = (fx, pagoHoje = false, mesAlvo = null) => {
    const mf = mesAlvo || mesFiltro;
    setEditingId(null);
    // Sem valor/fonte no fixo: sugere os do lançamento do mês anterior com o mesmo nome (o formulário mostra; só grava no Salvar).
    const [ay, am] = mf.split("-").map(Number);
    const mesAnt = am === 1 ? `${ay - 1}-12` : `${ay}-${String(am - 1).padStart(2, "0")}`;
    const ant = custos.find(c => c.mes === mesAnt && normDesc(c.descricao) === normDesc(fx.descricao));
    const valorBase = fx.valor || (ant ? Number(ant.valorPrevisto || ant.valorPago || 0) : 0);
    const pagoEsp = fx.valorPago || valorBase || 0;
    setForm({ ...blankCusto(mf), descricao: fx.descricao, vencimento: vencimentoDoFixo(fx, mf),
      valorPrevisto: valorBase ? fmtValorInput(valorBase) : "", fonte: fx.fonte || (ant && ant.fonte) || FONTES[0],
      ...(pagoHoje ? { status: "Pago", dataPagamento: todayISO(), valorPago: pagoEsp ? fmtValorInput(pagoEsp) : "" } : {}) });
    setShowForm(true);
  };

  useEffect(() => {
    if (!jumpTarget) return;
    const item = custos.find(c => c.id === jumpTarget.id);
    if (item) {
      setMesFiltro(item.mes); openEdit(item);
      // "Pagar" vindo do Início: abre o formulário já com Pago/hoje/valor previsto. Só grava quando você tocar em Salvar.
      if (jumpTarget.pagar) setForm(f => ({ ...f, status: "Pago", dataPagamento: todayISO(), valorPago: f.valorPago || (item.valorPrevisto ? fmtValorInput(item.valorPrevisto) : "") }));
    }
    if (jumpTarget.lancarFixo) {
      const fx = fixos.find(x => x.descricao === jumpTarget.lancarFixo);
      const m = jumpTarget.mes || currentMonth();
      if (fx) { setMesFiltro(m); lancarFixo(fx, false, m); }
    }
    if (jumpTarget.novo) openNew();
    if (jumpTarget.abrirFixos) setShowFixos(true);
    onConsumeJump();
  }, [jumpTarget]);

  const submit = async (e) => {
    e.preventDefault();
    if (saving) return;
    if (!form.descricao || !form.vencimento || form.valorPrevisto === "") return;
    const valorPrevistoNum = parseValorBR(form.valorPrevisto);
    const valorPagoNum = form.valorPago === "" ? null : parseValorBR(form.valorPago);
    if (!isValorValido(valorPrevistoNum) || (valorPagoNum !== null && !isValorValido(valorPagoNum))) {
      alert("O valor precisa ser maior que zero.");
      return;
    }
    if (!(await confirmDateSanity(form.vencimento, "data de vencimento"))) return;
    // Campos-chave: um custo Pago precisa de Valor Pago e Data de Pagamento, senão some dos saldos e do Anual.
    if (form.status === "Pago") {
      const faltam = [];
      if (valorPagoNum === null) faltam.push("Valor Pago");
      if (!form.dataPagamento) faltam.push("Data Pagamento");
      if (faltam.length) { alert(`Este custo está como Pago. Preencha: ${faltam.join(" e ")}.\n\nSe ainda não foi pago, mude o Status para Pendente.`); return; }
    }
    // Data de pagamento: erro de ano/mês (ex.: 03/12 em vez de 03/11) não pode passar batido. Só pergunta quando a data ou o status mudou.
    const origCusto = editingId ? custos.find(c => c.id === editingId) : null;
    if (form.status === "Pago" && form.dataPagamento) {
      if (!(await confirmDateSanity(form.dataPagamento, "data de pagamento"))) return;
      const mudouPgto = !origCusto || origCusto.status !== "Pago" || origCusto.dataPagamento !== form.dataPagamento;
      if (mudouPgto && form.dataPagamento > todayISO() && !(await confirmar(`A data do pagamento (${fmtDate(form.dataPagamento)}) ainda não chegou. Marcar como Pago mesmo assim?`, { titulo: "Data no futuro", okLabel: "Está correta" }))) return;
    }
    // Nome repetido: lançamento novo com a mesma descrição e a mesma fonte no mês (divisão em fontes diferentes não dispara).
    if (!editingId) {
      const igual = custos.find(c => c.mes === form.mes && normDesc(c.descricao) === normDesc(form.descricao) && c.fonte === form.fonte);
      if (igual && !(await confirmar(`Já existe "${igual.descricao}" em ${monthLabel(igual.mes)} com a fonte ${igual.fonte} (${fmtBRL(igual.valorPrevisto)}, ${igual.status}).\n\nSe for outra parte de um pagamento dividido, use outra fonte. Lançar mesmo assim?`, { titulo: "Possível duplicado", okLabel: "Lançar mesmo assim" }))) return;
    }
    // Valor muito acima do mesmo custo no mês anterior (erro de vírgula, ex.: 17600 em vez de 176,00). Só dispara em salto de 3x ou mais e de pelo menos R$ 300.
    if (!origCusto || Number(origCusto.valorPrevisto) !== valorPrevistoNum) {
      const mesAntC = shiftMonth(form.mes, -1);
      const baseC = Math.max(0, ...custos.filter(c => c.mes === mesAntC && normDesc(c.descricao) === normDesc(form.descricao)).map(c => Number(c.valorPrevisto || c.valorPago || 0)));
      if (baseC > 0 && valorPrevistoNum >= baseC * 3 && valorPrevistoNum - baseC >= 300
        && !(await confirmar(`O valor previsto (${fmtBRL(valorPrevistoNum)}) é mais de 3 vezes o de ${monthLabel(mesAntC)} (${fmtBRL(baseC)}) para "${form.descricao}".\n\nO valor está certo?`, { titulo: "Valor muito acima do normal", okLabel: "Está certo" }))) return;
    }
    // Confirmação da fonte: ao marcar Pago (ou trocar a fonte de um custo já Pago), mostra o resumo antes de gravar.
    if (form.status === "Pago") {
      const orig = origCusto;
      if (!orig || orig.status !== "Pago" || orig.fonte !== form.fonte) {
        const difPg = valorPagoNum - valorPrevistoNum;
        const linhaPrev = Math.abs(difPg) > 0.005 ? `\nPrevisto: ${fmtBRL(valorPrevistoNum)} (${difPg > 0 ? "+" : "−"}${fmtBRL(Math.abs(difPg))})` : "";
        const resumo = `${form.descricao}\nValor pago: ${fmtBRL(valorPagoNum)}${linhaPrev}\nFonte: ${form.fonte}\nData do pagamento: ${fmtDate(form.dataPagamento)}`;
        if (!(await confirmar(`${resumo}\n\nConfirma todos os dados?`, { titulo: "Marcar como Pago?", okLabel: "Sim, confirmar" }))) return;
      }
    }
    const valorPagoFinal = valorPagoNum;
    const payload = { ...form, valorPrevisto: valorPrevistoNum, valorPago: valorPagoFinal };
    setSaving(true);
    try {
      if (editingId) { await onUpdate(editingId, payload); } else { await onAdd(payload); }
      setForm(blankCusto(mesFiltro));
      setEditingId(null);
      setShowForm(false);
    } finally {
      setSaving(false);
    }
  };

  const grupos = useMemo(() => {
    const ordem = ["Atrasado", "Pendente", "Pago"];
    return ordem
      .map(status => {
        const items = filtrados.filter(c => statusEfetivo(c) === status);
        return { status, items, blocos: agruparDivididos(items) };
      })
      .filter(g => g.items.length > 0);
  }, [filtrados]);

  const GROUP_LABEL = { "Atrasado": "Atrasados", "Pendente": "A pagar", "Pago": "Pagos" };

  // Setas do modal de edição: percorrem os custos na MESMA ordem em que aparecem na lista.
  const ordemLista = useMemo(() => grupos.flatMap(g => g.blocos.flatMap(b => b.itens)), [grupos]);
  const idxEdicao = editingId ? ordemLista.findIndex(c => c.id === editingId) : -1;
  const navegarCusto = async (dir) => {
    const alvo = ordemLista[idxEdicao + dir];
    if (!alvo) return;
    const original = ordemLista[idxEdicao];
    const mudou = original && JSON.stringify(formDeCusto(original)) !== JSON.stringify(form);
    if (mudou && !(await confirmar("Há alterações não salvas neste custo. Descartar e ir para o outro?", { okLabel: "Descartar" }))) return;
    openEdit(alvo);
  };

  // One cost row. `embutido` = drawn inside the outline of a split cost (flatter, no own shadow).
  const renderLinhaCusto = (c, i, accent, embutido) => (
    <div key={c.id} className="tap in" onClick={() => openEdit(c)} style={{
      display: "flex", alignItems: "center", gap: 12, background: C.card,
      boxShadow: embutido ? "none" : shadowSm, padding: "14px 14px 14px 15px",
      borderRadius: embutido ? 13 : 14,
      cursor: "pointer", position: "relative", animationDelay: `${Math.min(i,6)*30}ms`
    }}>
      <div style={{ position: "absolute", left: 0, top: 10, bottom: 10, width: 3, borderRadius: 3, background: accent }} />
      <Wallet size={16} color={C.faint} style={{ flexShrink: 0 }} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: 700, color: C.ink, fontSize: 15 }}>
          {c.descricao}
        </div>
        <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2, display: "flex", alignItems: "center", flexWrap: "wrap", gap: "0 6px" }}>
          <span>Venc. {fmtDate(c.vencimento)}{c.dataPagamento ? ` · Pago ${fmtDate(c.dataPagamento)}` : ""} ·</span>
          <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
            <span style={{ width: 8, height: 8, borderRadius: 3, background: fonteColor(c.fonte), flexShrink: 0 }} />
            {c.fonte}
          </span>
        </div>
        {c.status !== "Pago" && (() => {
          const d = diasAteVenc(c.vencimento);
          if (d === null || d > 7) return null;
          const atrasado = d < 0;
          return (
            <div style={{ fontSize: 11, fontWeight: 700, marginTop: 3, color: atrasado ? C.coral : C.amber }}>
              {atrasado ? fmtAtraso(-d) : d === 0 ? "Vence hoje" : `Vence em ${d} ${d === 1 ? "dia" : "dias"}`}
            </div>
          );
        })()}
      </div>
      <div style={{ flexShrink: 0, textAlign: "right" }}>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 17, color: C.ink }}>
          {fmtBRL(c.valorPago ?? c.valorPrevisto)}
        </div>
        {c.valorPago != null && Math.abs(Number(c.valorPago) - Number(c.valorPrevisto || 0)) > 0.005 && (
          <div style={{ fontSize: 10.5, color: C.faint, marginTop: 1 }}>previsto {fmtBRL(c.valorPrevisto)}</div>
        )}
        {c.status === "Pago" && c.valorPago == null && (
          <div style={{ fontSize: 10.5, fontWeight: 700, color: C.amberDeep, marginTop: 1 }}>sem valor pago</div>
        )}
      </div>
      <button onClick={async (e) => { e.stopPropagation(); if (await confirmar("Remover este custo?", { perigo: true, okLabel: "Remover" })) onRemove(c.id); }} title="Remover" style={{
        border: "none", background: "transparent", color: C.faint, cursor: "pointer",
        padding: 4, display: "flex", alignItems: "center", flexShrink: 0
    }}>
        <Trash2 size={15} />
      </button>
                      </div>
  );


  return (
    <div>
      <TabHero titulo="Custos" mesLabel={monthLabel(mesFiltro)} onLabelClick={() => setShowSeletorMesC(true)}
        destaque={(totals.pendentes + totals.atrasados > 0)
          ? { rotulo: "Falta pagar", valor: fmtBRL(totals.faltaPagar), cor: totals.atrasados > 0 ? "#FF8A76" : "#FFC53D", sub: `Pago ${fmtBRL(totals.totalPago)} de ${fmtBRL(totals.totalPrevisto)}` }
          : { rotulo: "Pago no mês", valor: fmtBRL(totals.totalPago), cor: "#4FE8BE", sub: `de ${fmtBRL(totals.totalPrevisto)} previsto${filtrados.length > 0 && totals.diferenca < -0.005 ? ` · economia de ${fmtBRL(-totals.diferenca)}` : totals.diferenca > 0.005 ? ` · ${fmtBRL(totals.diferenca)} acima do previsto` : ""}` }}
        chips={[
          ...(totals.pagos > 0 ? [{ t: `${totals.pagos} ${totals.pagos === 1 ? "pago" : "pagos"}`, cor: "#4FE8BE" }] : []),
          ...(totals.pendentes > 0 ? [{ t: `${totals.pendentes} ${totals.pendentes === 1 ? "pendente" : "pendentes"}`, cor: "#FFC53D" }] : []),
          ...(totals.atrasados > 0 ? [{ t: `${totals.atrasados} ${totals.atrasados === 1 ? "atrasado" : "atrasados"}`, cor: "#FF8A76" }] : [])
        ]}
        pct={totals.totalPrevisto > 0 ? (totals.totalPago / totals.totalPrevisto) * 100 : 0}
        pctLabel={totals.totalPrevisto > 0 ? `${Math.round((totals.totalPago / totals.totalPrevisto) * 100)}% do previsto pago` : null}
        faixa={proximo ? (
          <button type="button" className="tap" onClick={() => openEdit(proximo.item.orig || proximo.item)}
            style={{ width: "100%", display: "flex", alignItems: "center", gap: 10, textAlign: "left", fontFamily: "inherit", cursor: "pointer", padding: "8px 12px", borderRadius: 12, background: "rgba(255,255,255,0.05)", border: "1px solid rgba(255,255,255,0.10)" }}>
            <span style={{ width: 10, height: 10, borderRadius: 5, background: proximo.cor, boxShadow: `0 0 9px ${proximo.cor}`, flexShrink: 0 }} />
            <span style={{ flex: 1, minWidth: 0 }}>
              <span style={{ display: "block", fontSize: 10, fontWeight: 700, color: "#8B93B5", textTransform: "uppercase", letterSpacing: 0.5 }}>Próximo vencimento</span>
              <span style={{ display: "block", fontSize: 13, fontWeight: 700, color: "#F3F5FA", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{proximo.item.descricao}{proximo.item.vencimento ? ` · ${fmtDate(proximo.item.vencimento)}` : ""}</span>
            </span>
            <span style={{ textAlign: "right", flexShrink: 0 }}>
              <span style={{ display: "block", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 700, color: "#F3F5FA" }}>{fmtBRL(proximo.item.valorPrevisto)}</span>
              <span style={{ display: "block", fontSize: 11, fontWeight: 700, color: proximo.cor }}>{proximo.texto}</span>
            </span>
          </button>
        ) : maiorCusto ? (
          <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 12px", borderRadius: 12, background: "rgba(255,255,255,0.05)", border: "1px solid rgba(255,255,255,0.10)" }}>
            <span style={{ flex: 1, minWidth: 0 }}>
              <span style={{ display: "block", fontSize: 10, fontWeight: 700, color: "#8B93B5", textTransform: "uppercase", letterSpacing: 0.5 }}>Maior custo do mês</span>
              <span style={{ display: "block", fontSize: 13, fontWeight: 700, color: "#F3F5FA", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{maiorCusto.descricao}</span>
            </span>
            <span style={{ textAlign: "right", flexShrink: 0 }}>
              <span style={{ display: "block", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 700, color: "#F3F5FA" }}>{fmtBRL(maiorCusto.valor)}</span>
              <span style={{ display: "block", fontSize: 11, fontWeight: 700, color: "#8B93B5" }}>{maiorCusto.pct}% do mês</span>
            </span>
          </div>
        ) : null}
        acao={{ label: "Novo", onClick: openNew }}
        onPrev={() => setMesFiltro(m => shiftMonth(m, -1))} onNext={() => setMesFiltro(m => shiftMonth(m, 1))}
        rodapeExtra={(
          <button type="button" className="tap" onClick={() => setShowFontes(true)} aria-label="Filtrar por fonte de pagamento"
            style={{ display: "flex", alignItems: "center", gap: 6, height: 34, padding: "0 12px", borderRadius: 11, cursor: "pointer", fontFamily: "inherit", fontSize: 12.5, fontWeight: 700,
              border: `1px solid ${fonteFiltro !== "Todas" ? "#00E0A8" : "rgba(255,255,255,0.14)"}`, background: fonteFiltro !== "Todas" ? "rgba(0,224,168,0.14)" : "rgba(255,255,255,0.06)", color: fonteFiltro !== "Todas" ? "#4FE8BE" : "#D5DBF5" }}>
            {fonteFiltro === "Todas" ? "Todas as fontes" : fonteFiltro}<ChevronDown size={14} />
          </button>
        )} />
      {showFontes && (
        <Modal title="Fonte de pagamento" onClose={() => setShowFontes(false)}>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {["Todas", ...FONTES_FORM, ...fontesLegadas].map(f => {
              const sel = fonteFiltro === f;
              return (
                <button key={f} type="button" className="tap" onClick={() => { setFonteFiltro(f); setShowFontes(false); }}
                  style={{ display: "flex", alignItems: "center", justifyContent: "space-between", textAlign: "left", fontFamily: "inherit", fontSize: 14, fontWeight: sel ? 800 : 600, padding: "11px 14px", borderRadius: 12, cursor: "pointer",
                    border: `1px solid ${sel ? C.emerald : C.border}`, background: sel ? C.emeraldSoft : "transparent", color: sel ? C.emerald : C.ink }}>
                  {f === "Todas" ? "Todas as fontes" : f}{sel && <Check size={16} />}
                </button>
              );
            })}
          </div>
        </Modal>
      )}
      {showSeletorMesC && <SeletorMes mes={mesFiltro} onPick={setMesFiltro} onClose={() => setShowSeletorMesC(false)} />}

      {Object.keys(totals.porFonte).length > 0 && (
        <div className={serie.temDados ? "grid2c" : undefined} style={{ marginBottom: 20 }}>
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 16, display: "flex", flexDirection: "column" }}>
          <div style={{ fontWeight: 700, fontSize: 12.5, color: C.ink, marginBottom: 10, textTransform: "uppercase", letterSpacing: 0.4 }}>Por fonte de pagamento</div>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 18, flexWrap: "wrap", flex: 1 }}>
            <div style={{ width: 124, height: 124, flexShrink: 0, position: "relative" }}>
              <DonutTotal total={Object.values(totals.porFonte).reduce((t, v) => t + v, 0)} />
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={Object.entries(totals.porFonte).map(([name,value])=>({name,value}))} dataKey="value" nameKey="name" innerRadius={36} outerRadius={58} paddingAngle={Object.keys(totals.porFonte).length > 1 ? 2 : 0}>
                    {Object.keys(totals.porFonte).map((f) => <Cell key={f} fill={fonteColor(f)} />)}
                  </Pie>
                  <Tooltip separator=": " formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: "0 1 auto" }}>
              {Object.entries(totals.porFonte).map(([f,v],i) => (
                <div key={f} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12 }}>
                  <span style={{ width: 8, height: 8, borderRadius: 3, background: fonteColor(f), flexShrink: 0 }} />
                  <span style={{ color: C.ink, minWidth: 112 }}>{f}</span>
                  <span style={{ fontSize: 10.5, color: C.faint, minWidth: 28, textAlign: "right" }}>{pctDe(v, Object.values(totals.porFonte).map(x => ({ value: x })))}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted, fontSize: 11.5 }}>{fmtBRL(v)}</span>
                </div>
              ))}
              {meiosCustos.length > 0 && <div style={{ fontSize: 10.5, color: C.faint, marginTop: 2 }}>{meiosCustos.join(" · ")}</div>}
            </div>
          </div>
        </div>
        {serie.temDados && <SaidasMesCard serie={serie} mesLabel={monthLabel(mesFiltro)} />}
        </div>
      )}

      <FixosLed fixos={fixos} custos={custos} mes={mesFiltro} erro={fixosErro} onOpen={() => setShowFixos(true)} />

      {filtrados.length > 0 && totals.pendentes === 0 && totals.atrasados === 0 && (
        <div className="in" style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16, padding: "11px 14px", borderRadius: 14, background: C.emeraldSoft, color: C.emeraldDeep, fontSize: 13, fontWeight: 600 }}>
          <Check size={16} strokeWidth={2.6} /> {totals.pagos === 1 ? "O único custo lançado está pago." : `Todos os ${totals.pagos} custos lançados estão pagos.`}
        </div>
      )}

      {grupos.length === 0 ? (
        <div style={{
          border: `1px dashed ${C.border}`, borderRadius: 14, padding: "32px 16px",
          textAlign: "center", color: C.muted, fontSize: 13.5, background: C.card
        }}>Nenhum custo lançado neste mês/fonte.</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 22 }}>
          {grupos.map(g => {
            const totalGrupo = g.items.reduce((s,c) => s + Number(c.valorPago ?? c.valorPrevisto ?? 0), 0);
            const accent = STATUS_ACCENT[g.status] || C.border;
            return (
              <div key={g.status}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, padding: "0 2px" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                    <span style={{ width: 7, height: 7, borderRadius: "50%", background: accent }} />
                    <span style={{ fontWeight: 700, fontSize: 12.5, color: C.pageInk, textTransform: "uppercase", letterSpacing: 0.5 }}>
                      {GROUP_LABEL[g.status]}
                    </span>
                    <span style={{ fontSize: 11, fontWeight: 700, color: C.pageMuted, background: C.pageChip, borderRadius: 10, padding: "1px 7px" }}>
                      {g.items.length}
                    </span>
                  </div>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: C.pageMuted }}>{fmtBRL(totalGrupo)}</span>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {g.blocos.map((b, bi) => {
                    const linhas = b.itens.map((c, i) => renderLinhaCusto(c, bi + i, accent, b.dividido));
                    if (!b.dividido) return linhas;
                    return (
                      <div key={b.key} className="in" style={{
                        border: `1px solid ${C.lavender}55`, background: "rgba(154,165,209,0.05)", borderRadius: 18,
                        padding: 7, display: "flex", flexDirection: "column", gap: 6
                      }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "3px 9px 1px", fontSize: 10.5, fontWeight: 700, letterSpacing: 0.4, textTransform: "uppercase", color: C.lavender }}>
                          <span>{new Set(b.itens.map(x => x.fonte)).size === 1 ? `${b.itens.length} pagamentos · ${b.itens[0].fonte}` : `Dividido em ${b.itens.length} fontes`}</span>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted, letterSpacing: 0 }}>{fmtBRL(b.total)}</span>
                        </div>
                        {linhas}
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      )}


      {showFixos && <FixosModal fixos={fixos} custos={custos} mes={mesFiltro} onLancar={(fx, pago) => { setShowFixos(false); lancarFixo(fx, pago); }} onAdd={onAddFixo} onUpdate={onUpdateFixo} onRemove={onRemoveFixo} onClose={() => setShowFixos(false)} />}

      {showForm && (
        <Modal
          title={editingId ? "Editar custo" : "Novo custo"}
          headerExtra={editingId && idxEdicao >= 0 ? (
            <div style={{ display: "flex", alignItems: "center", gap: 6, marginLeft: "auto", marginRight: 10 }}>
              <span style={{ fontSize: 11.5, color: C.faint, fontWeight: 600 }}>{idxEdicao + 1} de {ordemLista.length}</span>
              {[[-1, "Custo anterior", "‹"], [1, "Próximo custo", "›"]].map(([d, rot, ch]) => (
                <button key={d} type="button" aria-label={rot} title={rot} onClick={() => navegarCusto(d)} disabled={!ordemLista[idxEdicao + d]}
                  style={{ border: `1px solid ${C.border}`, background: C.card, color: C.ink, borderRadius: 10, width: 34, height: 32, fontSize: 18, fontWeight: 700, lineHeight: 1, cursor: ordemLista[idxEdicao + d] ? "pointer" : "default", opacity: ordemLista[idxEdicao + d] ? 1 : 0.35 }}>{ch}</button>
              ))}
            </div>
          ) : null}
          onClose={() => { setShowForm(false); setEditingId(null); }}
          footer={
            <button type="submit" form="custo-form" disabled={saving} style={{ ...addBtnStyle, width: "100%", justifyContent: "center", opacity: saving ? 0.7 : 1, cursor: saving ? "default" : "pointer" }}>
              <Check size={16}/> {saving ? "Salvando..." : (editingId ? "Salvar alterações" : "Salvar custo")}
            </button>
          }
        >
          <form id="custo-form" onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, padding: "4px 0 24px" }}>
            <Field label="Descrição / Categoria" span={2}><TextInput required value={form.descricao} onChange={e=>setForm({...form,descricao:e.target.value})} /></Field>
            <Field label="Vencimento"><TextInput type="date" required value={form.vencimento} onChange={e=>setForm({...form,vencimento:e.target.value})} /></Field>
            <Field label="Data Pagamento"><TextInput type="date" value={form.dataPagamento} onChange={e=>setForm({...form,dataPagamento:e.target.value})} /></Field>
            <Field label="Valor Previsto (R$)"><TextInput type="text" inputMode="decimal" placeholder="0,00" required value={form.valorPrevisto} onChange={e=>setForm({...form,valorPrevisto:e.target.value})} /></Field>
            <Field label="Valor Pago (R$)"><TextInput type="text" inputMode="decimal" placeholder="0,00" value={form.valorPago} onChange={e=>setForm({...form,valorPago:e.target.value})} /></Field>
            <Field label="Fonte de Pagamento"><SelectInput options={opcoesFonte(form.fonte)} value={form.fonte} onChange={e=>setForm({...form,fonte:e.target.value})} /></Field>
            <Field label="Status"><SelectInput options={STATUS_CUSTO} value={form.status} onChange={e => {
              // Marking a custo as Pago fills, right here in the form (still editable before saving),
              // the payment date and the paid value when they are empty.
              const st = e.target.value;
              setForm(f => {
                const n = { ...f, status: st };
                if (st === "Pago") {
                  if (!f.dataPagamento) n.dataPagamento = todayISO();
                  if (parseValorBR(f.valorPago) === 0 && f.valorPrevisto !== "") n.valorPago = f.valorPrevisto;
                }
                return n;
              });
            }} /></Field>
            <Field label="Observações" span={2}><TextArea rows={10} style={{ height: 260 }} value={form.obs} onChange={e=>setForm({...form,obs:e.target.value})} /></Field>
          </form>
        </Modal>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Receitas Mensais                                                   */
/* ------------------------------------------------------------------ */

function ReceitasTab({ receitas, onAdd, onRemove, onUpdate, inicioContagem, jumpTarget, onConsumeJump }) {
  const [mesFiltro, setMesFiltro] = useState(currentMonth());
  const [showSeletorMesR, setShowSeletorMesR] = useState(false);
  useMesCorrenteAuto(setMesFiltro);
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [saving, setSaving] = useState(false);

  const filtradas = useMemo(() => {
    const list = receitas.filter(r => r.mes === mesFiltro);
    return [...list].sort((a, b) => {
      if (!a.dataRecebimento && !b.dataRecebimento) return 0;
      if (!a.dataRecebimento) return 1;
      if (!b.dataRecebimento) return -1;
      return a.dataRecebimento.localeCompare(b.dataRecebimento);
    });
  }, [receitas, mesFiltro]);

  // Depósitos are listed, but are not income: totals, charts and the next-to-receive led only see the rest.
  const rendas = useMemo(() => filtradas.filter(r => !ehDeposito(r)), [filtradas]);
  const depositosLista = useMemo(() => filtradas.filter(ehDeposito), [filtradas]);

  const totals = useMemo(() => {
    const totalPrevisto = rendas.reduce((s,r) => s + Number(r.valorPrevisto||0), 0);
    const totalRecebido = rendas.reduce((s,r) => s + Number(r.valorRecebido||0), 0);
    const recebidas = rendas.filter(r => r.status === "Recebido").length;
    const pendentes = rendas.filter(r => r.status !== "Recebido").length;
    const aReceber = rendas.filter(r => r.status !== "Recebido").reduce((s,r) => s + Number(r.valorPrevisto||0), 0);
    return { totalPrevisto, totalRecebido, diferenca: totalRecebido-totalPrevisto, aReceber, recebidas, pendentes };
  }, [rendas]);

  const comoCustos = useMemo(() => receitasComoCustos(receitas.filter(r => !ehDeposito(r))), [receitas]);
  const serie = useMemo(() => serieSaidas(comoCustos, mesFiltro, "Todas", inicioContagem), [comoCustos, mesFiltro, inicioContagem]);
  const proximo = useMemo(() => proximoVencimento(receitasComoCustos(rendas)), [rendas]);
  const porConta = useMemo(() => {
    const m = {};
    rendas.forEach(r => { const v = Number(r.valorRecebido || 0); if (v > 0) { const k = contaDaFonte(r.contaDestino || "—"); m[k] = (m[k] || 0) + v; } });
    return Object.entries(m).sort((a, b) => b[1] - a[1]);
  }, [rendas]);
  const meiosRec = useMemo(() => notasMeios(receitasComoCustos(rendas).filter(c => c.status === "Pago")), [rendas]);

  // Receitas habituais: títulos que entraram no mês anterior e ainda não existem neste mês (só no mês corrente). Só sugerem: nada é gravado.
  const esperadas = useMemo(() => {
    if (mesFiltro !== currentMonth()) return [];
    const tit = (r) => (r.descricao || "Outras").toUpperCase();
    const jaTem = new Set(rendas.map(tit));
    const ant = shiftMonth(mesFiltro, -1);
    const grupos = new Map();
    receitas.filter(r => r.mes === ant && !ehDeposito(r)).forEach(r => {
      const t = tit(r);
      if (jaTem.has(t)) return;
      const g = grupos.get(t) || { titulo: t, descricao: r.descricao, valor: 0, categoria: r.categoria || "", conta: r.contaDestino || "", dia: 0 };
      g.valor += Number(r.valorRecebido ?? r.valorPrevisto ?? 0);
      const dd = Number(String(r.dataRecebimento || "").slice(8, 10)) || 0;
      if (dd > g.dia) g.dia = dd;
      grupos.set(t, g);
    });
    return [...grupos.values()].filter(g => g.valor > 0).sort((a, b) => b.valor - a.valor);
  }, [receitas, rendas, mesFiltro]);
  const abrirEsperada = (g) => {
    destinoManual.current = true;
    setEditingId(null);
    const dia = Math.min(g.dia || 1, daysInMonthOf(mesFiltro));
    setForm({ ...blankReceita(mesFiltro), descricao: g.descricao, categoria: g.categoria, contaDestino: g.conta || FONTES[1], valorPrevisto: fmtValorInput(g.valor), dataRecebimento: `${mesFiltro}-${String(dia).padStart(2, "0")}` });
    setShowForm(true);
  };

  function blankReceita(mes) {
    return { mes, ano: parseInt(mes.split("-")[0],10), descricao: "", dataRecebimento: "",
      valorPrevisto: "", valorRecebido: "", categoria: "", contaDestino: FONTES[1], status: "Pendente", obs: "" };
  }
  const [form, setForm] = useState(blankReceita(mesFiltro));
  const destinoManual = useRef(false);

  const openNew = () => { destinoManual.current = false; setEditingId(null); setForm(blankReceita(mesFiltro)); setShowForm(true); };
  useEffect(() => {
    if (!jumpTarget) return;
    if (jumpTarget.novo) openNew();
    onConsumeJump && onConsumeJump();
  }, [jumpTarget]);
  const openEdit = (r) => {
    destinoManual.current = true;
    setEditingId(r.id);
    setForm({
      mes: r.mes, ano: r.ano, descricao: r.descricao, dataRecebimento: r.dataRecebimento || "",
      valorPrevisto: fmtValorInput(r.valorPrevisto), valorRecebido: fmtValorInput(r.valorRecebido),
      categoria: r.categoria || "", contaDestino: r.contaDestino, status: r.status, obs: r.obs || ""
    });
    setShowForm(true);
  };

  // Same two-block layout as Custos: what is still to come first, then what already came in.
  const gruposRec = useMemo(() => [
    { chave: "aberto", rotulo: "A receber", accent: C.amber, items: rendas.filter(r => r.status !== "Recebido") },
    { chave: "recebido", rotulo: "Recebidas", accent: C.emerald, items: rendas.filter(r => r.status === "Recebido") },
    { chave: "deposito", rotulo: "Depósitos de dinheiro (não contam como receita)", accent: C.lavender, items: depositosLista },
  ].filter(g => g.items.length > 0), [rendas, depositosLista]);

  const renderLinhaReceita = (r) => (
    <>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: 700, color: C.ink, fontSize: 15, textTransform: "uppercase" }}>{r.descricao}</div>
        <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2, display: "flex", alignItems: "center", flexWrap: "wrap", gap: "0 6px" }}>
          <span>
            {r.status === "Pendente" ? "Previsto para" : "Recebido"} {fmtDate(r.dataRecebimento)}{r.categoria ? ` · ${r.categoria}` : ""}{r.contaDestino ? " ·" : ""}
          </span>
          {r.contaDestino && (
            <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
              <span style={{ width: 8, height: 8, borderRadius: 3, background: fonteColorIn(r.contaDestino), flexShrink: 0 }} />
              {r.contaDestino}
            </span>
          )}
        </div>
      </div>
      <div style={{ textAlign: "right", minWidth: 92 }}>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 14.5, color: C.ink }}>
          {fmtBRL(r.valorRecebido ?? r.valorPrevisto)}
        </div>
        {r.valorRecebido != null && Math.abs(Number(r.valorRecebido) - Number(r.valorPrevisto || 0)) > 0.005 && (
          <div style={{ fontSize: 10.5, color: C.faint, marginTop: 1 }}>previsto {fmtBRL(r.valorPrevisto)}</div>
        )}
        <div style={{ marginTop: 3 }}><StatusPill status={r.status} /></div>
      </div>
    </>
  );

  const submit = async (e) => {
    e.preventDefault();
    if (saving) return;
    if (!form.descricao || !form.dataRecebimento || form.valorPrevisto === "") return;
    const valorPrevistoNum = parseValorBR(form.valorPrevisto);
    const valorRecebidoNum = form.valorRecebido === "" ? null : parseValorBR(form.valorRecebido);
    if (!isValorValido(valorPrevistoNum) || (valorRecebidoNum !== null && !isValorValido(valorRecebidoNum))) {
      alert("O valor precisa ser maior que zero.");
      return;
    }
    if (!(await confirmDateSanity(form.dataRecebimento, "data de recebimento"))) return;
    if (form.status === "Recebido" && valorRecebidoNum === null) {
      alert("Esta receita está como Recebida. Preencha o Valor Recebido.\n\nSe ainda não entrou, mude o Status para Pendente.");
      return;
    }
    // Data de recebimento no futuro: só pergunta quando a data ou o status mudou.
    const origRec = editingId ? receitas.find(r => r.id === editingId) : null;
    if (form.status === "Recebido" && form.dataRecebimento) {
      const mudouRec = !origRec || origRec.status !== "Recebido" || origRec.dataRecebimento !== form.dataRecebimento;
      if (mudouRec && form.dataRecebimento > todayISO() && !(await confirmar(`A data do recebimento (${fmtDate(form.dataRecebimento)}) ainda não chegou. Marcar como Recebida mesmo assim?`, { titulo: "Data no futuro", okLabel: "Está correta" }))) return;
    }
    // Confirmação da conta de destino: ao marcar Recebido (ou trocar a conta de uma receita já Recebida).
    if (form.status === "Recebido") {
      const orig = origRec;
      if (!orig || orig.status !== "Recebido" || orig.contaDestino !== form.contaDestino) {
        const difRc = valorRecebidoNum - valorPrevistoNum;
        const linhaPrevR = Math.abs(difRc) > 0.005 ? `\nPrevisto: ${fmtBRL(valorPrevistoNum)} (${difRc > 0 ? "+" : "−"}${fmtBRL(Math.abs(difRc))})` : "";
        const resumo = `${form.descricao}\nValor recebido: ${fmtBRL(valorRecebidoNum)}${linhaPrevR}\nConta: ${form.contaDestino || "—"}\nData do recebimento: ${fmtDate(form.dataRecebimento)}`;
        if (!(await confirmar(`${resumo}\n\nConfirma todos os dados?`, { titulo: "Marcar como Recebida?", okLabel: "Sim, confirmar" }))) return;
      }
    }
    const payload = { ...form, valorPrevisto: valorPrevistoNum, valorRecebido: valorRecebidoNum };
    setSaving(true);
    try {
      if (editingId) { await onUpdate(editingId, payload); } else { await onAdd(payload); }
      setForm(blankReceita(mesFiltro));
      setEditingId(null);
      setShowForm(false);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div>
      <TabHero titulo="Receitas" mesLabel={monthLabel(mesFiltro)} onLabelClick={() => setShowSeletorMesR(true)}
        destaque={totals.pendentes > 0
          ? { rotulo: "A receber", valor: fmtBRL(totals.aReceber), cor: "#FFC53D", sub: `Recebido ${fmtBRL(totals.totalRecebido)} de ${fmtBRL(totals.totalPrevisto)}` }
          : { rotulo: "Recebido no mês", valor: fmtBRL(totals.totalRecebido), cor: "#4FE8BE", sub: `de ${fmtBRL(totals.totalPrevisto)} previsto${totals.diferenca < -0.005 ? ` · ${fmtBRL(-totals.diferenca)} abaixo do previsto` : ""}` }}
        chips={[
          ...(totals.recebidas > 0 ? [{ t: `${totals.recebidas} ${totals.recebidas === 1 ? "recebida" : "recebidas"}`, cor: "#4FE8BE" }] : []),
          ...(totals.pendentes > 0 ? [{ t: `${totals.pendentes} ${totals.pendentes === 1 ? "pendente" : "pendentes"}`, cor: "#FFC53D" }] : [])
        ]}
        pct={totals.totalPrevisto > 0 ? (totals.totalRecebido / totals.totalPrevisto) * 100 : 0}
        pctLabel={totals.totalPrevisto > 0 ? `${Math.round((totals.totalRecebido / totals.totalPrevisto) * 100)}% do previsto recebido` : null}
        faixa={esperadas.length > 0 ? (
          <div style={{ display: "flex", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
            <span style={{ fontSize: 11, fontWeight: 700, color: "#8B93B5", textTransform: "uppercase", letterSpacing: 0.5 }}>Esperadas</span>
            {esperadas.map(g => (
              <button key={g.titulo} type="button" className="tap" onClick={() => abrirEsperada(g)} title="Abrir formulário preenchido"
                style={{ fontSize: 12, fontWeight: 700, padding: "4px 11px", borderRadius: 999, background: "rgba(255,197,61,0.1)", border: "1px solid rgba(255,197,61,0.4)", color: "#FFC53D", cursor: "pointer", fontFamily: "inherit" }}>
                {g.titulo.toLowerCase().replace(/(^|\s)(\S)/g, (m, sp, ch) => sp + ch.toUpperCase())} · {fmtBRL(g.valor)} ›
              </button>
            ))}
          </div>
        ) : null}
        acao={{ label: "Nova", onClick: openNew }}
        onPrev={() => setMesFiltro(m => shiftMonth(m, -1))} onNext={() => setMesFiltro(m => shiftMonth(m, 1))} />
      {showSeletorMesR && <SeletorMes mes={mesFiltro} onPick={setMesFiltro} onClose={() => setShowSeletorMesR(false)} />}

      {proximo && <ProximoVencLed info={proximo} onOpen={openEdit} entrada />}

      {(porConta.length > 0 || serie.temDados) && (
        <div className={porConta.length > 0 && serie.temDados ? "grid2c" : undefined} style={{ marginBottom: 20 }}>
          {porConta.length > 0 && (
            <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 16, display: "flex", flexDirection: "column" }}>
              <div style={{ fontWeight: 700, fontSize: 12.5, color: C.ink, marginBottom: 10, textTransform: "uppercase", letterSpacing: 0.4 }}>Por conta de destino</div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 18, flexWrap: "wrap", flex: 1 }}>
                <div style={{ width: 124, height: 124, flexShrink: 0, position: "relative" }}>
                  <DonutTotal total={porConta.reduce((t, [, v]) => t + v, 0)} />
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={porConta.map(([name, value]) => ({ name, value }))} dataKey="value" nameKey="name" innerRadius={36} outerRadius={58} paddingAngle={porConta.length > 1 ? 2 : 0}>
                        {porConta.map(([f]) => <Cell key={f} fill={fonteColorIn(f)} />)}
                      </Pie>
                      <Tooltip separator=": " formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: "0 1 auto" }}>
                  {porConta.map(([f, v]) => (
                    <div key={f} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12 }}>
                      <span style={{ width: 8, height: 8, borderRadius: 3, background: fonteColorIn(f), flexShrink: 0 }} />
                      <span style={{ color: C.ink, minWidth: 112 }}>{f}</span>
                      <span style={{ fontSize: 10.5, color: C.faint, minWidth: 28, textAlign: "right" }}>{pctDe(v, porConta.map(([, x]) => ({ value: x })))}</span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted, fontSize: 11.5 }}>{fmtBRL(v)}</span>
                    </div>
                  ))}
                  {meiosRec.length > 0 && <div style={{ fontSize: 10.5, color: C.faint, marginTop: 2 }}>{meiosRec.join(" · ")}</div>}
                </div>
              </div>
            </div>
          )}
          {serie.temDados && <SaidasMesCard serie={serie} mesLabel={monthLabel(mesFiltro)} entrada />}
        </div>
      )}

      {gruposRec.length === 0 ? (
        <CardList empty="Nenhuma receita lançada neste mês." rows={[]} onRemove={onRemove} onEdit={openEdit} render={renderLinhaReceita} />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 22 }}>
          {gruposRec.map(g => (
            <div key={g.chave}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, padding: "0 2px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                  <span style={{ width: 7, height: 7, borderRadius: "50%", background: g.accent }} />
                  <span style={{ fontWeight: 700, fontSize: 12.5, color: C.pageInk, textTransform: "uppercase", letterSpacing: 0.5 }}>{g.rotulo}</span>
                  <span style={{ fontSize: 11, fontWeight: 700, color: C.pageMuted, background: C.pageChip, borderRadius: 10, padding: "1px 7px" }}>{g.items.length}</span>
                </div>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: C.pageMuted }}>
                  {fmtBRL(g.items.reduce((s, r) => s + Number(r.valorRecebido ?? r.valorPrevisto ?? 0), 0))}
                </span>
              </div>
              <CardList empty="" rows={g.items} onRemove={onRemove} onEdit={openEdit} render={renderLinhaReceita} />
            </div>
          ))}
        </div>
      )}

      {showForm && (
        <Modal
          title={editingId ? "Editar receita" : "Nova receita"}
          onClose={() => { setShowForm(false); setEditingId(null); }}
          footer={
            <button type="submit" form="receita-form" disabled={saving} style={{ ...addBtnStyle, width: "100%", justifyContent: "center", opacity: saving ? 0.7 : 1, cursor: saving ? "default" : "pointer" }}>
              <Check size={16}/> {saving ? "Salvando..." : (editingId ? "Salvar alterações" : "Salvar receita")}
            </button>
          }
        >
          <form id="receita-form" onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, padding: "4px 0 24px" }}>
            <Field label="Descrição / Origem" span={2}><TextInput required value={form.descricao} onChange={e => {
              const v = e.target.value;
              const sug = destinoPadraoReceita(v);
              // New receita only: suggest the usual destination until the person picks one by hand.
              setForm(f => ({ ...f, descricao: v, contaDestino: (!editingId && !destinoManual.current && sug) ? sug : f.contaDestino }));
            }} /></Field>
            <Field label="Data Recebimento"><TextInput type="date" required value={form.dataRecebimento} onChange={e=>setForm({...form,dataRecebimento:e.target.value})} /></Field>
            <Field label="Categoria"><TextInput value={form.categoria} onChange={e=>setForm({...form,categoria:e.target.value})} /></Field>
            <Field label="Valor Previsto (R$)"><TextInput type="text" inputMode="decimal" placeholder="0,00" required value={form.valorPrevisto} onChange={e=>setForm({...form,valorPrevisto:e.target.value})} /></Field>
            <Field label="Valor Recebido (R$)"><TextInput type="text" inputMode="decimal" placeholder="0,00" value={form.valorRecebido} onChange={e=>setForm({...form,valorRecebido:e.target.value})} /></Field>
            <Field label="Conta de Destino"><SelectInput options={opcoesFonte(form.contaDestino)} value={form.contaDestino} onChange={e => { destinoManual.current = true; setForm({ ...form, contaDestino: e.target.value }); }} /></Field>
            <Field label="Status"><SelectInput options={STATUS_RECEITA} value={form.status} onChange={e => {
              // Marking a receita as Recebido fills the received value from the previsto when it is empty.
              const st = e.target.value;
              setForm(f => {
                const n = { ...f, status: st };
                if (st === "Recebido" && parseValorBR(f.valorRecebido) === 0 && f.valorPrevisto !== "") n.valorRecebido = f.valorPrevisto;
                return n;
              });
            }} /></Field>
            <Field label="Observações" span={2}><TextArea rows={10} style={{ height: 260 }} value={form.obs} onChange={e=>setForm({...form,obs:e.target.value})} /></Field>
          </form>
        </Modal>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Resumo Anual                                                       */
/* ------------------------------------------------------------------ */

function MonthCard({ nome, pago, saldo, status, pctLiquidado, temDados, atual, foraContagem, onClick }) {
  const quitado = status === "Em dia";
  return (
    <div className="tap" onClick={temDados ? onClick : undefined} style={{
      background: C.card, boxShadow: atual ? `0 0 0 1.5px ${C.emerald}, ${shadowSm}` : shadowSm, borderRadius: 16, padding: 16,
      opacity: !temDados ? 0.55 : foraContagem ? 0.5 : 1, cursor: temDados ? "pointer" : "default"
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 6, marginBottom: 8 }}>
        <span style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 0 }}>
          <span style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>{nome}</span>
        </span>
        {temDados && (quitado
          ? <Check size={15} color={C.emerald} strokeWidth={2.6} style={{ flexShrink: 0 }} />
          : status === "Pendências" ? <AlertCircle size={15} color={C.amber} style={{ flexShrink: 0 }} /> : null)}
      </div>
      {temDados ? (
        <>
          <ProgressBar pct={quitado ? 100 : pctLiquidado} height={6} />
          <div style={{ fontSize: 10.5, color: C.faint, marginTop: 5, whiteSpace: "nowrap" }}>{fmtBRL(pago)} pago</div>
          <div style={{ fontSize: 10.5, fontWeight: 700, marginTop: 1, whiteSpace: "nowrap", color: quitado ? C.emerald : C.faint }}>
            {quitado ? "quitado" : `${pctLiquidado.toFixed(0)}% do previsto`}
          </div>
          <div style={{ marginTop: 9 }}>
            <div style={{ fontSize: 10.5, color: C.muted, fontWeight: 600 }}>Saldo</div>
            <div style={{
              fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 13.5, whiteSpace: "nowrap",
              color: saldoCor(saldo)
            }}>{fmtBRL(saldo)}</div>
          </div>
          {atual
            ? <div style={{ fontSize: 10, fontWeight: 700, color: C.emerald, marginTop: 6 }}>mês atual</div>
            : foraContagem && <div style={{ fontSize: 10, color: C.faint, marginTop: 6 }}>fora da contagem</div>}
        </>
      ) : (
        <div style={{ fontSize: 11.5, color: C.faint, padding: "8px 0" }}>Sem lançamentos</div>
      )}
    </div>
  );
}

function ResumoTab({ custos, receitas, inicioContagem, onChangeInicio }) {
  const anos = useMemo(() => {
    const s = new Set([...custos, ...receitas].map(x => x.ano).filter(Boolean));
    s.add(new Date().getFullYear());
    return Array.from(s).sort();
  }, [custos, receitas]);
  const [ano, setAno] = useState(new Date().getFullYear());
  const [mesDetalhe, setMesDetalhe] = useState(null);

  const rows = useMemo(() => {
    return MESES.map((nome, idx) => {
      const mesKey = `${ano}-${String(idx+1).padStart(2,"0")}`;
      const cs = custos.filter(c => c.mes === mesKey);
      const rs = receitas.filter(r => r.mes === mesKey);
      const totalPrevistoCusto = cs.reduce((s,c)=>s+Number(c.valorPrevisto||0),0);
      const totalPagoCusto = cs.reduce((s,c)=>s+Number(c.valorPago||0),0);
      const totalRecebido = rs.reduce((s,r)=>s+Number(r.valorRecebido||0),0);
      const hasPendencia = cs.some(c => c.status !== "Pago");
      // Share of the planned amount already settled: custos marked Pago count
      // in full (their previsto), so a month with everything paid reads 100%
      // even when a few centavos differ between previsto and pago.
      const previstoQuitado = cs.filter(c => c.status === "Pago").reduce((t, c) => t + Number(c.valorPrevisto || 0), 0);
      return {
        mes: nome, mesKey, mesAbr: MESES_ABR[idx], totalPrevistoCusto, totalPagoCusto,
        pctLiquidado: totalPrevistoCusto > 0 ? Math.min(100, (previstoQuitado / totalPrevistoCusto) * 100) : 0,
        totalRecebido, saldo: totalRecebido - totalPagoCusto,
        status: cs.length > 0 ? (hasPendencia ? "Pendências" : "Em dia") : "—",
        temDados: cs.length > 0 || rs.length > 0
      };
    });
  }, [custos, receitas, ano]);

  // Months before the "contar a partir de" month stay visible but don't count
  // toward the year's totals.
  const inicioIdx = inicioIdxDoAno(inicioContagem, ano);
  const anual = rows.reduce((acc, r, i) => i < inicioIdx ? acc : ({
    previsto: acc.previsto + r.totalPrevistoCusto, pago: acc.pago + r.totalPagoCusto, recebido: acc.recebido + r.totalRecebido,
  }), { previsto: 0, pago: 0, recebido: 0 });

  // The chart scale follows only the months that count; older months (outside
  // "contar a partir de") are clipped to the scale and drawn faded, so they
  // can't flatten the months that matter. The tooltip still shows the real value.
  const saldosContados = rows.filter((_, i) => i >= inicioIdx && rows[i].temDados).map(r => r.saldo);
  const topoC = Math.max(1, ...saldosContados, 0);
  const baseC = Math.min(0, ...saldosContados);
  const domMax = topoC * 1.35;
  const domMin = Math.min(baseC * 1.5, -topoC * 0.3);
  const chartData = rows.map((r, i) => ({
    mes: r.mesAbr, Saldo: r.saldo, contado: i >= inicioIdx, temDados: r.temDados,
    Plot: Math.max(domMin, Math.min(domMax, r.saldo))
  }));
  const temCortada = chartData.some(d => !d.contado && d.temDados && d.Plot !== d.Saldo);
  const vazios = rows.filter(r => !r.temDados && r.mesKey !== currentMonth());
  const nCards = rows.filter(r => r.temDados || r.mesKey === currentMonth()).length;

  return (
    <div>
      <TabHero titulo="Resumo anual" mesLabel={String(ano)}
        destaque={{ rotulo: inicioIdx > 0 ? `Saldo desde ${MESES_ABR[inicioIdx]}` : "Saldo do ano", valor: fmtBRL(anual.recebido - anual.pago), cor: anual.recebido - anual.pago < 0 ? "#FF8A76" : "#4FE8BE", sub: null }}
        chips={[{ t: `Recebido ${fmtBRL(anual.recebido)}`, cor: "#4FE8BE" }, { t: `Pago ${fmtBRL(anual.pago)}`, cor: "#FF8A76" }, { t: `Previsto ${fmtBRL(anual.previsto)}` }]}
        pct={anual.previsto > 0 ? (anual.pago / anual.previsto) * 100 : null}
        pctLabel={anual.previsto > 0 ? `${Math.round((anual.pago / anual.previsto) * 100)}% do previsto pago` : null}
        faixa={(
          <div style={{ display: "flex", gap: 4 }} aria-label="Saldo de cada mês: verde positivo, coral negativo, esmaecido fora da contagem">
            {rows.map((r, i) => {
              const contado = i >= inicioIdx && r.temDados;
              const cor = !contado ? "rgba(255,255,255,0.18)" : saldoCor(r.saldo) === C.coral ? "#FF8A76" : saldoCor(r.saldo) === C.emerald ? "#4FE8BE" : "#B8C2E8";
              const atual = r.mesKey === currentMonth();
              return (
                <div key={r.mesKey || i} style={{ flex: 1, minWidth: 0, textAlign: "center", padding: "5px 0 6px", borderRadius: 9, background: atual ? "rgba(0,224,168,0.14)" : "rgba(255,255,255,0.04)", border: atual ? "1px solid rgba(0,224,168,0.5)" : "1px solid transparent" }}>
                  <div style={{ fontSize: 9.5, fontWeight: 700, color: atual ? "#4FE8BE" : "#8B93B5", letterSpacing: 0.3 }}>{r.mesAbr}</div>
                  <div style={{ height: 10, display: "flex", alignItems: "center", justifyContent: "center", marginTop: 3 }}>
                    <span style={{ width: contado ? 14 : 8, height: contado ? 6 : 4, borderRadius: 3, background: cor, opacity: contado ? 1 : 0.7 }} />
                  </div>
                </div>
              );
            })}
          </div>
        )}
        onPrev={() => setAno(a => a - 1)} onNext={() => setAno(a => a + 1)}
        prevOff={!anos.includes(ano - 1)} nextOff={!anos.includes(ano + 1)} />

      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
        <span style={{ fontSize: 12.5, color: C.muted, fontWeight: 600 }}>Contar a partir de</span>
        <SelectInput options={MESES} value={MESES[inicioIdx]}
          onChange={e => { const idx = MESES.indexOf(e.target.value); onChangeInicio && onChangeInicio(`${ano}-${String(Math.max(0, idx) + 1).padStart(2, "0")}`); }}
          style={{ maxWidth: 150 }} />
        {inicioIdx > 0 && (
          <span style={{ fontSize: 11.5, color: C.pageFaint }}>Meses anteriores aparecem esmaecidos e ficam fora dos totais{temCortada ? "; o ≈ marca barras cortadas na escala" : ""}.</span>
        )}
      </div>

      <div style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 10 }}>Saldo mês a mês</div>
        <div style={{ width: "100%", height: 180 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={chartData} margin={{ top: 14, bottom: 4 }}>
              <XAxis dataKey="mes" tick={{ fontSize: 11, fill: C.faint }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis hide domain={[domMin, domMax]} />
              <ReferenceLine y={0} stroke={C.border} />
              <Tooltip separator=": " formatter={(v, n, item) => [fmtBRL(item.payload.Saldo), "Saldo"]} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12.5, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
              <Bar dataKey="Plot" radius={[4, 4, 4, 4]}>
                {chartData.map((d, i) => <Cell key={i} fill={saldoCor(d.Saldo)} fillOpacity={d.contado ? 1 : 0.3} />)}
                <LabelList dataKey="Plot" content={(pr) => {
                  const d = chartData[pr.index];
                  if (!d || !d.temDados) return null;
                  const topo = Math.min(pr.y, pr.y + pr.height), fundo = Math.max(pr.y, pr.y + pr.height);
                  const neg = d.Saldo < 0;
                  // Mês fora da contagem com barra cortada na escala: só uma marca "≈" dentro da barra (o valor real está no cartão e no tooltip).
                  if (!d.contado) return d.Plot !== d.Saldo ? <text x={pr.x + pr.width / 2} y={neg ? fundo - 6 : topo + 14} textAnchor="middle" fontSize={13} fontWeight={700} fill={C.muted}>≈</text> : null;
                  const v = Math.round(Math.abs(d.Saldo));
                  const txt = `${neg ? "−" : "+"}${v.toLocaleString("pt-BR")}`;
                  return <text x={pr.x + pr.width / 2} y={neg ? fundo + 12 : topo - 5} textAnchor="middle" fontSize={10.5} fontWeight={700} fill={saldoCor(d.Saldo)}>{txt}</text>;
                }} />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={{ display: "grid", gap: 12, gridTemplateColumns: nCards >= 5 && nCards <= 6 ? "repeat(auto-fill, minmax(max(140px, calc((100% - 24px) / 3 - 0.5px)), 1fr))" : "repeat(auto-fill, minmax(140px, 1fr))" }}>
        {rows.map((r, i) => r.temDados || r.mesKey === currentMonth() ? (
          <MonthCard key={r.mes} nome={r.mes} pago={r.totalPagoCusto} saldo={r.saldo} status={r.status}
            pctLiquidado={r.pctLiquidado} temDados={r.temDados} atual={r.mesKey === currentMonth()}
            foraContagem={i < inicioIdx} onClick={() => setMesDetalhe(r)} />
        ) : null)}
      </div>
      {vazios.length > 0 && (
        <div style={{ marginTop: 14, marginBottom: 20, fontSize: 12, color: C.pageFaint }}>
          Sem lançamentos: {vazios.map(r => r.mes).join(" · ")}
        </div>
      )}

      {mesDetalhe && (
        <MonthDetailModal
          resumo={mesDetalhe}
          custos={custos.filter(c => c.mes === mesDetalhe.mesKey)}
          receitas={receitas.filter(r => r.mes === mesDetalhe.mesKey)}
          onClose={() => setMesDetalhe(null)}
        />
      )}
    </div>
  );
}

function MonthDetailModal({ resumo, custos, receitas, onClose }) {
  const porFonte = useMemo(() => {
    const map = {};
    custos.forEach(c => {
      const v = Number(c.valorPago ?? c.valorPrevisto ?? 0);
      const k = contaDaFonte(c.fonte);
      map[k] = (map[k] || 0) + v;
    });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [custos]);

  const itens = useMemo(() => {
    const cs = custos.map(c => ({ id: c.id, tipo: "custo", desc: c.descricao, data: c.dataPagamento || c.vencimento, valor: Number(c.valorPago ?? c.valorPrevisto ?? 0), status: c.status }));
    const rs = receitas.map(r => ({ id: r.id, tipo: "receita", desc: r.descricao, data: r.dataRecebimento, valor: Number(r.valorRecebido ?? r.valorPrevisto ?? 0), status: r.status }));
    return [...cs, ...rs].sort((a, b) => (b.data || "").localeCompare(a.data || ""));
  }, [custos, receitas]);

  return (
    <Modal title={resumo.mes} onClose={onClose}>
      <StatRow stats={[
        { icon: <Wallet size={12}/>, label: "Previsto", value: fmtBRL(resumo.totalPrevistoCusto) },
        { icon: <Check size={12}/>, label: "Pago", value: fmtBRL(resumo.totalPagoCusto) },
        { icon: <PiggyBank size={12}/>, label: "Recebido", value: fmtBRL(resumo.totalRecebido) },
        { icon: resumo.saldo<0?<TrendingDown size={12}/>:<TrendingUp size={12}/>, label: "Saldo", value: fmtBRL(resumo.saldo), tone: resumo.saldo<0?"red":"green" },
      ]} />

      {porFonte.length > 0 && (
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 20 }}>
          <div style={{ width: 110, height: 110, flexShrink: 0 }}>
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie data={porFonte} dataKey="value" nameKey="name" innerRadius={30} outerRadius={52} paddingAngle={porFonte.length > 1 ? 2 : 0}>
                  {porFonte.map((f) => <Cell key={f.name} fill={fonteColor(f.name)} />)}
                </Pie>
                <Tooltip separator=": " formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 120 }}>
            {porFonte.map((f, i) => (
              <div key={f.name} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12 }}>
                <span style={{ width: 8, height: 8, borderRadius: 3, background: fonteColor(f.name), flexShrink: 0 }} />
                <span style={{ color: C.ink, flex: 1, whiteSpace: "nowrap" }}>{f.name}</span>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted, fontSize: 11.5 }}>{fmtBRL(f.value)}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      <div style={{ fontWeight: 700, fontSize: 13, color: C.ink, marginBottom: 10, textTransform: "uppercase", letterSpacing: 0.4 }}>Lançamentos do mês</div>
      {itens.length === 0 ? (
        <div style={{ color: C.muted, fontSize: 13, textAlign: "center", padding: "16px 0" }}>Nada lançado neste mês.</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
          {itens.map(it => (
            <div key={`${it.tipo}-${it.id}`} style={{
              display: "flex", alignItems: "center", gap: 10, background: C.page, borderRadius: 10, padding: "9px 11px"
            }}>
              {it.tipo === "custo" ? <Wallet size={14} color={C.faint} /> : <PiggyBank size={14} color={C.faint} />}
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.ink }}>{it.desc}</div>
                <div style={{ fontSize: 10.5, color: C.muted }}>{fmtDate(it.data)}</div>
              </div>
              <div style={{
                fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 700,
                color: it.tipo === "custo" ? C.ink : C.emerald
              }}>
                {it.tipo === "custo" ? "-" : "+"}{fmtBRL(it.valor)}
              </div>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

/* ------------------------------------------------------------------ */
/*  Demandas                                                            */
/* ------------------------------------------------------------------ */

function DemandasTab({ demandas, onAdd, onRemove, onUpdate, onStatusChange, jumpTarget, onConsumeJump }) {
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [saving, setSaving] = useState(false);

  const resumo = useMemo(() => ({
    pendente: demandas.filter(d=>d.status==="Pendente").length,
    andamento: demandas.filter(d=>d.status==="Em Andamento").length,
    concluida: demandas.filter(d=>d.status==="Concluída").length,
  }), [demandas]);

  function blank() { return { nome: "", contato: "", detalhes: "", prazo: "", prioridade: "Alta", status: "Pendente", obs: "" }; }
  const [form, setForm] = useState(blank());

  const openNew = () => { setEditingId(null); setForm(blank()); setShowForm(true); };
  const openEdit = (d) => {
    setEditingId(d.id);
    setForm({
      nome: d.nome, contato: d.contato || "", detalhes: d.detalhes || "",
      prazo: d.prazo || "", prioridade: d.prioridade, status: d.status, obs: d.obs || "", concluidaEm: d.concluidaEm || ""
    });
    setShowForm(true);
  };

  useEffect(() => {
    if (!jumpTarget) return;
    const item = demandas.find(d => d.id === jumpTarget.id);
    if (item) openEdit(item);
    if (jumpTarget.novo) openNew();
    if (jumpTarget.adiarParadas) abrirAdiar(true);
    onConsumeJump();
  }, [jumpTarget]);

  const submit = async (e) => {
    e.preventDefault();
    if (saving) return;
    if (!form.nome) return;
    if (form.prazo && !(await confirmDateSanity(form.prazo, "prazo"))) return;
    setSaving(true);
    try {
      if (editingId) {
        const antes = demandas.find(d => d.id === editingId);
        const extra = {};
        if (antes && antes.status !== "Concluída" && form.status === "Concluída") extra.concluidaEm = form.concluidaEm || todayISO();
        else if (antes && antes.status === "Concluída" && form.status !== "Concluída") extra.concluidaEm = "";
        await onUpdate(editingId, { ...form, ...extra });
      } else {
        await onAdd({ ...form, criadoEm: todayISO(), ...(form.status === "Concluída" ? { concluidaEm: form.concluidaEm || todayISO() } : { concluidaEm: "" }) });
      }
      setForm(blank());
      setEditingId(null);
      setShowForm(false);
    } finally {
      setSaving(false);
    }
  };

  const PRIORIDADE_COLOR = { "Alta": C.coral, "Média": C.amber, "Baixa": C.emeraldDeep };
  const porPrioridade = useMemo(() => {
    const map = { "Alta": 0, "Média": 0, "Baixa": 0 };
    demandas.forEach(d => { map[d.prioridade] = (map[d.prioridade] || 0) + 1; });
    return Object.entries(map).filter(([,v]) => v > 0);
  }, [demandas]);
  const totalDemandas = demandas.length;
  const listaAtrasadas = useMemo(() => demandas.filter(d => d.status !== "Concluída" && (diasAteVenc(d.prazo) ?? 0) < 0).sort((a, b) => String(a.prazo).localeCompare(String(b.prazo))), [demandas]);
  const atrasadasDem = listaAtrasadas.length;
  const maisUrgenteDem = useMemo(() => {
    const abertas = demandas.filter(d => d.status !== "Concluída" && d.prazo).sort((a, b) => String(a.prazo).localeCompare(String(b.prazo)));
    // Tarefas paradas há 30+ dias não são "a mais urgente": prefere a atrasada mais antiga que ainda é recente, ou a próxima a vencer.
    return abertas.find(d => (diasAteVenc(d.prazo) ?? 0) > -PARADA_DIAS) || abertas[0] || null;
  }, [demandas]);
  const reguaDias = useMemo(() => {
    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    const cont = {};
    demandas.filter(d => d.status !== "Concluída" && d.prazo).forEach(d => { cont[d.prazo] = (cont[d.prazo] || 0) + 1; });
    return Array.from({ length: 14 }, (_, i) => {
      const dt = new Date(hoje); dt.setDate(hoje.getDate() + i);
      const iso = localISO(dt);
      return { iso, dow: "DSTQQSS"[dt.getDay()], dia: dt.getDate(), n: cont[iso] || 0, hoje: i === 0 };
    });
  }, [demandas]);
  const [showAdiar, setShowAdiar] = useState(false);
  const [novoPrazo, setNovoPrazo] = useState("");
  const [adiarSel, setAdiarSel] = useState([]);
  const [adiando, setAdiando] = useState(false);
  const abrirAdiar = (soParadas = false) => { setNovoPrazo(localISO(new Date(Date.now() + 7 * 86400000))); setAdiarSel(listaAtrasadas.filter(d => !soParadas || (diasAteVenc(d.prazo) ?? 0) <= -PARADA_DIAS).map(d => d.id)); setShowAdiar(true); };
  const adiarSelecionadas = async () => {
    if (adiando || !novoPrazo || adiarSel.length === 0) return;
    if (novoPrazo < todayISO()) { alert("Escolha uma data de hoje em diante."); return; }
    if (!(await confirmar(`Mudar o prazo de ${adiarSel.length} ${adiarSel.length === 1 ? "demanda" : "demandas"} para ${fmtDate(novoPrazo)}?`, { titulo: "Adiar atrasadas", okLabel: "Adiar" }))) return;
    setAdiando(true);
    try { for (const id of adiarSel) await onUpdate(id, { prazo: novoPrazo }); setShowAdiar(false); } finally { setAdiando(false); }
  };
  const pctConcluido = totalDemandas > 0 ? (resumo.concluida / totalDemandas) * 100 : 0;
  const COLUMNS = [
    { status: "Pendente", label: "Pendente", count: resumo.pendente, next: "Em Andamento" },
    { status: "Em Andamento", label: "Em Andamento", count: resumo.andamento, next: "Concluída" },
    { status: "Concluída", label: "Concluída", count: resumo.concluida, next: null },
  ];

  return (
    <div>
      <TabHero titulo="Tarefas"
        destaque={{ rotulo: "Em aberto", valor: `${resumo.pendente + resumo.andamento} ${(resumo.pendente + resumo.andamento) === 1 ? "tarefa" : "tarefas"}`, cor: (resumo.pendente + resumo.andamento) > 0 ? "#F3F5FA" : "#4FE8BE", sub: `${resumo.pendente} pendentes · ${resumo.andamento} em andamento · ${resumo.concluida} concluídas` }}
        chips={[
          ...(atrasadasDem > 0 ? [{ t: `${atrasadasDem} ${atrasadasDem === 1 ? "atrasada" : "atrasadas"} · adiar ›`, cor: "#FF8A76", onClick: () => abrirAdiar(false) }] : []),
          ...(porPrioridade.length > 1 ? porPrioridade.map(([p, n]) => ({ t: `${p} ${n}`, cor: p === "Alta" ? "#FF8A76" : p === "Média" ? "#FFC53D" : "#4FE8BE" })) : [])
        ]}
        pct={totalDemandas > 0 ? pctConcluido : null}
        pctLabel={totalDemandas > 0 ? `${pctConcluido.toFixed(0)}% concluído` : null}
        faixa={<ReguaDias dias={reguaDias} />}
        rodapeTexto={maisUrgenteDem ? `Mais urgente: ${maisUrgenteDem.nome} · ${String(maisUrgenteDem.prazo).slice(0, 4) === String(new Date().getFullYear()) ? fmtDate(maisUrgenteDem.prazo).slice(0, 5) : fmtDate(maisUrgenteDem.prazo)}` : "Nenhum prazo em aberto"}
        rodapeCor={maisUrgenteDem && (diasAteVenc(maisUrgenteDem.prazo) ?? 0) < 0 ? "#FF8A76" : "#B8C2E8"}
        acao={{ label: "Nova", onClick: openNew }} />

      <div className="kanban" style={{ display: "grid", gridTemplateColumns: "1fr", gap: 16, "--kcols": COLUMNS.map(c => c.count > 0 ? "2fr" : "1fr").join(" ") }}>
        {COLUMNS.map(col => {
          const items = demandas.filter(d => d.status === col.status).sort((a, b) => {
            // Open columns: earliest prazo first (overdue on top), no prazo last. Done: latest prazo first.
            const pa = a.prazo || "", pb = b.prazo || "";
            if (col.status === "Concluída") return (b.concluidaEm || pb).localeCompare(a.concluidaEm || pa);
            if (!pa && !pb) return 0;
            if (!pa) return 1;
            if (!pb) return -1;
            if (pa !== pb) return pa.localeCompare(pb);
            const ord = { "Alta": 0, "Média": 1, "Baixa": 2 };
            return (ord[a.prioridade] ?? 1) - (ord[b.prioridade] ?? 1) || (a.nome || "").localeCompare(b.nome || "");
          });
          return (
            <div key={col.status}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, padding: "0 2px" }}>
                <span style={{ fontWeight: 700, fontSize: 13.5, color: C.pageInk, textTransform: "uppercase", letterSpacing: 0.4 }}>{col.label}</span>
                <span style={{
                  fontSize: 11, fontWeight: 700, color: C.pageMuted, background: C.pageChip,
                  borderRadius: 10, padding: "1px 8px"
                }}>{col.count}</span>
              </div>
              {items.length === 0 ? (
                <div style={{ border: `1px dashed ${C.border}`, borderRadius: 14, padding: "12px 14px", textAlign: "center", color: C.faint, fontSize: 12.5 }}>
                  Nada aqui.
                </div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {items.map(d => (
                    <div key={d.id} className="tap in" onClick={() => openEdit(d)} style={{
                      background: C.card, boxShadow: shadowSm, borderRadius: 14, padding: "13px 14px",
                      position: "relative", cursor: "pointer"
                    }}>
                      <div style={{ position: "absolute", left: 0, top: 10, bottom: 10, width: 3, borderRadius: 3, background: d.status === "Concluída" ? C.emeraldDeep : (d.prazo && d.prazo < todayISO() ? C.coral : (d.status === "Em Andamento" ? C.lavender : C.amber)) }} />
                      <div style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                        <div style={{ fontWeight: 700, color: C.ink, fontSize: 14.5, marginBottom: 3, flex: 1, minWidth: 0 }}>{d.nome}</div>
                        <button onClick={async (e) => { e.stopPropagation(); if (await confirmar("Remover esta demanda?", { perigo: true, okLabel: "Remover" })) onRemove(d.id); }} title="Remover" style={{
                          border: "none", background: "transparent", color: C.faint, cursor: "pointer",
                          padding: 2, display: "flex", alignItems: "center", flexShrink: 0, marginRight: 24
                        }}>
                          <Trash2 size={15} />
                        </button>
                      </div>
                      <div style={{ fontSize: 11.5, color: C.muted, marginBottom: 3 }}>
                        Prazo {fmtDate(d.prazo)}{porPrioridade.length > 1 ? ` · ${d.prioridade}` : ""}{d.contato ? <> · {renderTel(d.contato)}</> : null}{d.detalhes ? <> · {renderTel(d.detalhes)}</> : null}
                      </div>
                      {d.status !== "Concluída" && (diasAteVenc(d.prazo) ?? 0) < 0 && (
                        <div style={{ fontSize: 11, fontWeight: 700, color: C.coral, marginBottom: 3 }}>{fmtAtraso(-diasAteVenc(d.prazo))}</div>
                      )}
                      {(d.criadoEm || (d.status === "Concluída" && d.concluidaEm)) && (
                        <div style={{ fontSize: 10.5, color: C.faint, marginBottom: col.next ? 10 : 0 }}>
                          {d.status === "Concluída" && d.concluidaEm && <div style={{ color: C.emerald, fontWeight: 700, fontSize: 11, marginBottom: 2 }}>Concluída em {fmtDate(d.concluidaEm)}</div>}
                          {d.criadoEm && <div>Criada em {fmtDate(d.criadoEm)}</div>}
                        </div>
                      )}
                      {col.next && <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                        {col.next && (
                          <button onClick={(e) => { e.stopPropagation(); onStatusChange(d.id, col.next); }} style={{
                            display: "flex", alignItems: "center", gap: 5, border: "none", background: C.border,
                            color: C.ink, borderRadius: 9, padding: "6px 10px", fontSize: 11.5, fontWeight: 700, cursor: "pointer"
                          }}>
                            {col.next} <ArrowRight size={12} />
                          </button>
                        )}
                        {(diasAteVenc(d.prazo) ?? 0) < 0 && (
                          <button onClick={async (e) => {
                            e.stopPropagation();
                            const novo = localISO(new Date(Date.now() + 7 * 86400000));
                            if (await confirmar(`Adiar o prazo de "${d.nome}" para ${fmtDate(novo)}?`, { okLabel: "Adiar" })) onUpdate(d.id, { prazo: novo });
                          }} title="Adiar prazo em 7 dias" style={{
                            display: "flex", alignItems: "center", gap: 5, border: `1px solid ${C.border}`, background: "transparent",
                            color: C.muted, borderRadius: 9, padding: "5px 10px", fontSize: 11.5, fontWeight: 700, cursor: "pointer"
                          }}>
                            Adiar 7 dias
                          </button>
                        )}
                      </div>}
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {showAdiar && (
        <Modal title="Adiar atrasadas" onClose={() => setShowAdiar(false)}
          footer={
            <button type="button" onClick={adiarSelecionadas} disabled={adiando || adiarSel.length === 0} style={{ ...addBtnStyle, width: "100%", justifyContent: "center", opacity: (adiando || adiarSel.length === 0) ? 0.6 : 1 }}>
              <Check size={16} /> {adiando ? "Adiando..." : `Adiar ${adiarSel.length} ${adiarSel.length === 1 ? "demanda" : "demandas"}`}
            </button>
          }>
          <div style={{ paddingBottom: 20 }}>
            <div style={{ fontSize: 13, color: C.muted, marginBottom: 14 }}>Marque as que quer adiar e escolha o novo prazo. As desmarcadas não mudam.</div>
            <Field label="Novo prazo"><TextInput type="date" value={novoPrazo} onChange={e => setNovoPrazo(e.target.value)} /></Field>
            <div style={{ marginTop: 14 }}>
              {listaAtrasadas.map(d => {
                const marcado = adiarSel.includes(d.id);
                return (
                  <label key={d.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 0", borderTop: `1px solid ${C.border}`, cursor: "pointer" }}>
                    <input type="checkbox" checked={marcado} onChange={() => setAdiarSel(sel => marcado ? sel.filter(x => x !== d.id) : [...sel, d.id])} style={{ width: 18, height: 18, accentColor: C.emerald }} />
                    <span style={{ flex: 1, minWidth: 0, fontSize: 14, fontWeight: 600, color: C.ink }}>{d.nome}</span>
                    <span style={{ fontSize: 11.5, color: C.coral, fontWeight: 700, flexShrink: 0 }}>{fmtDate(d.prazo)}</span>
                  </label>
                );
              })}
            </div>
          </div>
        </Modal>
      )}

      {showForm && (
        <Modal
          title={editingId ? "Editar demanda" : "Nova demanda"}
          onClose={() => { setShowForm(false); setEditingId(null); }}
          footer={
            <button type="submit" form="demanda-form" disabled={saving} style={{ ...addBtnStyle, width: "100%", justifyContent: "center", opacity: saving ? 0.7 : 1, cursor: saving ? "default" : "pointer" }}>
              <Check size={16}/> {saving ? "Salvando..." : (editingId ? "Salvar alterações" : "Salvar demanda")}
            </button>
          }
        >
          <form id="demanda-form" onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, padding: "4px 0 24px" }}>
            <Field label="Nome da Demanda" span={2}><TextInput required value={form.nome} onChange={e=>setForm({...form,nome:e.target.value})} /></Field>
            <Field label="Contato"><TextInput value={form.contato} onChange={e=>setForm({...form,contato:e.target.value})} /></Field>
            <Field label="Prazo"><TextInput type="date" value={form.prazo} onChange={e=>setForm({...form,prazo:e.target.value})} /></Field>
            <Field label="Prioridade"><SelectInput options={PRIORIDADES} value={form.prioridade} onChange={e=>setForm({...form,prioridade:e.target.value})} /></Field>
            <Field label="Status"><SelectInput options={STATUS_DEMANDA} value={form.status} onChange={e=>setForm({...form,status:e.target.value})} /></Field>
            {form.status === "Concluída" && <Field label="Concluída em"><TextInput type="date" value={form.concluidaEm || ""} onChange={e=>setForm({...form,concluidaEm:e.target.value})} /></Field>}
            <Field label="Detalhes" span={2}><TextInput value={form.detalhes} onChange={e=>setForm({...form,detalhes:e.target.value})} /></Field>
            <Field label="Observações" span={2}><TextArea rows={10} style={{ height: 260 }} value={form.obs} onChange={e=>setForm({...form,obs:e.target.value})} /></Field>
          </form>
        </Modal>
      )}
    </div>
  );
}

const addBtnStyle = {
  display: "flex", alignItems: "center", gap: 6, background: C.emerald, color: "#fff",
  border: "none", borderRadius: 14, padding: "12px 18px", fontSize: 14, fontWeight: 700,
  cursor: "pointer", boxShadow: "0 6px 16px rgba(0,200,150,0.35)"
};
const monthNavBtnStyle = {
  border: `1px solid ${C.border}`, background: C.card, color: C.ink, borderRadius: 10,
  width: 40, height: 40, fontSize: 16, fontWeight: 700, cursor: "pointer",
  display: "flex", alignItems: "center", justifyContent: "center"
};
