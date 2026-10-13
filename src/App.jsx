import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Plus, Trash2, X, Wallet, TrendingUp, TrendingDown, ClipboardList,
  Calendar, Check, AlertCircle, Clock, Filter, ChevronDown, PiggyBank,
  Home as HomeIcon, ArrowRight, Download, Pencil, Eye, EyeOff
} from "lucide-react";
import * as XLSX from "xlsx";
import { auth, db } from "./firebase.js";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { collection, doc, addDoc, updateDoc, deleteDoc, onSnapshot, writeBatch } from "firebase/firestore";
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, ReferenceLine, AreaChart, Area } from "recharts";

/* ------------------------------------------------------------------ */
/*  Firestore collections — one document per record. Never write an   */
/*  auto-generated "empty" document; only explicit user actions       */
/*  (add / edit / delete / import) touch the database.                */
/* ------------------------------------------------------------------ */
const custosCol = collection(db, "custos");
const receitasCol = collection(db, "receitas");
const demandasCol = collection(db, "demandas");

const MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const FONTES = ["Conta Salário","Conta Corrente","Dinheiro","PIX","Transferência","Débito Automático","BB"];
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
  { re: /\bbb\b/i, fonte: "BB" },
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
  if (!dataStr) dataStr = new Date().toISOString().slice(0, 10);

  // "salário" alone suggests income ("recebi meu salário"), but "conta
  // salário" is just the fonte name — strip recognized fonte phrases first so
  // mentioning that account doesn't misclassify a custo as a receita.
  let lowerSemFonte = lower;
  for (const { re } of FONTE_KEYWORDS) lowerSemFonte = lowerSemFonte.replace(re, " ");
  const isReceita = /receb|sal[áa]rio|entrada|venda|renda|dep[óo]sito/i.test(lowerSemFonte);
  const jaRecebido = /\breceb/i.test(lowerSemFonte);
  const jaPago = /\bpagu?ei\b|\bpago\b/i.test(lowerSemFonte);
  const status = isReceita ? (jaRecebido ? "Recebido" : "Pendente") : (jaPago ? "Pago" : "Pendente");
  const tipo = isReceita ? "receita" : "custo";

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
      .replace(/\bno dia\b/gi, " ").replace(/\bpagamento de\b/gi, " ")
      .replace(/\brecebi\b/gi, " ").replace(/\bpaguei\b/gi, " ").replace(/\bem\b/gi, " ")
      .replace(/[,.]/g, " ")
      .replace(/\b(de|da|do|e)\b/gi, " ")
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
const todayISO = () => new Date().toISOString().slice(0,10);
const currentMonth = () => new Date().toISOString().slice(0,7);
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
const confirmDateSanity = (dateStr, label = "data") => {
  if (!isDateImplausible(dateStr)) return true;
  return window.confirm(
    `A ${label} informada (${fmtDate(dateStr)}) parece muito distante da data de hoje. Confirma que está correta?`
  );
};

function exportToExcel(data) {
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

  XLSX.writeFile(wb, `Controle_Financeiro_${todayISO()}.xlsx`);
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
};
const shadowSm = "0 1px 3px rgba(0,0,0,0.35), 0 0 0 1px rgba(255,255,255,0.04)";
const shadowMd = "0 10px 30px rgba(0,0,0,0.45), 0 0 0 1px rgba(255,255,255,0.05)";
// Closed tonal system: every chart color is a shade of emerald/amber/coral/ink —
// no unrelated hues, so multi-slice charts still read as "one app", not a mix.
// Saídas (custos): tons quentes primeiro; Entradas (receitas): tons frios/verdes primeiro.
// Cada paleta tem 8 cores bem distintas entre si, para nenhuma fatia se confundir com outra.
const PALETTE_OUT = ["#FF5A68", "#FF8A3D", "#FFC53D", "#F472B6", "#A78BFA", "#38BDF8", "#2DD4BF", "#9AA5D1"];
const PALETTE_IN = ["#00E0A8", "#22D3EE", "#4FA3FF", "#B5E655", "#818CF8", "#FDE047", "#F0ABFC", "#86EFAC"];
const CHART_PALETTE = PALETTE_OUT;
// Cor fixa por fonte: "Conta Salário" tem sempre a mesma cor em qualquer gráfico ou mês.
const fonteColor = (name) => PALETTE_OUT[Math.max(0, FONTES.indexOf(name)) % PALETTE_OUT.length];
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
function RichTooltip({ active, payload, label, total, colorMap }) {
  if (!active || !payload || !payload.length) return null;
  const sum = total ?? payload.reduce((s, p) => s + (Number(p.value) || 0), 0);
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
          {sum > 0 && <span style={{ color: C.faint, fontSize: 10.5, minWidth: 30, textAlign: "right" }}>{((Number(p.value) / sum) * 100).toFixed(0)}%</span>}
        </div>
      ))}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 5, fontSize: 13 }}>
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

function TextInput(props) {
  return <input {...props} style={{ ...inputStyle, ...(props.style||{}) }} />;
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

function Modal({ title, onClose, children, footer }) {
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
function QuickAddFab({ onAddCusto, onAddReceita, demandas = [], onResolveDemanda }) {
  const [open, setOpen] = useState(false);
  const [step, setStep] = useState("input");
  const [text, setText] = useState("");
  const [draft, setDraft] = useState(null);
  const [saving, setSaving] = useState(false);

  const reset = () => { setOpen(false); setStep("input"); setText(""); setDraft(null); setSaving(false); };

  // draft is always an array — 1 item normally, 2+ when the sentence named a
  // split payment across different fontes (see parseQuickEntry). Each custo
  // entry also gets a matchedDemanda suggestion when its descrição clearly
  // corresponds to an open demanda (see matchDemanda) — off by default,
  // the person ticks it to also mark that demanda Concluída on save.
  const analisar = () => {
    if (!text.trim()) return;
    const parsed = parseQuickEntry(text).map(d => {
      if (d.tipo !== "custo") return d;
      const matched = matchDemanda(d.descricao, demandas);
      return matched ? { ...d, matchedDemandaId: matched.id, matchedDemandaNome: matched.nome, resolverDemanda: true } : d;
    });
    if (parsed.every(p => !isValorValido(p.valor))) {
      const continuar = window.confirm(
        "Não encontrei nenhum valor no texto digitado. Quer continuar e preencher o valor manualmente na próxima tela?"
      );
      if (!continuar) return;
    }
    setDraft(parsed);
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
    if (draft.some(d => !d.data)) { alert("Preencha a data de todos os lançamentos."); return; }
    const valores = draft.map(d => parseValorBR(d.valor));
    if (valores.some(v => !isValorValido(v))) {
      alert("Cada lançamento precisa de um valor maior que zero.");
      return;
    }
    if (!confirmDateSanity(draft[0].data, draft[0].tipo === "custo" ? "data de vencimento" : "data de recebimento")) return;
    setSaving(true);
    try {
      for (let i = 0; i < draft.length; i++) {
        const d = draft[i];
        const valorNum = valores[i];
        const mes = d.data.slice(0, 7);
        const anoNum = parseInt(d.data.slice(0, 4), 10);
        const status = d.status || "Pendente";
        if (d.tipo === "custo") {
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
            categoria: d.categoria || "", contaDestino: d.fonte || FONTES[1],
            status, obs: ""
          });
        }
      }
      reset();
    } finally {
      setSaving(false);
    }
  };

  return (
    <>
      <button className="tap" onClick={() => setOpen(true)} title="Lançamento rápido" style={{
        position: "fixed", right: 18, bottom: "calc(92px + env(safe-area-inset-bottom))", zIndex: 42,
        width: 56, height: 56, borderRadius: 28, border: "none",
        background: C.emerald, color: "#0A0E1C", cursor: "pointer",
        display: "flex", alignItems: "center", justifyContent: "center",
        boxShadow: "0 10px 26px rgba(0,224,168,0.45)"
      }}>
        <Plus size={26} strokeWidth={2.6} />
      </button>

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
                    {["custo", "receita"].map(t => (
                      <button key={t} type="button" onClick={() => updateDraft(i, { tipo: t })} style={{
                        flex: 1, padding: "9px 0", borderRadius: 10, border: `1px solid ${d.tipo === t ? C.emerald : C.border}`,
                        background: d.tipo === t ? C.emeraldSoft : C.card, color: d.tipo === t ? C.emerald : C.muted,
                        fontWeight: 700, fontSize: 12.5, cursor: "pointer"
                      }}>{t === "custo" ? "Custo" : "Receita"}</button>
                    ))}
                  </div>
                </Field>
                <Field label="Valor (R$)"><TextInput type="text" inputMode="decimal" placeholder="0,00" value={d.valor} onChange={e => updateDraft(i, { valor: e.target.value })} /></Field>
                <Field label="Descrição"><TextInput style={{ gridColumn: "span 2" }} value={d.descricao} onChange={e => updateDraft(i, { descricao: e.target.value })} /></Field>
                <Field label={d.tipo === "custo" ? "Vencimento" : "Recebimento"}>
                  <TextInput type="date" value={d.data} onChange={e => updateDraft(i, { data: e.target.value })} />
                </Field>
                <Field label={d.tipo === "custo" ? "Fonte de Pagamento" : "Conta de Destino"}>
                  <SelectInput options={FONTES} value={d.fonte || (d.tipo === "custo" ? FONTES[0] : FONTES[1])} onChange={e => updateDraft(i, { fonte: e.target.value })} />
                </Field>
                <Field label="Status">
                  <SelectInput options={d.tipo === "custo" ? STATUS_CUSTO : STATUS_RECEITA} value={d.status || "Pendente"} onChange={e => updateDraft(i, { status: e.target.value })} />
                </Field>
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
      background: C.card, boxShadow: `0 -1px 0 ${C.border}, 0 -8px 24px rgba(0,0,0,0.35)`,
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
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 28, fontWeight: 700, color: "#F3F5FA" }}>Controle Financeiro</div>
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
  const [user, setUser] = useState(undefined); // undefined = checking, null = logged out
  const [authError, setAuthError] = useState("");
  const [authLoading, setAuthLoading] = useState(false);
  const [custos, setCustos] = useState([]);
  const [receitas, setReceitas] = useState([]);
  const [demandas, setDemandas] = useState([]);
  const [loadedFlags, setLoadedFlags] = useState({ custos: false, receitas: false, demandas: false });
  const [tab, setTab] = useState("dashboard");
  const [jumpTarget, setJumpTarget] = useState(null);
  const goToItem = (tab, id) => { setTab(tab); setJumpTarget({ tab, id }); };
  const [saveError, setSaveError] = useState(false);
  const fileInputRef = useRef(null);
  const [menuOpen, setMenuOpen] = useState(false);

  const loading = !(loadedFlags.custos && loadedFlags.receitas && loadedFlags.demandas);
  const data = useMemo(() => ({ custos, receitas, demandas }), [custos, receitas, demandas]);

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
    return () => { unsubC(); unsubR(); unsubD(); };
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
    try { await updateDoc(doc(db, "demandas", id), { status }); setSaveError(false); } catch (e) { setSaveError(true); }
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
      const hasExisting = custos.length || receitas.length || demandas.length;
      if (hasExisting) {
        const ok = window.confirm(
          `Isso vai substituir os dados atuais (${custos.length} custos, ${receitas.length} receitas, ${demandas.length} demandas) pelos do arquivo importado. Continuar?`
        );
        if (!ok) { e.target.value = ""; return; }
      }
      const batch = writeBatch(db);
      custos.forEach(c => batch.delete(doc(db, "custos", c.id)));
      receitas.forEach(r => batch.delete(doc(db, "receitas", r.id)));
      demandas.forEach(d => batch.delete(doc(db, "demandas", d.id)));
      parsed.custos.forEach(c => { const { id, ...rest } = c; batch.set(doc(custosCol), rest); });
      parsed.receitas.forEach(r => { const { id, ...rest } = r; batch.set(doc(receitasCol), rest); });
      parsed.demandas.forEach(d => { const { id, ...rest } = d; batch.set(doc(demandasCol), rest); });
      await batch.commit();
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

  const displayName = (user?.email || "").split("@")[0];
  const greeting = displayName ? displayName.charAt(0).toUpperCase() + displayName.slice(1) : "";

  return (
    <div style={{ minHeight: "100vh", background: C.page, fontFamily: "'Inter', sans-serif", paddingBottom: 92 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:wght@400;500;600;700&family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600;700&display=swap');
        * { box-sizing: border-box; }
        body { margin: 0; }
        @keyframes slideUp { from { transform: translateY(24px); opacity: 0 } to { transform: translateY(0); opacity: 1 } }
        @keyframes fadeIn { from { opacity: 0 } to { opacity: 1 } }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: #2A3150; border-radius: 4px; }
        button:focus-visible, select:focus-visible, input:focus-visible { outline: 2px solid ${C.emerald}; outline-offset: 1px; }
        .tap { transition: transform .15s cubic-bezier(.34,1.56,.64,1), box-shadow .15s ease, filter .15s ease; }
        @media (hover: hover) and (pointer: fine) {
          .tap:hover { filter: brightness(1.08); transform: translateY(-1px); }
        }
        .tap:active { transform: scale(0.96); }
        @keyframes fadeInUp { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        .in { animation: fadeInUp .4s cubic-bezier(.22,1,.36,1) backwards; }
        .grid2 { display: grid; grid-template-columns: 1fr; gap: 20px; }
        @media (min-width: 680px) { .grid2 { grid-template-columns: 1fr 1fr; } .grid2 > * { margin-bottom: 0 !important; } }
        @media (min-width: 1100px) { .grid2 { grid-template-columns: 1fr 1fr 1fr; } }
        @media (min-width: 760px) { .kanban { grid-template-columns: 1fr 1fr 1fr !important; } }
      `}</style>

      {/* Header — same elevated card surface as everything else, not floating text */}
      <div style={{ padding: "16px 20px 0" }}>
        <header style={{
          maxWidth: 1080, margin: "0 auto", background: C.card, boxShadow: shadowSm,
          borderRadius: 20, padding: "16px 18px"
        }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{
              width: 42, height: 42, borderRadius: 13, background: C.emeraldSoft,
              display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0
            }}>
              <Wallet size={20} color={C.emerald} strokeWidth={2.2} />
            </div>
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 12.5, color: C.muted, fontWeight: 600 }}>Olá, {greeting} 👋</span>
                <span style={{
                  fontSize: 10.5, fontWeight: 700, color: C.faint, background: C.border,
                  borderRadius: 8, padding: "1px 7px", textTransform: "capitalize"
                }}>
                  {new Date().toLocaleDateString("pt-BR", { day: "2-digit", month: "short" })}
                </span>
              </div>
              <div style={{ fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 22, color: C.ink, marginTop: 2 }}>
                Controle Financeiro
              </div>
            </div>
          </div>
          <div style={{ position: "relative" }}>
            <button className="tap" onClick={() => setMenuOpen(o => !o)} style={{
              width: 40, height: 40, borderRadius: 12, border: `1px solid ${C.border}`, background: C.card,
              boxShadow: shadowSm, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer"
            }}>
              <span style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                <span style={{ width: 4, height: 4, borderRadius: 2, background: C.ink }} />
                <span style={{ width: 4, height: 4, borderRadius: 2, background: C.ink }} />
                <span style={{ width: 4, height: 4, borderRadius: 2, background: C.ink }} />
              </span>
            </button>
            {menuOpen && (
              <>
                <div onClick={() => setMenuOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 45 }} />
                <div style={{
                  position: "absolute", right: 0, top: 46, background: C.card, borderRadius: 14,
                  boxShadow: shadowMd, padding: 6, zIndex: 46, width: 200, animation: "fadeIn .15s ease-out"
                }}>
                  <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileChange} style={{ display: "none" }} />
                  <button onClick={() => { setMenuOpen(false); handleImportClick(); }} style={menuItemStyle}>
                    <Download size={15} style={{ transform: "rotate(180deg)" }} /> Importar Excel
                  </button>
                  <button onClick={() => { setMenuOpen(false); exportToExcel(data); }} style={menuItemStyle}>
                    <Download size={15} /> Exportar Excel
                  </button>
                  <div style={{ height: 1, background: C.border, margin: "4px 6px" }} />
                  <button onClick={() => { setMenuOpen(false); handleLogout(); }} style={{ ...menuItemStyle, color: C.coral }}>
                    <X size={15} /> Sair
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      </header>
      </div>

      <main style={{ maxWidth: 1080, margin: "0 auto", padding: "20px 20px 0" }}>
        {tab === "dashboard" && <DashboardTab custos={data.custos} receitas={data.receitas} demandas={data.demandas} onNavigate={goToItem} />}
        {tab === "custos" && <CustosTab custos={data.custos} onAdd={addCusto} onRemove={removeCusto} onUpdate={updateCusto} jumpTarget={jumpTarget && jumpTarget.tab === "custos" ? jumpTarget : null} onConsumeJump={() => setJumpTarget(null)} />}
        {tab === "receitas" && <ReceitasTab receitas={data.receitas} onAdd={addReceita} onRemove={removeReceita} onUpdate={updateReceita} />}
        {tab === "resumo" && <ResumoTab custos={data.custos} receitas={data.receitas} />}
        {tab === "demandas" && <DemandasTab demandas={data.demandas} onAdd={addDemanda} onRemove={removeDemanda} onUpdate={updateDemanda} onStatusChange={updateDemandaStatus} jumpTarget={jumpTarget && jumpTarget.tab === "demandas" ? jumpTarget : null} onConsumeJump={() => setJumpTarget(null)} />}
      </main>

      <BottomNav tab={tab} setTab={setTab} />
      <QuickAddFab onAddCusto={addCusto} onAddReceita={addReceita} demandas={data.demandas} onResolveDemanda={(id) => updateDemandaStatus(id, "Concluída")} />

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

function BalanceHero({ saldo, trendAtual, trendAnterior, mesLabel, hidden, onToggle, saldosPorMes, porFonte, expanded, onToggleExpanded }) {
  const maxAbs = Math.max(1, ...saldosPorMes.map(s => Math.abs(s)));
  // Trend compares "recebido - pago até o mesmo dia do mês" on both sides (not
  // the full previous month — see the trend calc in DashboardTab for why) and,
  // within that, % only makes intuitive sense against a positive base: "466%"
  // over a negative previous saldo reads as nonsense, so that case falls back
  // to a plain R$ difference instead.
  const saldoDiff = trendAnterior !== null ? trendAtual - trendAnterior : null;
  const trendPct = (trendAnterior !== null && trendAnterior > 0) ? (saldoDiff / trendAnterior) * 100 : null;
  return (
    <div className="in" style={{
      background: `linear-gradient(150deg, #0B0F1F 0%, #101632 55%, #0F3D33 100%)`,
      borderRadius: 28, padding: "32px 26px", marginBottom: 24,
      boxShadow: "0 20px 48px rgba(11,15,31,0.35), 0 6px 16px rgba(11,15,31,0.2)",
      position: "relative", overflow: "hidden"
    }}>
      <div style={{
        position: "absolute", top: -70, right: -60, width: 220, height: 220, borderRadius: "50%",
        background: `radial-gradient(circle, ${C.emeraldGlow} 0%, transparent 70%)`
      }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", position: "relative" }}>
        <div className="tap" onClick={onToggleExpanded} style={{ cursor: "pointer", flex: 1, minWidth: 0 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
            <span style={{ fontSize: 12.5, color: C.lavender, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.6 }}>Saldo de {mesLabel}</span>
            {trendPct !== null ? (
              <span style={{ display: "flex", alignItems: "center", gap: 2, fontSize: 10.5, fontWeight: 700, color: trendPct >= 0 ? "#4FE8BE" : "#FF8A76" }}>
                {trendPct >= 0 ? <TrendingUp size={11} /> : <TrendingDown size={11} />}
                {Math.abs(trendPct).toFixed(0)}% vs mês ant.
              </span>
            ) : saldoDiff !== null && (
              <span style={{ display: "flex", alignItems: "center", gap: 2, fontSize: 10.5, fontWeight: 700, color: saldoDiff >= 0 ? "#4FE8BE" : "#FF8A76" }}>
                {saldoDiff >= 0 ? <TrendingUp size={11} /> : <TrendingDown size={11} />}
                {hidden ? "•••" : fmtBRL(Math.abs(saldoDiff))} vs mês ant.
              </span>
            )}
          </div>
          <div style={{
            fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 48, letterSpacing: -1.2,
            color: saldo < 0 ? "#FF8A76" : "#4FE8BE", fontVariantNumeric: "tabular-nums", lineHeight: 1.02
          }}>
            {hidden ? "R$ •••••" : fmtBRL(saldo)}
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
          <button className="tap" onClick={onToggleExpanded} aria-label="Ver saldo por fonte" style={{
            width: 38, height: 38, borderRadius: 12, border: "1px solid rgba(255,255,255,0.12)", background: "rgba(255,255,255,0.06)",
            color: C.lavender, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
            transform: expanded ? "rotate(45deg)" : "none", transition: "transform .2s ease"
          }}>
            <Plus size={16} />
          </button>
          <button className="tap" onClick={onToggle} style={{
            width: 38, height: 38, borderRadius: 12, border: "1px solid rgba(255,255,255,0.12)", background: "rgba(255,255,255,0.06)",
            color: C.lavender, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center"
          }}>
            {hidden ? <EyeOff size={16} /> : <Eye size={16} />}
          </button>
        </div>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 5, height: 56, marginTop: 26, position: "relative" }}>
        {saldosPorMes.map((s, i) => {
          const h = s === 0 ? 2 : Math.max(3, (Math.abs(s) / maxAbs) * 24);
          const positive = s >= 0;
          return (
            <div key={i} style={{ display: "flex", flexDirection: "column", alignItems: "center", flex: 1, gap: 3 }}>
              <div style={{ height: 24, display: "flex", alignItems: "flex-end" }}>
                <div style={{ width: "100%", minWidth: 4, borderRadius: "3px 3px 0 0", height: positive ? h : 2, background: positive ? "#4FE8BE" : "rgba(255,255,255,0.12)" }} />
              </div>
              <div style={{ height: 24, display: "flex", alignItems: "flex-start" }}>
                <div style={{ width: "100%", minWidth: 4, borderRadius: "0 0 3px 3px", height: positive ? 2 : h, background: positive ? "rgba(255,255,255,0.12)" : "#FF8A76" }} />
              </div>
              <span style={{ fontSize: 8.5, color: "#7C86AD", fontWeight: 600 }}>{MESES_ABR[i]}</span>
            </div>
          );
        })}
      </div>

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
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// Overdue or due-within-7-days custos/demandas, surfaced at the very top
// of the dashboard so nothing urgent needs to be hunted for.
function useAtencao(custos, demandas, dias = 7) {
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
      .map(c => ({ id: `c-${c.id}`, rawId: c.id, tab: "custos", label: c.descricao, data: c.vencimento, dateObj: toDateObj(c.vencimento) }))
      .filter(i => i.dateObj && i.dateObj <= limite);
    const itensDemanda = demandas
      .filter(d => d.status !== "Concluída" && d.prazo)
      .map(d => ({ id: `d-${d.id}`, rawId: d.id, tab: "demandas", label: d.nome, data: d.prazo, dateObj: toDateObj(d.prazo) }))
      .filter(i => i.dateObj && i.dateObj <= limite);
    return [...itensCusto, ...itensDemanda]
      .map(i => ({ ...i, atrasado: i.dateObj < hoje }))
      .sort((a, b) => a.dateObj - b.dateObj);
  }, [custos, demandas, dias]);
}

function AttentionCard({ itens, onNavigate, title = "Precisa de atenção", limit = 6 }) {
  if (itens.length === 0) return null;
  return (
    <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 18, padding: "16px 18px", marginBottom: 20 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <AlertCircle size={16} color={C.coral} />
        <span style={{ fontWeight: 700, fontSize: 13, color: C.ink, textTransform: "uppercase", letterSpacing: 0.4 }}>{title}</span>
        <span style={{ fontSize: 11, fontWeight: 700, color: C.coral, background: C.coralSoft, borderRadius: 10, padding: "1px 8px" }}>
          {itens.length}
        </span>
      </div>
      <div style={{ display: "flex", flexDirection: "column" }}>
        {itens.slice(0, limit).map((item, i) => (
          <div
            key={item.id}
            className="tap"
            onClick={() => onNavigate && onNavigate(item.tab, item.rawId)}
            style={{
              display: "flex", alignItems: "center", gap: 10, cursor: onNavigate ? "pointer" : "default",
              padding: "9px 2px", borderBottom: i < Math.min(itens.length, limit) - 1 ? `1px solid ${C.border}` : "none"
            }}
          >
            <Clock size={14} color={item.atrasado ? C.coral : C.amber} style={{ flexShrink: 0 }} />
            <span style={{ flex: 1, fontSize: 13, color: C.ink, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {item.label}
            </span>
            <span style={{ fontSize: 11.5, fontWeight: 700, color: item.atrasado ? C.coral : C.amber, flexShrink: 0 }}>
              {item.atrasado ? "Atrasado" : fmtDate(item.data)}
            </span>
          </div>
        ))}
      </div>
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

function DashboardTab({ custos, receitas, demandas, onNavigate }) {
  const [hidden, setHidden] = useState(false);
  const [saldoAberto, setSaldoAberto] = useState(false);
  const ano = new Date().getFullYear();
  const mesAtual = currentMonth();
  const mesAnterior = shiftMonth(mesAtual, -1);
  const atencao = useAtencao(custos, demandas);
  const proximos30 = useAtencao(custos, demandas, 30);
  const ritmo = useRitmoMes(custos, mesAtual);

  const receitasComparativo = useMemo(() => {
    return receitas.filter(r => r.mes === mesAtual)
      .map(r => ({ nome: r.descricao, Previsto: Number(r.valorPrevisto || 0), Recebido: Number(r.valorRecebido || 0) }))
      .sort((a, b) => b.Previsto - a.Previsto)
      .slice(0, 6);
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
  const saldoPorFonte = useMemo(() => {
    const map = {};
    const get = (f) => (map[f] || (map[f] = { recebido: 0, pago: 0, pendente: 0, itensPendentes: [] }));
    custos.filter(c => c.mes === mesAtual).forEach(c => {
      const f = get(c.fonte || "—");
      if (c.status === "Pago") {
        f.pago += Number(c.valorPago || 0);
      } else {
        const v = Number(c.valorPrevisto || 0);
        f.pendente += v;
        f.itensPendentes.push({ descricao: c.descricao, valor: v, vencimento: c.vencimento });
      }
    });
    receitas.filter(r => r.mes === mesAtual).forEach(r => {
      get(r.contaDestino || "—").recebido += Number(r.valorRecebido || 0);
    });
    return Object.entries(map)
      .map(([fonte, v]) => {
        const saldo = v.recebido - v.pago;
        const saldoProjetado = saldo - v.pendente;
        return { fonte, ...v, saldo, saldoProjetado, vaiNegativar: v.pendente > 0 && saldoProjetado < 0 };
      })
      .filter(f => f.recebido > 0 || f.pago > 0 || f.pendente > 0)
      .sort((a, b) => (a.vaiNegativar !== b.vaiNegativar ? (a.vaiNegativar ? -1 : 1) : b.saldo - a.saldo));
  }, [custos, receitas, mesAtual]);

  const porFonte = useMemo(() => {
    const map = {};
    custos.filter(c => c.mes === mesAtual).forEach(c => {
      const v = Number(c.valorPago ?? c.valorPrevisto ?? 0);
      map[c.fonte] = (map[c.fonte] || 0) + v;
    });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [custos, mesAtual]);

  const porCategoria = useMemo(() => {
    const map = {};
    receitas.filter(r => r.mes === mesAtual).forEach(r => {
      const v = Number(r.valorRecebido ?? r.valorPrevisto ?? 0);
      const key = r.descricao || "Outras";
      map[key] = (map[key] || 0) + v;
    });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [receitas, mesAtual]);

  const catColor = useMemo(() => {
    const titulos = [...new Set(receitas.filter(r => r.mes === mesAtual).map(r => r.descricao || "Outras"))].sort();
    return (name) => PALETTE_IN[Math.max(0, titulos.indexOf(name)) % PALETTE_IN.length];
  }, [receitas, mesAtual]);

  const tendencia = useMemo(() => {
    const out = [];
    const now = new Date();
    for (let i = 5; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const mesKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const pago = custos.filter(c => c.mes === mesKey).reduce((s, c) => s + Number(c.valorPago || 0), 0);
      const recebido = receitas.filter(r => r.mes === mesKey).reduce((s, r) => s + Number(r.valorRecebido || 0), 0);
      out.push({ mes: MESES_ABR[d.getMonth()], Pago: pago, Recebido: recebido });
    }
    return out;
  }, [custos, receitas]);

  const topCustos = useMemo(() => {
    return [...custos.filter(c => c.mes === mesAtual)]
      .sort((a, b) => Number(b.valorPago ?? b.valorPrevisto ?? 0) - Number(a.valorPago ?? a.valorPrevisto ?? 0))
      .slice(0, 5);
  }, [custos, mesAtual]);
  const topMax = Math.max(1, ...topCustos.map(c => Number(c.valorPago ?? c.valorPrevisto ?? 0)));

  const projecao = useMemo(() => {
    const cs = custos.filter(c => c.mes === mesAtual);
    const rs = receitas.filter(r => r.mes === mesAtual);
    const pago = cs.reduce((t, c) => t + Number(c.valorPago || 0), 0);
    const aPagar = cs.filter(c => c.status !== "Pago").reduce((t, c) => t + Number(c.valorPrevisto || 0), 0);
    const recebido = rs.reduce((t, r) => t + Number(r.valorRecebido || 0), 0);
    const aReceber = rs.filter(r => r.status !== "Recebido").reduce((t, r) => t + Math.max(0, Number(r.valorPrevisto || 0) - Number(r.valorRecebido || 0)), 0);
    return { pago, aPagar, recebido, aReceber, saldoProjetado: (recebido + aReceber) - (pago + aPagar) };
  }, [custos, receitas, mesAtual]);
  const projMax = Math.max(1, projecao.recebido + projecao.aReceber, projecao.pago + projecao.aPagar);

  const comparativo = useMemo(() => {
    const soma = (mes, fonte) => custos.filter(c => c.mes === mes && c.fonte === fonte).reduce((t, c) => t + Number(c.valorPago ?? c.valorPrevisto ?? 0), 0);
    return FONTES.map(f => ({ fonte: f, Atual: soma(mesAtual, f), Anterior: soma(mesAnterior, f) })).filter(r => r.Atual > 0 || r.Anterior > 0);
  }, [custos, mesAtual, mesAnterior]);
  const variacoes = comparativo
    .filter(r => r.Anterior > 0)
    .map(r => ({ fonte: r.fonte, pct: ((r.Atual - r.Anterior) / r.Anterior) * 100 }))
    .sort((a, b) => Math.abs(b.pct) - Math.abs(a.pct));

  const acumulado = useMemo(() => {
    let cum = 0;
    const mesIdx = new Date().getMonth();
    return MESES_ABR.map((m, i) => { cum += saldosPorMes[i]; return { mes: m, Acumulado: i <= mesIdx ? cum : null }; });
  }, [saldosPorMes]);
  const acumAtual = [...acumulado].reverse().find(d => d.Acumulado !== null)?.Acumulado ?? 0;
  const acumColor = acumAtual < 0 ? C.coral : C.emerald;

  const mesLabel = monthLabel(mesAtual);

  return (
    <div>
      <AttentionCard itens={atencao} onNavigate={onNavigate} />

      <BalanceHero
        saldo={kpis.saldo}
        trendAtual={saldoAteDiaAtual}
        trendAnterior={saldoAteDiaAnterior}
        mesLabel={mesLabel}
        hidden={hidden}
        onToggle={() => setHidden(h => !h)}
        saldosPorMes={saldosPorMes}
        porFonte={saldoPorFonte}
        expanded={saldoAberto}
        onToggleExpanded={() => setSaldoAberto(a => !a)}
      />

      <StatRow stats={[
        { icon: <Wallet size={12} />, label: "Previsto", value: hidden ? "•••" : fmtBRL(kpis.totalPrevisto) },
        {
          icon: <Check size={12} />, label: "Pago", value: hidden ? "•••" : fmtBRL(kpis.totalPago),
          trend: pagoTrendPct !== null ? { pct: pagoTrendPct, good: pagoTrendPct <= 0 } : null
        },
        {
          icon: <PiggyBank size={12} />, label: "Recebido", value: hidden ? "•••" : fmtBRL(kpis.totalRecebido),
          trend: recebidoTrendPct !== null ? { pct: recebidoTrendPct, good: recebidoTrendPct >= 0 } : null
        },
      ]} />

      <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 18, padding: "20px 22px", marginBottom: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 16 }}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink }}>Projeção de fechamento — {mesLabel}</div>
            <div style={{ fontSize: 11.5, color: C.faint, marginTop: 3 }}>Se tudo o que está previsto for pago e recebido</div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 24, fontWeight: 700, color: projecao.saldoProjetado < 0 ? C.coral : C.emerald }}>
              {hidden ? "R$ •••••" : fmtBRL(projecao.saldoProjetado)}
            </div>
            <div style={{ fontSize: 11, color: C.faint }}>saldo previsto no fim do mês</div>
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <ProjBar label="Entradas" done={projecao.recebido} pending={projecao.aReceber} color={C.emerald} max={projMax} hidden={hidden} />
          <ProjBar label="Saídas" done={projecao.pago} pending={projecao.aPagar} color={C.coral} max={projMax} hidden={hidden} />
        </div>
        <div style={{ display: "flex", gap: 16, marginTop: 12, fontSize: 11, color: C.faint, flexWrap: "wrap" }}>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 10, height: 10, borderRadius: 3, background: C.muted }} />já realizado</span>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 10, height: 10, borderRadius: 3, background: `repeating-linear-gradient(135deg, ${C.muted}88 0 3px, ${C.muted}33 3px 6px)` }} />ainda previsto</span>
        </div>
        <div style={{ marginTop: 18 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
            <span style={{ fontWeight: 700, fontSize: 12.5, color: C.ink }}>Ritmo do mês</span>
            {ritmo.disponivel && (
              <span style={{ fontSize: 12.5, fontWeight: 700, color: ritmo.pct > 150 ? C.coral : ritmo.pct > 100 ? C.amber : C.emerald }}>
                {ritmo.pct.toFixed(0)}% do ritmo normal
              </span>
            )}
          </div>
          {ritmo.disponivel ? (
            <>
              <div style={{ fontSize: 11, color: C.faint, marginBottom: 8 }}>
                Comparado ao que você normalmente já tinha pago até o dia {ritmo.diaRef}, nos últimos {ritmo.nMeses === 1 ? "mês" : `${ritmo.nMeses} meses`} — não à média do mês inteiro
              </div>
              <ProgressBar
                pct={ritmo.pct}
                color="linear-gradient(90deg, #00E0A8, #22D3EE)"
                overColor={ritmo.pct > 150 ? "linear-gradient(90deg,#FFB929,#FF5A68)" : "linear-gradient(90deg,#FFC53D,#FFB929)"}
              />
              <div style={{ fontSize: 11, color: C.faint, marginTop: 8 }}>
                Pago até agora: <strong style={{ color: C.ink }}>{hidden ? "•••" : fmtBRL(ritmo.pagoAteAgora)}</strong>
                {" · "}Normal até o dia {ritmo.diaRef}: <strong style={{ color: C.ink }}>{hidden ? "•••" : fmtBRL(ritmo.mediaAteODia)}</strong>
                {" · "}Média mensal: <strong style={{ color: C.ink }}>{hidden ? "•••" : fmtBRL(ritmo.mediaMensal)}</strong>
              </div>
            </>
          ) : (
            <div style={{ fontSize: 12, color: C.muted, padding: "6px 0" }}>
              Ainda não há meses anteriores com custos pagos suficientes para comparar o ritmo deste mês.
            </div>
          )}
        </div>
      </div>

      <div className="grid2" style={{ marginBottom: 20 }}>
        {/* Orçamento agora vive dentro do card de projeção */}

        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "40ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 10 }}>Custos por fonte — {mesLabel}</div>
          {porFonte.length === 0 ? (
            <div style={{ color: C.muted, fontSize: 13, padding: "16px 0", textAlign: "center" }}>Sem custos lançados neste mês.</div>
          ) : (
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <div style={{ width: 130, height: 130, flexShrink: 0 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie data={porFonte} dataKey="value" nameKey="name" innerRadius={36} outerRadius={60} paddingAngle={2}>
                      {porFonte.map((f) => <Cell key={f.name} fill={fonteColor(f.name)} />)}
                    </Pie>
                    <Tooltip content={<RichTooltip />} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 120 }}>
                {porFonte.map((f, i) => (
                  <div key={f.name} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>
                    <span style={{ width: 9, height: 9, borderRadius: 3, background: fonteColor(f.name), flexShrink: 0 }} />
                    <span style={{ color: C.ink, flex: 1 }}>{f.name}</span>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted, fontSize: 12 }}>{fmtBRL(f.value)}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "120ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 10 }}>Receitas por título — {mesLabel}</div>
          {porCategoria.length === 0 ? (
            <div style={{ color: C.muted, fontSize: 13, padding: "16px 0", textAlign: "center" }}>Sem receitas lançadas neste mês.</div>
          ) : (
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <div style={{ width: 130, height: 130, flexShrink: 0 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie data={porCategoria} dataKey="value" nameKey="name" innerRadius={36} outerRadius={60} paddingAngle={2}>
                      {porCategoria.map((f) => <Cell key={f.name} fill={catColor(f.name)} />)}
                    </Pie>
                    <Tooltip content={<RichTooltip />} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 120 }}>
                {porCategoria.map((f, i) => (
                  <div key={f.name} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>
                    <span style={{ width: 9, height: 9, borderRadius: 3, background: catColor(f.name), flexShrink: 0 }} />
                    <span style={{ color: C.ink, flex: 1 }}>{f.name}</span>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted, fontSize: 12 }}>{fmtBRL(f.value)}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "160ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 10 }}>Pago x Recebido — últimos 6 meses</div>
          <div style={{ width: "100%", height: 170 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={tendencia} barGap={4}>
                <defs>
                  <linearGradient id="gPago" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#FF7A85" /><stop offset="100%" stopColor="#E8434F" />
                  </linearGradient>
                  <linearGradient id="gRecebido" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#2CF0BD" /><stop offset="100%" stopColor="#00B98A" />
                  </linearGradient>
                </defs>
                <XAxis dataKey="mes" tick={{ fontSize: 11, fill: C.faint }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis hide />
                <Tooltip content={<RichTooltip colorMap={{ Pago: "#E8434F", Recebido: "#00B98A" }} />} />
                <Bar dataKey="Pago" fill="url(#gPago)" radius={[4, 4, 0, 0]} />
                <Bar dataKey="Recebido" fill="url(#gRecebido)" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "200ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 10 }}>Maiores custos — {mesLabel}</div>
          {topCustos.length === 0 ? (
            <div style={{ color: C.muted, fontSize: 13, padding: "16px 0", textAlign: "center" }}>Sem custos lançados neste mês.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {topCustos.map((c, i) => {
                const v = Number(c.valorPago ?? c.valorPrevisto ?? 0);
                const intensidade = v / topMax; // 0..1, share of the heaviest cost this month
                const tint = intensidade > 0.7 ? C.coralSoft : intensidade > 0.4 ? C.amberSoft : "transparent";
                const valueColor = intensidade > 0.7 ? C.coral : intensidade > 0.4 ? C.amber : C.muted;
                return (
                  <div key={c.id} style={{ background: tint, borderRadius: 8, padding: "5px 7px", margin: "-5px -7px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 4 }}>
                      <span style={{ color: C.ink, fontWeight: 600 }}>{c.descricao}</span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: valueColor, fontWeight: intensidade > 0.7 ? 700 : 400 }}>{fmtBRL(v)}</span>
                    </div>
                    <ProgressBar pct={(v / topMax) * 100} color={fonteColor(c.fonte)} height={6} />
                  </div>
                );
              })}
            </div>
          )}
        </div>
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "240ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 4 }}>Este mês x mês anterior — por fonte</div>
          <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 10 }}>{monthLabel(mesAnterior)} (cinza) x {mesLabel}</div>
          {comparativo.length === 0 ? (
            <div style={{ color: C.muted, fontSize: 13, padding: "16px 0", textAlign: "center" }}>Sem custos nos dois meses para comparar.</div>
          ) : (
            <>
              <div style={{ width: "100%", height: 200 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={comparativo} barGap={3} margin={{ bottom: 24 }}>
                    <XAxis dataKey="fonte" interval={0} angle={-35} textAnchor="end" height={50}
                      tick={{ fontSize: 9.5, fill: C.faint }} axisLine={{ stroke: C.border }} tickLine={false}
                      tickFormatter={(v) => v.replace("Conta ", "").replace("Débito Automático", "Déb. Auto.")} />
                    <YAxis hide />
                    <Tooltip content={<RichTooltip colorMap={{ Anterior: "#5C6485", Atual: "#E8434F" }} />} />
                    <Bar dataKey="Anterior" fill="#5C6485" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="Atual" fill="url(#gPago)" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
              {variacoes[0] && (
                <div style={{ fontSize: 12, color: C.muted, marginTop: 8 }}>
                  Maior variação: <strong style={{ color: C.ink }}>{variacoes[0].fonte}</strong>{" "}
                  <strong style={{ color: variacoes[0].pct > 0 ? C.coral : C.emerald }}>{variacoes[0].pct > 0 ? "+" : ""}{variacoes[0].pct.toFixed(0)}%</strong>
                </div>
              )}
            </>
          )}
        </div>

        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "280ms" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
            <span style={{ fontWeight: 700, fontSize: 13.5, color: C.ink }}>Saldo acumulado — {ano}</span>
            <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 700, color: acumColor }}>{hidden ? "•••" : fmtBRL(acumAtual)}</span>
          </div>
          <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 10 }}>Soma dos saldos mês a mês desde janeiro</div>
          <div style={{ width: "100%", height: 170 }}>
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={acumulado}>
                <defs>
                  <linearGradient id="gAcum" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={acumColor} stopOpacity={0.45} /><stop offset="100%" stopColor={acumColor} stopOpacity={0.02} />
                  </linearGradient>
                </defs>
                <XAxis dataKey="mes" tick={{ fontSize: 11, fill: C.faint }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis hide />
                <ReferenceLine y={0} stroke={C.border} />
                <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12.5, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
                <Area type="monotone" dataKey="Acumulado" stroke={acumColor} strokeWidth={2.5} fill="url(#gAcum)" dot={false} connectNulls={false} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "320ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 4 }}>Receitas previstas x recebidas — {mesLabel}</div>
          <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 10 }}>Cinza é o previsto, verde é o que já entrou</div>
          {receitasComparativo.length === 0 ? (
            <div style={{ color: C.muted, fontSize: 13, padding: "16px 0", textAlign: "center" }}>Sem receitas lançadas neste mês.</div>
          ) : (
            <div style={{ width: "100%", height: 220 }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={receitasComparativo} barGap={3} margin={{ bottom: 30 }}>
                  <XAxis dataKey="nome" interval={0} angle={-35} textAnchor="end" height={60}
                    tick={{ fontSize: 9, fill: C.faint }} axisLine={{ stroke: C.border }} tickLine={false}
                    tickFormatter={(v) => v.length > 8 ? v.slice(0, 8) + "…" : v} />
                  <YAxis hide />
                  <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12.5, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
                  <Bar dataKey="Previsto" fill="#5C6485" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Recebido" fill={C.emerald} radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>

        <AttentionCard itens={proximos30} onNavigate={onNavigate} title="Próximos vencimentos (30 dias)" limit={8} />
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
          <button onClick={(e) => { e.stopPropagation(); if (window.confirm("Remover este item?")) onRemove(row.id); }} title="Remover" style={{
            border: "none", background: "transparent", color: C.faint, cursor: "pointer",
            padding: 4, display: "flex", alignItems: "center", flexShrink: 0
          }}>
            <Trash2 size={16} />
          </button>
        </div>
      ))}
    </div>
  );
}

function CustosTab({ custos, onAdd, onRemove, onUpdate, jumpTarget, onConsumeJump }) {
  const [mesFiltro, setMesFiltro] = useState(currentMonth());
  const [fonteFiltro, setFonteFiltro] = useState("Todas");
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [saving, setSaving] = useState(false);

  const filtrados = useMemo(() => {
    const list = custos.filter(c => c.mes === mesFiltro && (fonteFiltro === "Todas" || c.fonte === fonteFiltro));
    return [...list].sort((a, b) => {
      if (!a.dataPagamento && !b.dataPagamento) return 0;
      if (!a.dataPagamento) return 1;
      if (!b.dataPagamento) return -1;
      return a.dataPagamento.localeCompare(b.dataPagamento);
    });
  }, [custos, mesFiltro, fonteFiltro]);

  const totals = useMemo(() => {
    const totalPrevisto = filtrados.reduce((s,c) => s + Number(c.valorPrevisto||0), 0);
    const totalPago = filtrados.reduce((s,c) => s + Number(c.valorPago||0), 0);
    const pagos = filtrados.filter(c => c.status === "Pago").length;
    const pendentes = filtrados.filter(c => c.status === "Pendente").length;
    const atrasados = filtrados.filter(c => c.status === "Atrasado").length;
    const porFonte = {};
    filtrados.forEach(c => { porFonte[c.fonte] = (porFonte[c.fonte]||0) + Number(c.valorPago||0); });
    return { totalPrevisto, totalPago, diferenca: totalPago-totalPrevisto, pagos, pendentes, atrasados, porFonte };
  }, [filtrados]);

  function blankCusto(mes) {
    return { mes, ano: parseInt(mes.split("-")[0],10), descricao: "", vencimento: "", dataPagamento: "",
      valorPrevisto: "", valorPago: "", fonte: FONTES[0], status: "Pendente", obs: "" };
  }
  const [form, setForm] = useState(blankCusto(mesFiltro));

  const openNew = () => { setEditingId(null); setForm(blankCusto(mesFiltro)); setShowForm(true); };
  const openEdit = (c) => {
    setEditingId(c.id);
    setForm({
      mes: c.mes, ano: c.ano, descricao: c.descricao, vencimento: c.vencimento || "",
      dataPagamento: c.dataPagamento || "", valorPrevisto: c.valorPrevisto ?? "",
      valorPago: c.valorPago ?? "", fonte: c.fonte, status: c.status, obs: c.obs || ""
    });
    setShowForm(true);
  };

  useEffect(() => {
    if (!jumpTarget) return;
    const item = custos.find(c => c.id === jumpTarget.id);
    if (item) { setMesFiltro(item.mes); openEdit(item); }
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
    if (!confirmDateSanity(form.vencimento, "data de vencimento")) return;
    const payload = { ...form, valorPrevisto: valorPrevistoNum, valorPago: valorPagoNum };
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
      .map(status => ({ status, items: filtrados.filter(c => c.status === status) }))
      .filter(g => g.items.length > 0);
  }, [filtrados]);

  const GROUP_LABEL = { "Atrasado": "Atrasados", "Pendente": "A pagar", "Pago": "Pagos" };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <h2 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 25, letterSpacing: -0.4, color: C.ink }}>Custos</h2>
        <button className="tap" onClick={openNew} style={addBtnStyle}><Plus size={16} /> Novo</button>
      </div>

      <FilterBar>
        <Field label="Mês"><TextInput type="month" value={mesFiltro} onChange={e => setMesFiltro(e.target.value)} /></Field>
        <div style={{ display: "flex", gap: 6 }}>
          <button type="button" onClick={() => setMesFiltro(m => shiftMonth(m, -1))} style={monthNavBtnStyle} aria-label="Mês anterior">‹</button>
          <button type="button" onClick={() => setMesFiltro(m => shiftMonth(m, 1))} style={monthNavBtnStyle} aria-label="Próximo mês">›</button>
        </div>
        <Field label="Fonte"><SelectInput options={["Todas", ...FONTES]} value={fonteFiltro} onChange={e => setFonteFiltro(e.target.value)} /></Field>
      </FilterBar>

      <StatRow stats={[
        { icon: <Wallet size={12} />, label: "Previsto", value: fmtBRL(totals.totalPrevisto) },
        { icon: <Check size={12} />, label: "Pago", value: fmtBRL(totals.totalPago) },
        { icon: totals.diferenca < 0 ? <TrendingDown size={12}/> : <TrendingUp size={12}/>, label: "Diferença", value: fmtBRL(totals.diferenca), tone: totals.diferenca < 0 ? "red" : "green" },
        { icon: <Clock size={12} />, label: "Pago/Pend./Atr.", value: `${totals.pagos}/${totals.pendentes}/${totals.atrasados}` },
      ]} />

      {Object.keys(totals.porFonte).length > 0 && (
        <div className="in" style={{ marginBottom: 20, background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 16 }}>
          <div style={{ fontWeight: 700, fontSize: 12.5, color: C.ink, marginBottom: 10, textTransform: "uppercase", letterSpacing: 0.4 }}>Por fonte de pagamento</div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            <div style={{ width: 100, height: 100, flexShrink: 0 }}>
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={Object.entries(totals.porFonte).map(([name,value])=>({name,value}))} dataKey="value" nameKey="name" innerRadius={28} outerRadius={48} paddingAngle={2}>
                    {Object.keys(totals.porFonte).map((f) => <Cell key={f} fill={fonteColor(f)} />)}
                  </Pie>
                  <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 130 }}>
              {Object.entries(totals.porFonte).map(([f,v],i) => (
                <div key={f} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12 }}>
                  <span style={{ width: 8, height: 8, borderRadius: 3, background: fonteColor(f), flexShrink: 0 }} />
                  <span style={{ color: C.ink, flex: 1 }}>{f}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted, fontSize: 11.5 }}>{fmtBRL(v)}</span>
                </div>
              ))}
            </div>
          </div>
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
                    <span style={{ fontWeight: 700, fontSize: 12.5, color: C.ink, textTransform: "uppercase", letterSpacing: 0.5 }}>
                      {GROUP_LABEL[g.status]}
                    </span>
                    <span style={{ fontSize: 11, fontWeight: 700, color: C.muted, background: C.border, borderRadius: 10, padding: "1px 7px" }}>
                      {g.items.length}
                    </span>
                  </div>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: C.muted }}>{fmtBRL(totalGrupo)}</span>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {g.items.map((c, i) => (
                    <div key={c.id} className="tap in" onClick={() => openEdit(c)} style={{
                      display: "flex", alignItems: "center", gap: 12, background: C.card,
                      boxShadow: shadowSm, borderRadius: 14, padding: "14px 14px 14px 15px",
                      cursor: "pointer", position: "relative", animationDelay: `${Math.min(i,6)*30}ms`
                    }}>
                      <div style={{ position: "absolute", left: 0, top: 10, bottom: 10, width: 3, borderRadius: 3, background: accent }} />
                      <Wallet size={16} color={C.faint} style={{ flexShrink: 0 }} />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 700, color: C.ink, fontSize: 15 }}>{c.descricao}</div>
                        <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>
                          Venc. {fmtDate(c.vencimento)}{c.dataPagamento ? ` · Pago ${fmtDate(c.dataPagamento)}` : ""} · {c.fonte}
                        </div>
                      </div>
                      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 17, color: C.ink, flexShrink: 0 }}>
                        {fmtBRL(c.valorPago ?? c.valorPrevisto)}
                      </div>
                      <button onClick={(e) => { e.stopPropagation(); if (window.confirm("Remover este custo?")) onRemove(c.id); }} title="Remover" style={{
                        border: "none", background: "transparent", color: C.faint, cursor: "pointer",
                        padding: 4, display: "flex", alignItems: "center", flexShrink: 0
                      }}>
                        <Trash2 size={15} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {showForm && (
        <Modal
          title={editingId ? "Editar custo" : "Novo custo"}
          onClose={() => { setShowForm(false); setEditingId(null); }}
          footer={
            <button type="submit" form="custo-form" disabled={saving} style={{ ...addBtnStyle, width: "100%", justifyContent: "center", opacity: saving ? 0.7 : 1, cursor: saving ? "default" : "pointer" }}>
              <Check size={16}/> {saving ? "Salvando..." : (editingId ? "Salvar alterações" : "Salvar custo")}
            </button>
          }
        >
          <form id="custo-form" onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, padding: "4px 0 24px" }}>
            <Field label="Descrição / Categoria"><TextInput style={{gridColumn:"span 2"}} required value={form.descricao} onChange={e=>setForm({...form,descricao:e.target.value})} /></Field>
            <Field label="Vencimento"><TextInput type="date" required value={form.vencimento} onChange={e=>setForm({...form,vencimento:e.target.value})} /></Field>
            <Field label="Data Pagamento"><TextInput type="date" value={form.dataPagamento} onChange={e=>setForm({...form,dataPagamento:e.target.value})} /></Field>
            <Field label="Valor Previsto (R$)"><TextInput type="text" inputMode="decimal" placeholder="0,00" required value={form.valorPrevisto} onChange={e=>setForm({...form,valorPrevisto:e.target.value})} /></Field>
            <Field label="Valor Pago (R$)"><TextInput type="text" inputMode="decimal" placeholder="0,00" value={form.valorPago} onChange={e=>setForm({...form,valorPago:e.target.value})} /></Field>
            <Field label="Fonte de Pagamento"><SelectInput options={FONTES} value={form.fonte} onChange={e=>setForm({...form,fonte:e.target.value})} /></Field>
            <Field label="Status"><SelectInput options={STATUS_CUSTO} value={form.status} onChange={e=>setForm({...form,status:e.target.value})} /></Field>
            <Field label="Observações"><TextArea rows={10} style={{gridColumn:"span 2", height: 260}} value={form.obs} onChange={e=>setForm({...form,obs:e.target.value})} /></Field>
          </form>
        </Modal>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Receitas Mensais                                                   */
/* ------------------------------------------------------------------ */

function ReceitasTab({ receitas, onAdd, onRemove, onUpdate }) {
  const [mesFiltro, setMesFiltro] = useState(currentMonth());
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

  const totals = useMemo(() => {
    const totalPrevisto = filtradas.reduce((s,r) => s + Number(r.valorPrevisto||0), 0);
    const totalRecebido = filtradas.reduce((s,r) => s + Number(r.valorRecebido||0), 0);
    const recebidas = filtradas.filter(r => r.status === "Recebido").length;
    const pendentes = filtradas.filter(r => r.status !== "Recebido").length;
    return { totalPrevisto, totalRecebido, diferenca: totalRecebido-totalPrevisto, recebidas, pendentes };
  }, [filtradas]);

  function blankReceita(mes) {
    return { mes, ano: parseInt(mes.split("-")[0],10), descricao: "", dataRecebimento: "",
      valorPrevisto: "", valorRecebido: "", categoria: "", contaDestino: FONTES[1], status: "Pendente", obs: "" };
  }
  const [form, setForm] = useState(blankReceita(mesFiltro));

  const openNew = () => { setEditingId(null); setForm(blankReceita(mesFiltro)); setShowForm(true); };
  const openEdit = (r) => {
    setEditingId(r.id);
    setForm({
      mes: r.mes, ano: r.ano, descricao: r.descricao, dataRecebimento: r.dataRecebimento || "",
      valorPrevisto: r.valorPrevisto ?? "", valorRecebido: r.valorRecebido ?? "",
      categoria: r.categoria || "", contaDestino: r.contaDestino, status: r.status, obs: r.obs || ""
    });
    setShowForm(true);
  };

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
    if (!confirmDateSanity(form.dataRecebimento, "data de recebimento")) return;
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
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <h2 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 25, letterSpacing: -0.4, color: C.ink }}>Receitas</h2>
        <button className="tap" onClick={openNew} style={addBtnStyle}><Plus size={16} /> Nova</button>
      </div>

      <FilterBar>
        <Field label="Mês"><TextInput type="month" value={mesFiltro} onChange={e => setMesFiltro(e.target.value)} /></Field>
        <div style={{ display: "flex", gap: 6 }}>
          <button type="button" onClick={() => setMesFiltro(m => shiftMonth(m, -1))} style={monthNavBtnStyle} aria-label="Mês anterior">‹</button>
          <button type="button" onClick={() => setMesFiltro(m => shiftMonth(m, 1))} style={monthNavBtnStyle} aria-label="Próximo mês">›</button>
        </div>
      </FilterBar>

      <StatRow stats={[
        { icon: <PiggyBank size={12} />, label: "Previsto", value: fmtBRL(totals.totalPrevisto) },
        { icon: <Check size={12} />, label: "Recebido", value: fmtBRL(totals.totalRecebido) },
        { icon: totals.diferenca < 0 ? <TrendingDown size={12}/> : <TrendingUp size={12}/>, label: "Diferença", value: fmtBRL(totals.diferenca), tone: totals.diferenca < 0 ? "red" : "green" },
        { icon: <Clock size={12} />, label: "Receb./Pend.", value: `${totals.recebidas}/${totals.pendentes}` },
      ]} />

      <CardList
        empty="Nenhuma receita lançada neste mês."
        rows={filtradas}
        onRemove={onRemove}
        onEdit={openEdit}
        render={(r) => (
          <>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 700, color: C.ink, fontSize: 15, textTransform: "uppercase" }}>{r.descricao}</div>
              <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>
                Recebido {fmtDate(r.dataRecebimento)}{r.categoria ? ` · ${r.categoria}` : ""}{r.contaDestino ? ` · ${r.contaDestino}` : ""}
              </div>
            </div>
            <div style={{ textAlign: "right", minWidth: 92 }}>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 14.5, color: C.ink }}>
                {fmtBRL(r.valorRecebido ?? r.valorPrevisto)}
              </div>
              <div style={{ marginTop: 3 }}><StatusPill status={r.status} /></div>
            </div>
          </>
        )}
      />

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
            <Field label="Descrição / Origem"><TextInput style={{gridColumn:"span 2"}} required value={form.descricao} onChange={e=>setForm({...form,descricao:e.target.value})} /></Field>
            <Field label="Data Recebimento"><TextInput type="date" required value={form.dataRecebimento} onChange={e=>setForm({...form,dataRecebimento:e.target.value})} /></Field>
            <Field label="Categoria"><TextInput value={form.categoria} onChange={e=>setForm({...form,categoria:e.target.value})} /></Field>
            <Field label="Valor Previsto (R$)"><TextInput type="text" inputMode="decimal" placeholder="0,00" required value={form.valorPrevisto} onChange={e=>setForm({...form,valorPrevisto:e.target.value})} /></Field>
            <Field label="Valor Recebido (R$)"><TextInput type="text" inputMode="decimal" placeholder="0,00" value={form.valorRecebido} onChange={e=>setForm({...form,valorRecebido:e.target.value})} /></Field>
            <Field label="Conta de Destino"><SelectInput options={FONTES} value={form.contaDestino} onChange={e=>setForm({...form,contaDestino:e.target.value})} /></Field>
            <Field label="Status"><SelectInput options={STATUS_RECEITA} value={form.status} onChange={e=>setForm({...form,status:e.target.value})} /></Field>
            <Field label="Observações"><TextArea rows={10} style={{gridColumn:"span 2", height: 260}} value={form.obs} onChange={e=>setForm({...form,obs:e.target.value})} /></Field>
          </form>
        </Modal>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Resumo Anual                                                       */
/* ------------------------------------------------------------------ */

function MonthCard({ nome, previsto, pago, recebido, saldo, status, temDados, onClick }) {
  const pct = previsto > 0 ? (pago / previsto) * 100 : 0;
  return (
    <div className="tap" onClick={temDados ? onClick : undefined} style={{
      background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 16,
      opacity: temDados ? 1 : 0.55, cursor: temDados ? "pointer" : "default"
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <span style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>{nome}</span>
        {temDados && (
          <span style={{ fontSize: 10.5 }}>{status === "Em dia" ? "✅" : status === "Pendências" ? "⚠️" : ""}</span>
        )}
      </div>
      {temDados ? (
        <>
          <ProgressBar pct={pct} height={6} />
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10.5, color: C.faint, marginTop: 4, marginBottom: 8 }}>
            <span>{fmtBRL(pago)} pago</span>
            <span>{pct.toFixed(0)}%</span>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
            <span style={{ fontSize: 10.5, color: C.muted, fontWeight: 600 }}>Saldo</span>
            <span style={{
              fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 13.5,
              color: saldo < 0 ? C.coral : C.emerald
            }}>{fmtBRL(saldo)}</span>
          </div>
        </>
      ) : (
        <div style={{ fontSize: 11.5, color: C.faint, padding: "8px 0" }}>Sem lançamentos</div>
      )}
    </div>
  );
}

function ResumoTab({ custos, receitas }) {
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
      return {
        mes: nome, mesKey, mesAbr: MESES_ABR[idx], totalPrevistoCusto, totalPagoCusto,
        totalRecebido, saldo: totalRecebido - totalPagoCusto,
        status: cs.length > 0 ? (hasPendencia ? "Pendências" : "Em dia") : "—",
        temDados: cs.length > 0 || rs.length > 0
      };
    });
  }, [custos, receitas, ano]);

  const anual = rows.reduce((acc,r) => ({
    previsto: acc.previsto + r.totalPrevistoCusto, pago: acc.pago + r.totalPagoCusto, recebido: acc.recebido + r.totalRecebido,
  }), { previsto: 0, pago: 0, recebido: 0 });

  const chartData = rows.map(r => ({ mes: r.mesAbr, Saldo: r.saldo }));

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <h2 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 25, letterSpacing: -0.4, color: C.ink }}>Resumo Anual</h2>
        <SelectInput options={anos.map(String)} value={String(ano)} onChange={e=>setAno(parseInt(e.target.value,10))} style={{maxWidth:110}} />
      </div>

      <StatRow stats={[
        { icon: <Wallet size={12}/>, label: "Previsto", value: fmtBRL(anual.previsto) },
        { icon: <Check size={12}/>, label: "Pago", value: fmtBRL(anual.pago) },
        { icon: <PiggyBank size={12}/>, label: "Recebido", value: fmtBRL(anual.recebido) },
        { icon: anual.recebido-anual.pago<0?<TrendingDown size={12}/>:<TrendingUp size={12}/>, label: "Saldo do Ano", value: fmtBRL(anual.recebido-anual.pago), tone: anual.recebido-anual.pago<0?"red":"green" },
      ]} />

      <div style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 10 }}>Saldo mês a mês</div>
        <div style={{ width: "100%", height: 180 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={chartData}>
              <XAxis dataKey="mes" tick={{ fontSize: 11, fill: C.faint }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis hide />
              <ReferenceLine y={0} stroke={C.border} />
              <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12.5, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
              <Bar dataKey="Saldo" radius={[4, 4, 4, 4]}>
                {chartData.map((d, i) => <Cell key={i} fill={d.Saldo < 0 ? C.coral : C.emerald} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))", gap: 12 }}>
        {rows.map(r => (
          <MonthCard key={r.mes} nome={r.mes} previsto={r.totalPrevistoCusto} pago={r.totalPagoCusto}
            recebido={r.totalRecebido} saldo={r.saldo} status={r.status} temDados={r.temDados}
            onClick={() => setMesDetalhe(r)} />
        ))}
      </div>

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
      map[c.fonte] = (map[c.fonte] || 0) + v;
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
                <Pie data={porFonte} dataKey="value" nameKey="name" innerRadius={30} outerRadius={52} paddingAngle={2}>
                  {porFonte.map((f) => <Cell key={f.name} fill={fonteColor(f.name)} />)}
                </Pie>
                <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 120 }}>
            {porFonte.map((f, i) => (
              <div key={f.name} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12 }}>
                <span style={{ width: 8, height: 8, borderRadius: 3, background: fonteColor(f.name), flexShrink: 0 }} />
                <span style={{ color: C.ink, flex: 1 }}>{f.name}</span>
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
      prazo: d.prazo || "", prioridade: d.prioridade, status: d.status, obs: d.obs || ""
    });
    setShowForm(true);
  };

  useEffect(() => {
    if (!jumpTarget) return;
    const item = demandas.find(d => d.id === jumpTarget.id);
    if (item) openEdit(item);
    onConsumeJump();
  }, [jumpTarget]);

  const submit = async (e) => {
    e.preventDefault();
    if (saving) return;
    if (!form.nome) return;
    setSaving(true);
    try {
      if (editingId) { await onUpdate(editingId, form); } else { await onAdd({ ...form, criadoEm: todayISO() }); }
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
  const pctConcluido = totalDemandas > 0 ? (resumo.concluida / totalDemandas) * 100 : 0;
  const COLUMNS = [
    { status: "Pendente", label: "Pendente", count: resumo.pendente, next: "Em Andamento" },
    { status: "Em Andamento", label: "Em Andamento", count: resumo.andamento, next: "Concluída" },
    { status: "Concluída", label: "Concluída", count: resumo.concluida, next: null },
  ];

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <h2 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 25, letterSpacing: -0.4, color: C.ink }}>Demandas</h2>
        <button className="tap" onClick={openNew} style={addBtnStyle}><Plus size={16} /> Nova</button>
      </div>

      <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 18, marginBottom: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 8 }}>
          <span style={{ fontWeight: 700, fontSize: 13, color: C.ink, textTransform: "uppercase", letterSpacing: 0.4 }}>Progresso geral</span>
          <span style={{ fontSize: 12.5, fontWeight: 700, color: C.emerald }}>{pctConcluido.toFixed(0)}% concluído</span>
        </div>
        <ProgressBar pct={pctConcluido} />
        {porPrioridade.length > 0 && (
          <div style={{ display: "flex", gap: 16, marginTop: 14 }}>
            {porPrioridade.map(([p, count]) => (
              <div key={p} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ width: 8, height: 8, borderRadius: "50%", background: PRIORIDADE_COLOR[p] }} />
                <span style={{ fontSize: 12, color: C.muted }}>{p} <strong style={{ color: C.ink }}>{count}</strong></span>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="kanban" style={{ display: "grid", gridTemplateColumns: "1fr", gap: 16 }}>
        {COLUMNS.map(col => {
          const items = demandas.filter(d => d.status === col.status);
          return (
            <div key={col.status}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, padding: "0 2px" }}>
                <span style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, textTransform: "uppercase", letterSpacing: 0.4 }}>{col.label}</span>
                <span style={{
                  fontSize: 11, fontWeight: 700, color: C.muted, background: C.border,
                  borderRadius: 10, padding: "1px 8px"
                }}>{col.count}</span>
              </div>
              {items.length === 0 ? (
                <div style={{ border: `1px dashed ${C.border}`, borderRadius: 14, padding: "20px 14px", textAlign: "center", color: C.faint, fontSize: 12.5 }}>
                  Nada aqui.
                </div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {items.map(d => (
                    <div key={d.id} className="tap in" onClick={() => openEdit(d)} style={{
                      background: C.card, boxShadow: shadowSm, borderRadius: 14, padding: "13px 14px",
                      position: "relative", cursor: "pointer"
                    }}>
                      <div style={{ position: "absolute", left: 0, top: 10, bottom: 10, width: 3, borderRadius: 3, background: PRIORIDADE_COLOR[d.prioridade] || C.border }} />
                      <div style={{ fontWeight: 700, color: C.ink, fontSize: 14.5, marginBottom: 3 }}>{d.nome}</div>
                      <div style={{ fontSize: 11.5, color: C.muted, marginBottom: 3 }}>
                        Prazo {fmtDate(d.prazo)} · {d.prioridade}{d.contato ? ` · ${d.contato}` : ""}{d.detalhes ? ` · ${d.detalhes}` : ""}
                      </div>
                      {d.criadoEm && (
                        <div style={{ fontSize: 10.5, color: C.faint, marginBottom: 10 }}>
                          Criada em {fmtDate(d.criadoEm)}
                        </div>
                      )}
                      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        {col.next && (
                          <button onClick={(e) => { e.stopPropagation(); onStatusChange(d.id, col.next); }} style={{
                            display: "flex", alignItems: "center", gap: 5, border: "none", background: C.border,
                            color: C.ink, borderRadius: 9, padding: "6px 10px", fontSize: 11.5, fontWeight: 700, cursor: "pointer"
                          }}>
                            {col.next} <ArrowRight size={12} />
                          </button>
                        )}
                        <button onClick={(e) => { e.stopPropagation(); if (window.confirm("Remover esta demanda?")) onRemove(d.id); }} title="Remover" style={{
                          border: "none", background: "transparent", color: C.faint, cursor: "pointer",
                          padding: 4, display: "flex", alignItems: "center", marginLeft: "auto"
                        }}>
                          <Trash2 size={15} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

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
            <Field label="Nome da Demanda"><TextInput style={{gridColumn:"span 2"}} required value={form.nome} onChange={e=>setForm({...form,nome:e.target.value})} /></Field>
            <Field label="Contato"><TextInput value={form.contato} onChange={e=>setForm({...form,contato:e.target.value})} /></Field>
            <Field label="Prazo"><TextInput type="date" value={form.prazo} onChange={e=>setForm({...form,prazo:e.target.value})} /></Field>
            <Field label="Prioridade"><SelectInput options={PRIORIDADES} value={form.prioridade} onChange={e=>setForm({...form,prioridade:e.target.value})} /></Field>
            <Field label="Status"><SelectInput options={STATUS_DEMANDA} value={form.status} onChange={e=>setForm({...form,status:e.target.value})} /></Field>
            <Field label="Detalhes"><TextInput style={{gridColumn:"span 2"}} value={form.detalhes} onChange={e=>setForm({...form,detalhes:e.target.value})} /></Field>
            <Field label="Observações"><TextArea rows={10} style={{gridColumn:"span 2", height: 260}} value={form.obs} onChange={e=>setForm({...form,obs:e.target.value})} /></Field>
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
