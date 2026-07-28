import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  Plus, Trash2, X, Wallet, TrendingUp, TrendingDown, ClipboardList,
  Calendar, Check, AlertCircle, Clock, Filter, ChevronDown, PiggyBank,
  Home as HomeIcon, ArrowRight, Download
} from "lucide-react";
import * as XLSX from "xlsx";

/* ------------------------------------------------------------------ */
/*  Seed data — extracted from Controle_Financeiro_v9.xlsx            */
/* ------------------------------------------------------------------ */
const SEED = {"custos": [{"id": "c1", "mes": "2026-05", "ano": 2026, "descricao": "Luz - Energia Elétrica", "vencimento": "2026-04-16", "dataPagamento": "2026-05-04", "valorPrevisto": 820.94, "valorPago": 820.94, "fonte": "Dinheiro", "status": "Pago", "obs": null}, {"id": "c2", "mes": "2026-05", "ano": 2026, "descricao": "Prestação Pólo", "vencimento": "2026-05-03", "dataPagamento": "2026-05-04", "valorPrevisto": 1773.86, "valorPago": 1773.86, "fonte": "Transferência", "status": "Pago", "obs": null}, {"id": "c3", "mes": "2026-05", "ano": 2026, "descricao": "Vivo", "vencimento": "2026-05-01", "dataPagamento": "2026-05-04", "valorPrevisto": 133.22, "valorPago": 133.22, "fonte": "Débito Automático", "status": "Pago", "obs": null}, {"id": "c4", "mes": "2026-05", "ano": 2026, "descricao": "Pensão Rubem", "vencimento": "2026-05-05", "dataPagamento": "2026-05-04", "valorPrevisto": 1000, "valorPago": 1000, "fonte": "Transferência", "status": "Pago", "obs": null}, {"id": "c5", "mes": "2026-05", "ano": 2026, "descricao": "Facul Roberto", "vencimento": "2026-05-10", "dataPagamento": "2026-05-08", "valorPrevisto": 631.42, "valorPago": 631.42, "fonte": "Dinheiro", "status": "Pago", "obs": null}, {"id": "c6", "mes": "2026-05", "ano": 2026, "descricao": "Facul Rubem", "vencimento": "2026-05-10", "dataPagamento": "2026-05-08", "valorPrevisto": 791.99, "valorPago": 791.99, "fonte": "Dinheiro", "status": "Pago", "obs": null}, {"id": "c7", "mes": "2026-05", "ano": 2026, "descricao": "Pedal BOSS", "vencimento": "2026-05-08", "dataPagamento": "2026-05-08", "valorPrevisto": 300, "valorPago": 300, "fonte": "PIX", "status": "Pago", "obs": null}, {"id": "c8", "mes": "2026-05", "ano": 2026, "descricao": "Presente Aniversário Zannie", "vencimento": "2026-05-11", "dataPagamento": "2026-05-11", "valorPrevisto": 500, "valorPago": 500, "fonte": "PIX", "status": "Pago", "obs": null}, {"id": "c9", "mes": "2026-05", "ano": 2026, "descricao": "Aluguel", "vencimento": "2026-05-10", "dataPagamento": "2026-05-12", "valorPrevisto": 2675, "valorPago": 2675, "fonte": "Conta Corrente", "status": "Pago", "obs": null}, {"id": "c10", "mes": "2026-05", "ano": 2026, "descricao": "Pensão Zannie", "vencimento": "2026-05-05", "dataPagamento": "2026-05-20", "valorPrevisto": 700, "valorPago": 700, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c11", "mes": "2026-05", "ano": 2026, "descricao": "Prestação Gibson", "vencimento": "2026-05-30", "dataPagamento": "2026-05-22", "valorPrevisto": 1500, "valorPago": 1500, "fonte": "Dinheiro", "status": "Pago", "obs": null}, {"id": "c12", "mes": "2026-05", "ano": 2026, "descricao": "Unimed Rubem", "vencimento": "2026-05-15", "dataPagamento": "2026-05-29", "valorPrevisto": 607.36, "valorPago": 622.3, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c13", "mes": "2026-06", "ano": 2026, "descricao": "Pix Cartório Chapada", "vencimento": "2026-06-01", "dataPagamento": "2026-06-01", "valorPrevisto": 390.27, "valorPago": 390.27, "fonte": "PIX", "status": "Pago", "obs": null}, {"id": "c14", "mes": "2026-06", "ano": 2026, "descricao": "Pix Cantina Zannie", "vencimento": "2026-06-01", "dataPagamento": "2026-06-01", "valorPrevisto": 294, "valorPago": 294, "fonte": "PIX", "status": "Pago", "obs": null}, {"id": "c15", "mes": "2026-06", "ano": 2026, "descricao": "Pix Maria", "vencimento": "2026-06-01", "dataPagamento": "2026-06-01", "valorPrevisto": 750, "valorPago": 750, "fonte": "PIX", "status": "Pago", "obs": null}, {"id": "c16", "mes": "2026-06", "ano": 2026, "descricao": "Prestação Pólo", "vencimento": "2026-06-03", "dataPagamento": "2026-06-02", "valorPrevisto": 1773.86, "valorPago": 1773.86, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c17", "mes": "2026-06", "ano": 2026, "descricao": "Unimed Rubem", "vencimento": "2026-06-15", "dataPagamento": "2026-06-02", "valorPrevisto": 667.36, "valorPago": 667.36, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c18", "mes": "2026-06", "ano": 2026, "descricao": "Pensão Rubem", "vencimento": "2026-06-05", "dataPagamento": "2026-06-02", "valorPrevisto": 1000, "valorPago": 1000, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c19", "mes": "2026-06", "ano": 2026, "descricao": "Pensão Zannie", "vencimento": "2026-06-05", "dataPagamento": "2026-06-02", "valorPrevisto": 700, "valorPago": 700, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c20", "mes": "2026-06", "ano": 2026, "descricao": "Facul Roberto", "vencimento": "2026-06-10", "dataPagamento": "2026-06-05", "valorPrevisto": 631.42, "valorPago": 631.42, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c21", "mes": "2026-06", "ano": 2026, "descricao": "Facul Rubem", "vencimento": "2026-06-10", "dataPagamento": "2026-06-05", "valorPrevisto": 791.99, "valorPago": 791.99, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c22", "mes": "2026-06", "ano": 2026, "descricao": "Condomínio Altos da Serra", "vencimento": "2026-06-10", "dataPagamento": "2026-06-10", "valorPrevisto": 813.47, "valorPago": 813.47, "fonte": "Transferência", "status": "Pago", "obs": null}, {"id": "c23", "mes": "2026-06", "ano": 2026, "descricao": "Condomínio Via Ipiranga", "vencimento": "2026-06-10", "dataPagamento": "2026-06-10", "valorPrevisto": 752.47, "valorPago": 752.47, "fonte": "Transferência", "status": "Pago", "obs": null}, {"id": "c24", "mes": "2026-06", "ano": 2026, "descricao": "Luz - Energia Elétrica", "vencimento": "2026-05-15", "dataPagamento": "2026-06-12", "valorPrevisto": 721.49, "valorPago": 721.49, "fonte": "Dinheiro", "status": "Pago", "obs": null}, {"id": "c25", "mes": "2026-06", "ano": 2026, "descricao": "Aluguel", "vencimento": "2026-06-12", "dataPagamento": "2026-06-12", "valorPrevisto": 2675, "valorPago": 2675, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c26", "mes": "2026-06", "ano": 2026, "descricao": "Prestação Gibson", "vencimento": "2026-06-30", "dataPagamento": "2026-06-26", "valorPrevisto": 1500, "valorPago": 1500, "fonte": "Dinheiro", "status": "Pago", "obs": null}, {"id": "c27", "mes": "2026-05", "ano": 2026, "descricao": "Condomínio Altos da Serra", "vencimento": "2026-05-10", "dataPagamento": "2026-05-29", "valorPrevisto": 813.47, "valorPago": 834.86, "fonte": "Dinheiro", "status": "Pago", "obs": null}, {"id": "c28", "mes": "2026-05", "ano": 2026, "descricao": "Condomínio Via Ipiranga", "vencimento": "2026-05-10", "dataPagamento": "2026-05-29", "valorPrevisto": 752.47, "valorPago": 772.26, "fonte": "Dinheiro", "status": "Pago", "obs": null}, {"id": "c29", "mes": "2026-07", "ano": 2026, "descricao": "Prestação Pólo", "vencimento": "2026-07-03", "dataPagamento": "2026-07-02", "valorPrevisto": 1773.86, "valorPago": 1773.86, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c30", "mes": "2026-07", "ano": 2026, "descricao": "BO's IPTU AP503", "vencimento": "2026-07-02", "dataPagamento": "2026-07-02", "valorPrevisto": 922.85, "valorPago": 922.85, "fonte": "Conta Salário", "status": "Pago", "obs": "209,30 / 209,30 / 504,25"}, {"id": "c31", "mes": "2026-07", "ano": 2026, "descricao": "Unimed Rubem", "vencimento": "2026-07-15", "dataPagamento": "2026-07-02", "valorPrevisto": 607.36, "valorPago": 607.36, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c32", "mes": "2026-07", "ano": 2026, "descricao": "Pensão Rubem", "vencimento": "2026-07-05", "dataPagamento": "2026-07-02", "valorPrevisto": 1000, "valorPago": 1000, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c33", "mes": "2026-07", "ano": 2026, "descricao": "Pensão Zannie", "vencimento": "2026-07-05", "dataPagamento": "2026-07-02", "valorPrevisto": 700, "valorPago": 700, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c34", "mes": "2026-07", "ano": 2026, "descricao": "Luz - Energia Elétrica", "vencimento": "2026-06-16", "dataPagamento": "2026-07-03", "valorPrevisto": 529.57, "valorPago": 529.57, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c35", "mes": "2026-07", "ano": 2026, "descricao": "Pix Estúdio", "vencimento": "2026-07-01", "dataPagamento": "2026-07-06", "valorPrevisto": 155.2, "valorPago": 155.2, "fonte": "PIX", "status": "Pago", "obs": "8 horas de ensaios"}, {"id": "c36", "mes": "2026-07", "ano": 2026, "descricao": "Aluguel", "vencimento": "2026-07-10", "dataPagamento": "2026-07-10", "valorPrevisto": 2675, "valorPago": 2675, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c37", "mes": "2026-07", "ano": 2026, "descricao": "Facul Roberto", "vencimento": "2026-07-10", "dataPagamento": "2026-07-10", "valorPrevisto": 1165.89, "valorPago": 811.1, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c38", "mes": "2026-07", "ano": 2026, "descricao": "Facul Rubem", "vencimento": "2026-07-10", "dataPagamento": "2026-07-10", "valorPrevisto": 1282.85, "valorPago": 862.9, "fonte": "Conta Salário", "status": "Pago", "obs": null}, {"id": "c39", "mes": "2026-07", "ano": 2026, "descricao": "Pix Zannie Férias", "vencimento": "2026-07-16", "dataPagamento": "2026-07-16", "valorPrevisto": 350, "valorPago": 350, "fonte": "PIX", "status": "Pago", "obs": null}, {"id": "c40", "mes": "2026-07", "ano": 2026, "descricao": "Pix Mercado Zannie", "vencimento": "2026-07-22", "dataPagamento": "2026-07-22", "valorPrevisto": 560, "valorPago": 560, "fonte": "PIX", "status": "Pago", "obs": null}, {"id": "c41", "mes": "2026-07", "ano": 2026, "descricao": "Prestação Gibson", "vencimento": "2026-07-27", "dataPagamento": "2026-07-27", "valorPrevisto": 1500, "valorPago": 1500, "fonte": "Dinheiro", "status": "Pago", "obs": null}], "receitas": [{"id": "r1", "mes": "2026-05", "ano": 2026, "descricao": "Aluguel Chapada", "dataRecebimento": "2026-05-08", "valorPrevisto": 3500, "valorRecebido": 2830, "categoria": "Aluguel", "contaDestino": "BB", "status": "Parcial", "obs": null}, {"id": "r2", "mes": "2026-05", "ano": 2026, "descricao": "Aluguel Via Ipiranga", "dataRecebimento": "2026-05-12", "valorPrevisto": 3200, "valorRecebido": 3200, "categoria": "Aluguel", "contaDestino": "Conta Corrente", "status": "Recebido", "obs": null}, {"id": "r3", "mes": "2026-06", "ano": 2026, "descricao": "Aluguel Chapada", "dataRecebimento": "2026-06-10", "valorPrevisto": 3500, "valorRecebido": 713.53, "categoria": "Aluguel", "contaDestino": "Conta Corrente", "status": "Parcial", "obs": null}], "demandas": [{"id": "d1", "nome": "VENDA ITENS DISTRIBUIDORA", "contato": null, "detalhes": null, "prazo": "2026-05-31", "prioridade": "Alta", "status": "Em Andamento", "obs": null}, {"id": "d2", "nome": "QUESTÕES JURÍDICAS INQUILINO VIA IPIRANGA", "contato": null, "detalhes": "Dr Marilton 99241-0810", "prazo": "2026-05-31", "prioridade": "Alta", "status": "Pendente", "obs": null}, {"id": "d3", "nome": "QUESTÕES CASA DE CHAPADA", "contato": null, "detalhes": null, "prazo": "2026-05-31", "prioridade": "Alta", "status": "Concluída", "obs": null}, {"id": "d4", "nome": "PROTESTO CARTÓRIO CHAPADA", "contato": null, "detalhes": null, "prazo": "2026-05-31", "prioridade": "Alta", "status": "Concluída", "obs": null}, {"id": "d5", "nome": "PROTESTO CARTÓRIO CUIABÁ", "contato": null, "detalhes": null, "prazo": "2026-06-30", "prioridade": "Alta", "status": "Pendente", "obs": null}, {"id": "d6", "nome": "IMPOSTO DE RENDA", "contato": null, "detalhes": null, "prazo": "2026-05-08", "prioridade": "Alta", "status": "Concluída", "obs": null}, {"id": "d7", "nome": "CURSO DE CORRETOR", "contato": null, "detalhes": null, "prazo": "2026-05-08", "prioridade": "Alta", "status": "Pendente", "obs": null}, {"id": "d8", "nome": "CONSULTA ENDÓCRINO DAYANNE CAROLINE", "contato": null, "detalhes": null, "prazo": "2026-05-08", "prioridade": "Alta", "status": "Em Andamento", "obs": null}]};

const STORAGE_KEY = "financeiro:app-data";

const MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const FONTES = ["Conta Salário","Conta Corrente","Dinheiro","PIX","Transferência","Débito Automático","Cartão de Crédito","BB"];
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
const uid = (p) => `${p}${Date.now().toString(36)}${Math.random().toString(36).slice(2,7)}`;
const todayISO = () => new Date().toISOString().slice(0,10);
const currentMonth = () => new Date().toISOString().slice(0,7);

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

/* ------------------------------------------------------------------ */
/*  Reusable bits                                                      */
/* ------------------------------------------------------------------ */

function StatusPill({ status }) {
  const map = {
    "Pago": { bg: "#12312222", fg: "#0F5132", dot: "#1E7A47" },
    "Recebido": { bg: "#12312222", fg: "#0F5132", dot: "#1E7A47" },
    "Concluída": { bg: "#12312222", fg: "#0F5132", dot: "#1E7A47" },
    "Parcial": { bg: "#5b3a0022", fg: "#7A4A00", dot: "#B8860B" },
    "Em Andamento": { bg: "#5b3a0022", fg: "#7A4A00", dot: "#B8860B" },
    "Pendente": { bg: "#5b3a0022", fg: "#7A4A00", dot: "#B8860B" },
    "Atrasado": { bg: "#5b141422", fg: "#7A1E1E", dot: "#9B2C2C" },
  };
  const s = map[status] || map["Pendente"];
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 6,
      background: s.bg, color: s.fg, fontSize: 12, fontWeight: 600,
      padding: "3px 9px", borderRadius: 20, whiteSpace: "nowrap"
    }}>
      <span style={{ width: 6, height: 6, borderRadius: "50%", background: s.dot, display: "inline-block" }} />
      {status}
    </span>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 5, fontSize: 13 }}>
      <span style={{ color: "#6B7268", fontWeight: 600, letterSpacing: 0.2 }}>{label}</span>
      {children}
    </label>
  );
}

const inputStyle = {
  border: "1px solid #E4DFD3", borderRadius: 8, padding: "9px 10px",
  fontSize: 14, fontFamily: "'Inter', sans-serif", background: "#FFFFFF",
  color: "#1E2A24", outline: "none",
};

function TextInput(props) {
  return <input {...props} style={{ ...inputStyle, ...(props.style||{}) }} />;
}
function SelectInput({ options, ...props }) {
  return (
    <select {...props} style={{ ...inputStyle, ...(props.style||{}) }}>
      {options.map(o => <option key={o} value={o}>{o}</option>)}
    </select>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div style={{
      position: "fixed", inset: 0, background: "#1E2A2499", zIndex: 50,
      display: "flex", alignItems: "flex-end", justifyContent: "center"
    }} onClick={onClose}>
      <div
        onClick={e => e.stopPropagation()}
        style={{
          background: "#FBFAF6", width: "100%", maxWidth: 560, maxHeight: "88vh",
          borderRadius: "18px 18px 0 0", padding: "20px 20px 28px", overflowY: "auto",
          boxShadow: "0 -8px 30px #0003", animation: "slideUp .25s ease-out"
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontSize: 20, color: "#1E2A24" }}>{title}</h3>
          <button onClick={onClose} style={{
            border: "none", background: "#EFEBE0", borderRadius: "50%", width: 30, height: 30,
            display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: "#1E2A24"
          }}><X size={16} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}

function SummaryCard({ icon, label, value, tone = "ink" }) {
  const colors = {
    ink: "#1E2A24", green: "#146356", red: "#9B2C2C", gold: "#8A6A16"
  };
  return (
    <div style={{
      background: "#FFFFFF", border: "1px solid #E4DFD3", borderRadius: 12,
      padding: "12px 14px", flex: "1 1 140px", minWidth: 140
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, color: "#6B7268", fontSize: 12, fontWeight: 600, marginBottom: 6 }}>
        {icon}{label}
      </div>
      <div style={{
        fontFamily: "'IBM Plex Mono', monospace", fontSize: 17, fontWeight: 600,
        color: colors[tone], fontVariantNumeric: "tabular-nums"
      }}>{value}</div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Main App                                                           */
/* ------------------------------------------------------------------ */

export default function App() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("custos");
  const [saveError, setSaveError] = useState(false);

  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (raw) {
        setData(JSON.parse(raw));
      } else {
        setData(SEED);
        window.localStorage.setItem(STORAGE_KEY, JSON.stringify(SEED));
      }
    } catch (e) {
      setData(SEED);
    } finally {
      setLoading(false);
    }
  }, []);

  const persist = useCallback((next) => {
    setData(next);
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
      setSaveError(false);
    } catch (e) {
      setSaveError(true);
    }
  }, []);

  const addCusto = (c) => persist({ ...data, custos: [{ ...c, id: uid("c") }, ...data.custos] });
  const removeCusto = (id) => persist({ ...data, custos: data.custos.filter(c => c.id !== id) });
  const addReceita = (r) => persist({ ...data, receitas: [{ ...r, id: uid("r") }, ...data.receitas] });
  const removeReceita = (id) => persist({ ...data, receitas: data.receitas.filter(r => r.id !== id) });
  const addDemanda = (d) => persist({ ...data, demandas: [{ ...d, id: uid("d") }, ...data.demandas] });
  const removeDemanda = (id) => persist({ ...data, demandas: data.demandas.filter(d => d.id !== id) });
  const updateDemandaStatus = (id, status) => persist({
    ...data, demandas: data.demandas.map(d => d.id === id ? { ...d, status } : d)
  });

  if (loading || !data) {
    return (
      <div style={{
        minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
        background: "#F7F5F0", fontFamily: "'Inter', sans-serif", color: "#6B7268"
      }}>
        Carregando…
      </div>
    );
  }

  return (
    <div style={{ minHeight: "100vh", background: "#F7F5F0", fontFamily: "'Inter', sans-serif", paddingBottom: 40 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:wght@400;500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        * { box-sizing: border-box; }
        body { margin: 0; }
        @keyframes slideUp { from { transform: translateY(24px); opacity: 0 } to { transform: translateY(0); opacity: 1 } }
        table { border-collapse: collapse; width: 100%; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: #D8D2C2; border-radius: 4px; }
        button:focus-visible, select:focus-visible, input:focus-visible { outline: 2px solid #146356; outline-offset: 1px; }
      `}</style>

      {/* Header */}
      <header style={{
        background: "#1E2A24", color: "#F7F5F0", padding: "22px 18px 18px",
        borderBottom: "3px double #3A4A3F"
      }}>
        <div style={{ maxWidth: 900, margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "flex-end", gap: 12, flexWrap: "wrap" }}>
          <div>
            <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
              <span style={{ fontFamily: "'Fraunces', serif", fontSize: 15, color: "#8FAE9F" }}>Livro-Razão</span>
            </div>
            <h1 style={{ margin: "2px 0 0", fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 26 }}>
              Controle Financeiro
            </h1>
          </div>
          <button onClick={() => exportToExcel(data)} style={{
            display: "flex", alignItems: "center", gap: 7, background: "#F7F5F0", color: "#1E2A24",
            border: "none", borderRadius: 9, padding: "9px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer"
          }}>
            <Download size={15} /> Exportar Excel
          </button>
        </div>
      </header>

      {/* Tabs */}
      <nav style={{
        maxWidth: 900, margin: "0 auto", display: "flex", gap: 4, padding: "14px 18px 0",
        overflowX: "auto"
      }}>
        {[
          { key: "custos", label: "Custos", icon: <Wallet size={15} /> },
          { key: "receitas", label: "Receitas", icon: <PiggyBank size={15} /> },
          { key: "resumo", label: "Resumo Anual", icon: <TrendingUp size={15} /> },
          { key: "demandas", label: "Demandas", icon: <ClipboardList size={15} /> },
        ].map(t => (
          <button key={t.key} onClick={() => setTab(t.key)} style={{
            display: "flex", alignItems: "center", gap: 6,
            padding: "9px 14px", borderRadius: "10px 10px 0 0", border: "none",
            cursor: "pointer", fontSize: 13.5, fontWeight: 600, whiteSpace: "nowrap",
            background: tab === t.key ? "#FBFAF6" : "transparent",
            color: tab === t.key ? "#1E2A24" : "#6B7268",
            borderBottom: tab === t.key ? "2px solid #146356" : "2px solid transparent"
          }}>
            {t.icon}{t.label}
          </button>
        ))}
      </nav>

      <main style={{ maxWidth: 900, margin: "0 auto", padding: "0 18px" }}>
        <div style={{ background: "#FBFAF6", border: "1px solid #E4DFD3", borderTop: "none", borderRadius: "0 0 14px 14px", padding: 18 }}>
          {tab === "custos" && (
            <CustosTab custos={data.custos} onAdd={addCusto} onRemove={removeCusto} />
          )}
          {tab === "receitas" && (
            <ReceitasTab receitas={data.receitas} onAdd={addReceita} onRemove={removeReceita} />
          )}
          {tab === "resumo" && (
            <ResumoTab custos={data.custos} receitas={data.receitas} />
          )}
          {tab === "demandas" && (
            <DemandasTab demandas={data.demandas} onAdd={addDemanda} onRemove={removeDemanda} onStatusChange={updateDemandaStatus} />
          )}
        </div>
      </main>

      {saveError && (
        <div style={{
          position: "fixed", bottom: 16, left: "50%", transform: "translateX(-50%)",
          background: "#7A1E1E", color: "#fff", padding: "8px 16px", borderRadius: 10,
          fontSize: 13, display: "flex", alignItems: "center", gap: 6
        }}>
          <AlertCircle size={15} /> Não foi possível salvar. Tente novamente.
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Custos Mensais                                                     */
/* ------------------------------------------------------------------ */

function CustosTab({ custos, onAdd, onRemove }) {
  const [mesFiltro, setMesFiltro] = useState(currentMonth());
  const [fonteFiltro, setFonteFiltro] = useState("Todas");
  const [showForm, setShowForm] = useState(false);

  const filtrados = useMemo(() => {
    const list = custos.filter(c =>
      c.mes === mesFiltro && (fonteFiltro === "Todas" || c.fonte === fonteFiltro)
    );
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
    filtrados.forEach(c => {
      porFonte[c.fonte] = (porFonte[c.fonte]||0) + Number(c.valorPago||0);
    });
    return { totalPrevisto, totalPago, diferenca: totalPago-totalPrevisto, pagos, pendentes, atrasados, porFonte };
  }, [filtrados]);

  const [form, setForm] = useState(blankCusto(mesFiltro));
  function blankCusto(mes) {
    return { mes, ano: parseInt(mes.split("-")[0],10), descricao: "", vencimento: "", dataPagamento: "",
      valorPrevisto: "", valorPago: "", fonte: FONTES[0], status: "Pendente", obs: "" };
  }

  const submit = (e) => {
    e.preventDefault();
    if (!form.descricao || !form.vencimento || form.valorPrevisto === "") return;
    onAdd({ ...form, valorPrevisto: parseFloat(form.valorPrevisto)||0, valorPago: form.valorPago === "" ? null : parseFloat(form.valorPago) });
    setForm(blankCusto(mesFiltro));
    setShowForm(false);
  };

  return (
    <div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "flex-end", marginBottom: 14 }}>
        <Field label="Mês">
          <TextInput type="month" value={mesFiltro} onChange={e => setMesFiltro(e.target.value)} />
        </Field>
        <Field label="Fonte">
          <SelectInput options={["Todas", ...FONTES]} value={fonteFiltro} onChange={e => setFonteFiltro(e.target.value)} />
        </Field>
        <button onClick={() => { setForm(blankCusto(mesFiltro)); setShowForm(true); }} style={addBtnStyle}>
          <Plus size={16} /> Novo custo
        </button>
      </div>

      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
        <SummaryCard icon={<Wallet size={13} />} label="Previsto" value={fmtBRL(totals.totalPrevisto)} />
        <SummaryCard icon={<Check size={13} />} label="Pago" value={fmtBRL(totals.totalPago)} tone="green" />
        <SummaryCard icon={totals.diferenca < 0 ? <TrendingDown size={13}/> : <TrendingUp size={13}/>} label="Diferença" value={fmtBRL(totals.diferenca)} tone={totals.diferenca < 0 ? "red" : "green"} />
        <SummaryCard icon={<Clock size={13} />} label="Pagos / Pendentes / Atrasados" value={`${totals.pagos} / ${totals.pendentes} / ${totals.atrasados}`} />
      </div>

      {Object.keys(totals.porFonte).length > 0 && (
        <div style={{ marginBottom: 16, fontSize: 12.5, color: "#6B7268" }}>
          <span style={{ fontWeight: 700, color: "#1E2A24" }}>Resumo por fonte: </span>
          {Object.entries(totals.porFonte).map(([f,v],i) => (
            <span key={f}>{i>0 && " · "}{f}: <strong style={{ color: "#1E2A24" }}>{fmtBRL(v)}</strong></span>
          ))}
        </div>
      )}

      <RowList
        empty="Nenhum custo lançado neste mês/fonte."
        rows={filtrados}
        onRemove={onRemove}
        render={(c) => (
          <>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 600, color: "#1E2A24", fontSize: 14 }}>{c.descricao}</div>
              <div style={{ fontSize: 12, color: "#6B7268" }}>
                Venc. {fmtDate(c.vencimento)} · Pago em {fmtDate(c.dataPagamento)} · {c.fonte}{c.obs ? ` · ${c.obs}` : ""}
              </div>
            </div>
            <div style={{ textAlign: "right", minWidth: 100 }}>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, fontSize: 14, color: "#1E2A24" }}>
                {fmtBRL(c.valorPago ?? c.valorPrevisto)}
              </div>
              <StatusPill status={c.status} />
            </div>
          </>
        )}
      />

      {showForm && (
        <Modal title="Novo custo" onClose={() => setShowForm(false)}>
          <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <Field label="Descrição / Categoria"><TextInput style={{gridColumn:"span 2"}} required value={form.descricao} onChange={e=>setForm({...form,descricao:e.target.value})} /></Field>
            <Field label="Vencimento"><TextInput type="date" required value={form.vencimento} onChange={e=>setForm({...form,vencimento:e.target.value})} /></Field>
            <Field label="Data Pagamento"><TextInput type="date" value={form.dataPagamento} onChange={e=>setForm({...form,dataPagamento:e.target.value})} /></Field>
            <Field label="Valor Previsto (R$)"><TextInput type="number" step="0.01" required value={form.valorPrevisto} onChange={e=>setForm({...form,valorPrevisto:e.target.value})} /></Field>
            <Field label="Valor Pago (R$)"><TextInput type="number" step="0.01" value={form.valorPago} onChange={e=>setForm({...form,valorPago:e.target.value})} /></Field>
            <Field label="Fonte de Pagamento"><SelectInput options={FONTES} value={form.fonte} onChange={e=>setForm({...form,fonte:e.target.value})} /></Field>
            <Field label="Status"><SelectInput options={STATUS_CUSTO} value={form.status} onChange={e=>setForm({...form,status:e.target.value})} /></Field>
            <Field label="Observações"><TextInput style={{gridColumn:"span 2"}} value={form.obs} onChange={e=>setForm({...form,obs:e.target.value})} /></Field>
            <button type="submit" style={{ ...addBtnStyle, gridColumn: "span 2", justifyContent: "center", marginTop: 4 }}>
              <Check size={16}/> Salvar custo
            </button>
          </form>
        </Modal>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Receitas Mensais                                                   */
/* ------------------------------------------------------------------ */

function ReceitasTab({ receitas, onAdd, onRemove }) {
  const [mesFiltro, setMesFiltro] = useState(currentMonth());
  const [showForm, setShowForm] = useState(false);

  const filtradas = useMemo(() => receitas.filter(r => r.mes === mesFiltro), [receitas, mesFiltro]);

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

  const submit = (e) => {
    e.preventDefault();
    if (!form.descricao || !form.dataRecebimento || form.valorPrevisto === "") return;
    onAdd({ ...form, valorPrevisto: parseFloat(form.valorPrevisto)||0, valorRecebido: form.valorRecebido === "" ? null : parseFloat(form.valorRecebido) });
    setForm(blankReceita(mesFiltro));
    setShowForm(false);
  };

  return (
    <div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "flex-end", marginBottom: 14 }}>
        <Field label="Mês">
          <TextInput type="month" value={mesFiltro} onChange={e => setMesFiltro(e.target.value)} />
        </Field>
        <button onClick={() => { setForm(blankReceita(mesFiltro)); setShowForm(true); }} style={addBtnStyle}>
          <Plus size={16} /> Nova receita
        </button>
      </div>

      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
        <SummaryCard icon={<PiggyBank size={13} />} label="Previsto" value={fmtBRL(totals.totalPrevisto)} />
        <SummaryCard icon={<Check size={13} />} label="Recebido" value={fmtBRL(totals.totalRecebido)} tone="green" />
        <SummaryCard icon={totals.diferenca < 0 ? <TrendingDown size={13}/> : <TrendingUp size={13}/>} label="Diferença" value={fmtBRL(totals.diferenca)} tone={totals.diferenca < 0 ? "red" : "green"} />
        <SummaryCard icon={<Clock size={13} />} label="Recebidas / Pendentes" value={`${totals.recebidas} / ${totals.pendentes}`} />
      </div>

      <RowList
        empty="Nenhuma receita lançada neste mês."
        rows={filtradas}
        onRemove={onRemove}
        render={(r) => (
          <>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 600, color: "#1E2A24", fontSize: 14 }}>{r.descricao}</div>
              <div style={{ fontSize: 12, color: "#6B7268" }}>
                Receb. {fmtDate(r.dataRecebimento)} · {r.categoria}{r.contaDestino ? ` · ${r.contaDestino}` : ""}
              </div>
            </div>
            <div style={{ textAlign: "right", minWidth: 100 }}>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, fontSize: 14, color: "#1E2A24" }}>
                {fmtBRL(r.valorRecebido ?? r.valorPrevisto)}
              </div>
              <StatusPill status={r.status} />
            </div>
          </>
        )}
      />

      {showForm && (
        <Modal title="Nova receita" onClose={() => setShowForm(false)}>
          <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <Field label="Descrição / Origem"><TextInput style={{gridColumn:"span 2"}} required value={form.descricao} onChange={e=>setForm({...form,descricao:e.target.value})} /></Field>
            <Field label="Data Recebimento"><TextInput type="date" required value={form.dataRecebimento} onChange={e=>setForm({...form,dataRecebimento:e.target.value})} /></Field>
            <Field label="Categoria"><TextInput value={form.categoria} onChange={e=>setForm({...form,categoria:e.target.value})} /></Field>
            <Field label="Valor Previsto (R$)"><TextInput type="number" step="0.01" required value={form.valorPrevisto} onChange={e=>setForm({...form,valorPrevisto:e.target.value})} /></Field>
            <Field label="Valor Recebido (R$)"><TextInput type="number" step="0.01" value={form.valorRecebido} onChange={e=>setForm({...form,valorRecebido:e.target.value})} /></Field>
            <Field label="Conta de Destino"><SelectInput options={FONTES} value={form.contaDestino} onChange={e=>setForm({...form,contaDestino:e.target.value})} /></Field>
            <Field label="Status"><SelectInput options={STATUS_RECEITA} value={form.status} onChange={e=>setForm({...form,status:e.target.value})} /></Field>
            <Field label="Observações"><TextInput style={{gridColumn:"span 2"}} value={form.obs} onChange={e=>setForm({...form,obs:e.target.value})} /></Field>
            <button type="submit" style={{ ...addBtnStyle, gridColumn: "span 2", justifyContent: "center", marginTop: 4 }}>
              <Check size={16}/> Salvar receita
            </button>
          </form>
        </Modal>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Resumo Anual                                                       */
/* ------------------------------------------------------------------ */

function ResumoTab({ custos, receitas }) {
  const anos = useMemo(() => {
    const s = new Set([...custos, ...receitas].map(x => x.ano).filter(Boolean));
    s.add(new Date().getFullYear());
    return Array.from(s).sort();
  }, [custos, receitas]);
  const [ano, setAno] = useState(new Date().getFullYear());

  const rows = useMemo(() => {
    return MESES.map((nome, idx) => {
      const mesKey = `${ano}-${String(idx+1).padStart(2,"0")}`;
      const cs = custos.filter(c => c.mes === mesKey);
      const rs = receitas.filter(r => r.mes === mesKey);
      const totalPrevistoCusto = cs.reduce((s,c)=>s+Number(c.valorPrevisto||0),0);
      const totalPagoCusto = cs.reduce((s,c)=>s+Number(c.valorPago||0),0);
      const totalRecebido = rs.reduce((s,r)=>s+Number(r.valorRecebido||0),0);
      const qtdPagos = cs.filter(c=>c.status==="Pago").length;
      const hasPendencia = cs.some(c => c.status !== "Pago");
      let status = "—";
      if (cs.length > 0) status = hasPendencia ? "⚠️ Pendências" : "✅ Em dia";
      return {
        mes: nome, totalPrevistoCusto, totalPagoCusto,
        diferenca: totalPagoCusto-totalPrevistoCusto, qtdPagos, status,
        totalRecebido, saldo: totalRecebido - totalPagoCusto
      };
    });
  }, [custos, receitas, ano]);

  const anual = rows.reduce((acc,r) => ({
    previsto: acc.previsto + r.totalPrevistoCusto,
    pago: acc.pago + r.totalPagoCusto,
    recebido: acc.recebido + r.totalRecebido,
  }), { previsto: 0, pago: 0, recebido: 0 });

  return (
    <div>
      <div style={{ marginBottom: 14 }}>
        <Field label="Ano">
          <SelectInput options={anos.map(String)} value={String(ano)} onChange={e=>setAno(parseInt(e.target.value,10))} style={{maxWidth:120}} />
        </Field>
      </div>

      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 18 }}>
        <SummaryCard icon={<Wallet size={13}/>} label="Total Previsto (custos)" value={fmtBRL(anual.previsto)} />
        <SummaryCard icon={<Check size={13}/>} label="Total Pago" value={fmtBRL(anual.pago)} tone="green" />
        <SummaryCard icon={<PiggyBank size={13}/>} label="Total Recebido" value={fmtBRL(anual.recebido)} />
        <SummaryCard icon={anual.recebido-anual.pago<0?<TrendingDown size={13}/>:<TrendingUp size={13}/>} label="Saldo do Ano" value={fmtBRL(anual.recebido-anual.pago)} tone={anual.recebido-anual.pago<0?"red":"green"} />
      </div>

      <div style={{ overflowX: "auto" }}>
        <table>
          <thead>
            <tr style={{ borderBottom: "2px solid #1E2A24" }}>
              {["Mês","Previsto","Pago","Diferença","Recebido","Saldo","Status"].map(h => (
                <th key={h} style={{ textAlign: h==="Mês"?"left":"right", padding: "8px 6px", fontSize: 12, color: "#6B7268", fontWeight: 700, whiteSpace:"nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(r => (
              <tr key={r.mes} style={{ borderBottom: "1px solid #E4DFD3" }}>
                <td style={{ padding: "8px 6px", fontSize: 13.5, fontWeight: 600, color: "#1E2A24" }}>{r.mes}</td>
                <td style={numCell}>{fmtBRL(r.totalPrevistoCusto)}</td>
                <td style={numCell}>{fmtBRL(r.totalPagoCusto)}</td>
                <td style={{ ...numCell, color: r.diferenca < 0 ? "#9B2C2C" : "#146356" }}>{fmtBRL(r.diferenca)}</td>
                <td style={numCell}>{fmtBRL(r.totalRecebido)}</td>
                <td style={{ ...numCell, color: r.saldo < 0 ? "#9B2C2C" : "#146356" }}>{fmtBRL(r.saldo)}</td>
                <td style={{ padding: "8px 6px", fontSize: 13, textAlign: "right", whiteSpace: "nowrap" }}>{r.status}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
const numCell = { padding: "8px 6px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, textAlign: "right", color: "#1E2A24", fontVariantNumeric: "tabular-nums" };

/* ------------------------------------------------------------------ */
/*  Demandas                                                            */
/* ------------------------------------------------------------------ */

function DemandasTab({ demandas, onAdd, onRemove, onStatusChange }) {
  const [statusFiltro, setStatusFiltro] = useState("Todos");
  const [showForm, setShowForm] = useState(false);

  const filtradas = useMemo(() => demandas.filter(d => statusFiltro === "Todos" || d.status === statusFiltro), [demandas, statusFiltro]);

  const resumo = useMemo(() => ({
    pendente: demandas.filter(d=>d.status==="Pendente").length,
    andamento: demandas.filter(d=>d.status==="Em Andamento").length,
    concluida: demandas.filter(d=>d.status==="Concluída").length,
  }), [demandas]);

  function blank() { return { nome: "", contato: "", detalhes: "", prazo: "", prioridade: "Alta", status: "Pendente", obs: "" }; }
  const [form, setForm] = useState(blank());

  const submit = (e) => {
    e.preventDefault();
    if (!form.nome) return;
    onAdd(form);
    setForm(blank());
    setShowForm(false);
  };

  return (
    <div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "flex-end", marginBottom: 14 }}>
        <Field label="Status">
          <SelectInput options={["Todos", ...STATUS_DEMANDA]} value={statusFiltro} onChange={e=>setStatusFiltro(e.target.value)} />
        </Field>
        <button onClick={() => { setForm(blank()); setShowForm(true); }} style={addBtnStyle}>
          <Plus size={16} /> Nova demanda
        </button>
      </div>

      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
        <SummaryCard icon={<Clock size={13}/>} label="Pendentes" value={resumo.pendente} tone="gold" />
        <SummaryCard icon={<ArrowRight size={13}/>} label="Em Andamento" value={resumo.andamento} tone="gold" />
        <SummaryCard icon={<Check size={13}/>} label="Concluídas" value={resumo.concluida} tone="green" />
      </div>

      <RowList
        empty="Nenhuma demanda com este status."
        rows={filtradas}
        onRemove={onRemove}
        render={(d) => (
          <>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 600, color: "#1E2A24", fontSize: 14 }}>{d.nome}</div>
              <div style={{ fontSize: 12, color: "#6B7268" }}>
                Prazo {fmtDate(d.prazo)} · Prioridade {d.prioridade}{d.contato ? ` · ${d.contato}` : ""}{d.detalhes ? ` · ${d.detalhes}` : ""}
              </div>
            </div>
            <div style={{ textAlign: "right", minWidth: 140, display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6 }}>
              <select value={d.status} onChange={e => onStatusChange(d.id, e.target.value)} style={{ ...inputStyle, padding: "4px 8px", fontSize: 12 }}>
                {STATUS_DEMANDA.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          </>
        )}
      />

      {showForm && (
        <Modal title="Nova demanda" onClose={() => setShowForm(false)}>
          <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <Field label="Nome da Demanda"><TextInput style={{gridColumn:"span 2"}} required value={form.nome} onChange={e=>setForm({...form,nome:e.target.value})} /></Field>
            <Field label="Contato"><TextInput value={form.contato} onChange={e=>setForm({...form,contato:e.target.value})} /></Field>
            <Field label="Prazo"><TextInput type="date" value={form.prazo} onChange={e=>setForm({...form,prazo:e.target.value})} /></Field>
            <Field label="Prioridade"><SelectInput options={PRIORIDADES} value={form.prioridade} onChange={e=>setForm({...form,prioridade:e.target.value})} /></Field>
            <Field label="Status"><SelectInput options={STATUS_DEMANDA} value={form.status} onChange={e=>setForm({...form,status:e.target.value})} /></Field>
            <Field label="Detalhes"><TextInput style={{gridColumn:"span 2"}} value={form.detalhes} onChange={e=>setForm({...form,detalhes:e.target.value})} /></Field>
            <Field label="Observações"><TextInput style={{gridColumn:"span 2"}} value={form.obs} onChange={e=>setForm({...form,obs:e.target.value})} /></Field>
            <button type="submit" style={{ ...addBtnStyle, gridColumn: "span 2", justifyContent: "center", marginTop: 4 }}>
              <Check size={16}/> Salvar demanda
            </button>
          </form>
        </Modal>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Shared row list                                                     */
/* ------------------------------------------------------------------ */

function RowList({ rows, render, onRemove, empty }) {
  if (rows.length === 0) {
    return (
      <div style={{
        border: "1px dashed #D8D2C2", borderRadius: 10, padding: "28px 16px",
        textAlign: "center", color: "#6B7268", fontSize: 13.5
      }}>{empty}</div>
    );
  }
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      {rows.map(row => (
        <div key={row.id} style={{
          display: "flex", alignItems: "center", gap: 12, background: "#FFFFFF",
          border: "1px solid #E4DFD3", borderRadius: 10, padding: "10px 12px"
        }}>
          {render(row)}
          <button onClick={() => onRemove(row.id)} title="Remover" style={{
            border: "none", background: "transparent", color: "#B0A990", cursor: "pointer",
            padding: 4, display: "flex", alignItems: "center"
          }}>
            <Trash2 size={16} />
          </button>
        </div>
      ))}
    </div>
  );
}

const addBtnStyle = {
  display: "flex", alignItems: "center", gap: 6, background: "#146356", color: "#fff",
  border: "none", borderRadius: 9, padding: "9px 14px", fontSize: 13.5, fontWeight: 600,
  cursor: "pointer"
};
