import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  Plus, Trash2, X, Wallet, TrendingUp, TrendingDown, ClipboardList,
  Calendar, Check, AlertCircle, Clock, Filter, ChevronDown, PiggyBank,
  Home as HomeIcon, ArrowRight, Download, Pencil
} from "lucide-react";
import * as XLSX from "xlsx";

/* ------------------------------------------------------------------ */
/*  Seed data — empty by default. Use "Importar Excel" to load your   */
/*  data from a spreadsheet (nothing sensitive lives in this file).   */
/* ------------------------------------------------------------------ */
const SEED = { custos: [], receitas: [], demandas: [] };

const STORAGE_KEY = "financeiro:app-data";
const PWD_HASH_KEY = "financeiro:pwd-hash";
const PWD_SESSION_KEY = "financeiro:unlocked";

async function sha256(text) {
  const enc = new TextEncoder().encode(text);
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, "0")).join("");
}

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

function LockScreen({ mode, onSetup, onUnlock, error }) {
  const [value, setValue] = useState("");
  const [confirmValue, setConfirmValue] = useState("");

  const submit = (e) => {
    e.preventDefault();
    if (mode === "setup") onSetup(value, confirmValue);
    else onUnlock(value);
  };

  return (
    <div style={{
      minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
      background: "#1E2A24", fontFamily: "'Inter', sans-serif", padding: 20
    }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:wght@600&family=Inter:wght@400;500;600&display=swap');`}</style>
      <form onSubmit={submit} style={{
        background: "#FBFAF6", borderRadius: 16, padding: "28px 24px", width: "100%", maxWidth: 340,
        display: "flex", flexDirection: "column", gap: 14
      }}>
        <div>
          <h2 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontSize: 20, color: "#1E2A24" }}>
            {mode === "setup" ? "Criar senha de acesso" : "Controle Financeiro"}
          </h2>
          <p style={{ margin: "6px 0 0", fontSize: 12.5, color: "#6B7268" }}>
            {mode === "setup"
              ? "Defina uma senha para proteger o acesso a este app. Ela fica só neste navegador."
              : "Digite a senha para continuar."}
          </p>
        </div>
        <Field label={mode === "setup" ? "Nova senha" : "Senha"}>
          <TextInput type="password" autoFocus value={value} onChange={e => setValue(e.target.value)} />
        </Field>
        {mode === "setup" && (
          <Field label="Confirmar senha">
            <TextInput type="password" value={confirmValue} onChange={e => setConfirmValue(e.target.value)} />
          </Field>
        )}
        {error && <div style={{ fontSize: 12.5, color: "#9B2C2C" }}>{error}</div>}
        <button type="submit" style={{ ...addBtnStyle, justifyContent: "center" }}>
          {mode === "setup" ? "Criar senha e entrar" : "Entrar"}
        </button>
      </form>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Main App                                                           */
/* ------------------------------------------------------------------ */

export default function App() {
  const [authMode, setAuthMode] = useState(null); // 'setup' | 'locked' | 'unlocked'
  const [authError, setAuthError] = useState("");
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("custos");
  const [saveError, setSaveError] = useState(false);
  const fileInputRef = useRef(null);

  useEffect(() => {
    const hash = window.localStorage.getItem(PWD_HASH_KEY);
    const sessionOk = window.sessionStorage.getItem(PWD_SESSION_KEY) === "1";
    if (sessionOk) { setAuthMode("unlocked"); return; }
    setAuthMode(hash ? "locked" : "setup");
  }, []);

  const handleSetup = async (value, confirmValue) => {
    setAuthError("");
    if (value.length < 4) { setAuthError("Use pelo menos 4 caracteres."); return; }
    if (value !== confirmValue) { setAuthError("As senhas não coincidem."); return; }
    const hash = await sha256(value);
    window.localStorage.setItem(PWD_HASH_KEY, hash);
    window.sessionStorage.setItem(PWD_SESSION_KEY, "1");
    setAuthMode("unlocked");
  };

  const handleUnlock = async (value) => {
    setAuthError("");
    const hash = await sha256(value);
    const stored = window.localStorage.getItem(PWD_HASH_KEY);
    if (hash === stored) {
      window.sessionStorage.setItem(PWD_SESSION_KEY, "1");
      setAuthMode("unlocked");
    } else {
      setAuthError("Senha incorreta.");
    }
  };

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
  const updateCusto = (id, c) => persist({ ...data, custos: data.custos.map(x => x.id === id ? { ...c, id } : x) });
  const addReceita = (r) => persist({ ...data, receitas: [{ ...r, id: uid("r") }, ...data.receitas] });
  const removeReceita = (id) => persist({ ...data, receitas: data.receitas.filter(r => r.id !== id) });
  const updateReceita = (id, r) => persist({ ...data, receitas: data.receitas.map(x => x.id === id ? { ...r, id } : x) });
  const addDemanda = (d) => persist({ ...data, demandas: [{ ...d, id: uid("d") }, ...data.demandas] });
  const removeDemanda = (id) => persist({ ...data, demandas: data.demandas.filter(d => d.id !== id) });
  const updateDemandaStatus = (id, status) => persist({
    ...data, demandas: data.demandas.map(d => d.id === id ? { ...d, status } : d)
  });

  const handleImportClick = () => fileInputRef.current?.click();
  const handleFileChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const parsed = await importFromExcel(file);
      const hasExisting = data.custos.length || data.receitas.length || data.demandas.length;
      if (hasExisting) {
        const ok = window.confirm(
          `Isso vai substituir os dados atuais (${data.custos.length} custos, ${data.receitas.length} receitas, ${data.demandas.length} demandas) pelos do arquivo importado. Continuar?`
        );
        if (!ok) { e.target.value = ""; return; }
      }
      await persist(parsed);
    } catch (err) {
      alert("Não foi possível importar este arquivo. Confirme se é um Excel exportado por este app.");
    } finally {
      e.target.value = "";
    }
  };

  if (authMode === null) {
    return (
      <div style={{ minHeight: "100vh", background: "#1E2A24" }} />
    );
  }
  if (authMode !== "unlocked") {
    return (
      <LockScreen
        mode={authMode}
        onSetup={handleSetup}
        onUnlock={handleUnlock}
        error={authError}
      />
    );
  }

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
          <div style={{ display: "flex", gap: 8 }}>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileChange}
              style={{ display: "none" }}
            />
            <button onClick={handleImportClick} style={{
              display: "flex", alignItems: "center", gap: 7, background: "transparent", color: "#F7F5F0",
              border: "1px solid #3A4A3F", borderRadius: 9, padding: "9px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer"
            }}>
              <Download size={15} style={{ transform: "rotate(180deg)" }} /> Importar Excel
            </button>
            <button onClick={() => exportToExcel(data)} style={{
              display: "flex", alignItems: "center", gap: 7, background: "#F7F5F0", color: "#1E2A24",
              border: "none", borderRadius: 9, padding: "9px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer"
            }}>
              <Download size={15} /> Exportar Excel
            </button>
          </div>
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
            <CustosTab custos={data.custos} onAdd={addCusto} onRemove={removeCusto} onUpdate={updateCusto} />
          )}
          {tab === "receitas" && (
            <ReceitasTab receitas={data.receitas} onAdd={addReceita} onRemove={removeReceita} onUpdate={updateReceita} />
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

function CustosTab({ custos, onAdd, onRemove, onUpdate }) {
  const [mesFiltro, setMesFiltro] = useState(currentMonth());
  const [fonteFiltro, setFonteFiltro] = useState("Todas");
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);

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

  const submit = (e) => {
    e.preventDefault();
    if (!form.descricao || !form.vencimento || form.valorPrevisto === "") return;
    const payload = { ...form, valorPrevisto: parseFloat(form.valorPrevisto)||0, valorPago: form.valorPago === "" ? null : parseFloat(form.valorPago) };
    if (editingId) {
      onUpdate(editingId, payload);
    } else {
      onAdd(payload);
    }
    setForm(blankCusto(mesFiltro));
    setEditingId(null);
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
        <button onClick={openNew} style={addBtnStyle}>
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
        onEdit={openEdit}
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
        <Modal title={editingId ? "Editar custo" : "Novo custo"} onClose={() => { setShowForm(false); setEditingId(null); }}>
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
              <Check size={16}/> {editingId ? "Salvar alterações" : "Salvar custo"}
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

function ReceitasTab({ receitas, onAdd, onRemove, onUpdate }) {
  const [mesFiltro, setMesFiltro] = useState(currentMonth());
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);

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

  const submit = (e) => {
    e.preventDefault();
    if (!form.descricao || !form.dataRecebimento || form.valorPrevisto === "") return;
    const payload = { ...form, valorPrevisto: parseFloat(form.valorPrevisto)||0, valorRecebido: form.valorRecebido === "" ? null : parseFloat(form.valorRecebido) };
    if (editingId) {
      onUpdate(editingId, payload);
    } else {
      onAdd(payload);
    }
    setForm(blankReceita(mesFiltro));
    setEditingId(null);
    setShowForm(false);
  };

  return (
    <div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "flex-end", marginBottom: 14 }}>
        <Field label="Mês">
          <TextInput type="month" value={mesFiltro} onChange={e => setMesFiltro(e.target.value)} />
        </Field>
        <button onClick={openNew} style={addBtnStyle}>
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
        onEdit={openEdit}
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
        <Modal title={editingId ? "Editar receita" : "Nova receita"} onClose={() => { setShowForm(false); setEditingId(null); }}>
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
              <Check size={16}/> {editingId ? "Salvar alterações" : "Salvar receita"}
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

function RowList({ rows, render, onRemove, onEdit, empty }) {
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
          <div style={{ display: "flex", gap: 2 }}>
            {onEdit && (
              <button onClick={() => onEdit(row)} title="Editar" style={{
                border: "none", background: "transparent", color: "#6B7268", cursor: "pointer",
                padding: 4, display: "flex", alignItems: "center"
              }}>
                <Pencil size={15} />
              </button>
            )}
            <button onClick={() => onRemove(row.id)} title="Remover" style={{
              border: "none", background: "transparent", color: "#B0A990", cursor: "pointer",
              padding: 4, display: "flex", alignItems: "center"
            }}>
              <Trash2 size={16} />
            </button>
          </div>
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
