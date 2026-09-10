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
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, ReferenceLine } from "recharts";

/* ------------------------------------------------------------------ */
/*  Firestore collections — one document per record. Never write an   */
/*  auto-generated "empty" document; only explicit user actions       */
/*  (add / edit / delete / import) touch the database.                */
/* ------------------------------------------------------------------ */
const custosCol = collection(db, "custos");
const receitasCol = collection(db, "receitas");
const demandasCol = collection(db, "demandas");

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
const CHART_PALETTE = [C.emerald, C.amber, C.coral, C.emeraldDeep, C.amberDeep, C.coralDeep, C.inkSoft, C.emeraldLight];
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

function ProgressBar({ pct, color = C.emerald, trackColor = C.border, height = 8 }) {
  const clamped = Math.max(0, Math.min(100, pct));
  const over = pct > 100;
  return (
    <div style={{ width: "100%", height, borderRadius: height, background: trackColor, overflow: "hidden" }}>
      <div style={{
        width: `${clamped}%`, height: "100%", borderRadius: height,
        background: over ? C.coral : color, transition: "width .3s ease"
      }} />
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
  border: "1px solid " + C.border, borderRadius: 10, padding: "10px 12px",
  fontSize: 14, fontFamily: "'Inter', sans-serif", background: C.card,
  color: C.ink, outline: "none",
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
      position: "fixed", inset: 0, background: "#151A2E99", zIndex: 50,
      display: "flex", alignItems: "flex-end", justifyContent: "center"
    }} onClick={onClose}>
      <div
        onClick={e => e.stopPropagation()}
        style={{
          background: C.card, width: "100%", maxWidth: 560, maxHeight: "88vh",
          borderRadius: "22px 22px 0 0", padding: "22px 20px 28px", overflowY: "auto",
          boxShadow: "0 -8px 30px #0004", animation: "slideUp .25s ease-out"
        }}
      >
        <div style={{ width: 36, height: 4, borderRadius: 4, background: C.border, margin: "0 auto 16px" }} />
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
          <h3 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontSize: 21, color: C.ink }}>{title}</h3>
          <button onClick={onClose} style={{
            border: "none", background: C.border, borderRadius: "50%", width: 32, height: 32,
            display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: C.ink
          }}><X size={16} /></button>
        </div>
        {children}
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
        {tab === "dashboard" && <DashboardTab custos={data.custos} receitas={data.receitas} />}
        {tab === "custos" && <CustosTab custos={data.custos} onAdd={addCusto} onRemove={removeCusto} onUpdate={updateCusto} />}
        {tab === "receitas" && <ReceitasTab receitas={data.receitas} onAdd={addReceita} onRemove={removeReceita} onUpdate={updateReceita} />}
        {tab === "resumo" && <ResumoTab custos={data.custos} receitas={data.receitas} />}
        {tab === "demandas" && <DemandasTab demandas={data.demandas} onAdd={addDemanda} onRemove={removeDemanda} onUpdate={updateDemanda} onStatusChange={updateDemandaStatus} />}
      </main>

      <BottomNav tab={tab} setTab={setTab} />

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

function BalanceHero({ saldo, mesLabel, hidden, onToggle, saldosPorMes }) {
  const maxAbs = Math.max(1, ...saldosPorMes.map(s => Math.abs(s)));
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
        <div>
          <div style={{ fontSize: 12.5, color: C.lavender, fontWeight: 700, marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.6 }}>Saldo de {mesLabel}</div>
          <div style={{
            fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 48, letterSpacing: -1.2,
            color: saldo < 0 ? "#FF8A76" : "#4FE8BE", fontVariantNumeric: "tabular-nums", lineHeight: 1.02
          }}>
            {hidden ? "R$ •••••" : fmtBRL(saldo)}
          </div>
        </div>
        <button className="tap" onClick={onToggle} style={{
          width: 38, height: 38, borderRadius: 12, border: "1px solid rgba(255,255,255,0.12)", background: "rgba(255,255,255,0.06)",
          color: C.lavender, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0
        }}>
          {hidden ? <EyeOff size={16} /> : <Eye size={16} />}
        </button>
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
    </div>
  );
}

function DashboardTab({ custos, receitas }) {
  const [hidden, setHidden] = useState(false);
  const ano = new Date().getFullYear();
  const mesAtual = currentMonth();

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
      const key = r.categoria || "Outras";
      map[key] = (map[key] || 0) + v;
    });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
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

  const mesLabel = monthLabel(mesAtual);

  return (
    <div>
      <BalanceHero saldo={kpis.saldo} mesLabel={mesLabel} hidden={hidden} onToggle={() => setHidden(h => !h)} saldosPorMes={saldosPorMes} />

      <StatRow stats={[
        { icon: <Wallet size={12} />, label: "Previsto", value: hidden ? "•••" : fmtBRL(kpis.totalPrevisto) },
        { icon: <Check size={12} />, label: "Pago", value: hidden ? "•••" : fmtBRL(kpis.totalPago) },
        { icon: <PiggyBank size={12} />, label: "Recebido", value: hidden ? "•••" : fmtBRL(kpis.totalRecebido) },
      ]} />

      <div className="grid2" style={{ marginBottom: 20 }}>
        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "0ms" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 8 }}>
            <span style={{ fontWeight: 700, fontSize: 13.5, color: C.ink }}>Orçamento do mês</span>
            <span style={{ fontSize: 12.5, fontWeight: 700, color: kpis.pctUsado > 100 ? C.coral : C.emerald }}>
              {kpis.pctUsado.toFixed(0)}% usado
            </span>
          </div>
          <ProgressBar pct={kpis.pctUsado} />
        </div>

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
                      {porFonte.map((_, i) => <Cell key={i} fill={CHART_PALETTE[i % CHART_PALETTE.length]} />)}
                    </Pie>
                    <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12.5, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 120 }}>
                {porFonte.map((f, i) => (
                  <div key={f.name} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>
                    <span style={{ width: 9, height: 9, borderRadius: 3, background: CHART_PALETTE[i % CHART_PALETTE.length], flexShrink: 0 }} />
                    <span style={{ color: C.ink, flex: 1 }}>{f.name}</span>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted, fontSize: 12 }}>{fmtBRL(f.value)}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        <div className="in" style={{ background: C.card, boxShadow: shadowSm, borderRadius: 16, padding: 20, marginBottom: 20, animationDelay: "120ms" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink, marginBottom: 10 }}>Receitas por categoria — {mesLabel}</div>
          {porCategoria.length === 0 ? (
            <div style={{ color: C.muted, fontSize: 13, padding: "16px 0", textAlign: "center" }}>Sem receitas lançadas neste mês.</div>
          ) : (
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <div style={{ width: 130, height: 130, flexShrink: 0 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie data={porCategoria} dataKey="value" nameKey="name" innerRadius={36} outerRadius={60} paddingAngle={2}>
                      {porCategoria.map((_, i) => <Cell key={i} fill={CHART_PALETTE[i % CHART_PALETTE.length]} />)}
                    </Pie>
                    <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12.5, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 120 }}>
                {porCategoria.map((f, i) => (
                  <div key={f.name} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>
                    <span style={{ width: 9, height: 9, borderRadius: 3, background: CHART_PALETTE[i % CHART_PALETTE.length], flexShrink: 0 }} />
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
                <XAxis dataKey="mes" tick={{ fontSize: 11, fill: C.faint }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis hide />
                <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12.5, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
                <Bar dataKey="Pago" fill={C.coral} radius={[4, 4, 0, 0]} />
                <Bar dataKey="Recebido" fill={C.emerald} radius={[4, 4, 0, 0]} />
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
                return (
                  <div key={c.id}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 4 }}>
                      <span style={{ color: C.ink, fontWeight: 600 }}>{c.descricao}</span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.muted }}>{fmtBRL(v)}</span>
                    </div>
                    <ProgressBar pct={(v / topMax) * 100} color={CHART_PALETTE[i % CHART_PALETTE.length]} height={6} />
                  </div>
                );
              })}
            </div>
          )}
        </div>
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
          <button onClick={(e) => { e.stopPropagation(); onRemove(row.id); }} title="Remover" style={{
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

function CustosTab({ custos, onAdd, onRemove, onUpdate }) {
  const [mesFiltro, setMesFiltro] = useState(currentMonth());
  const [fonteFiltro, setFonteFiltro] = useState("Todas");
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);

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

  const submit = (e) => {
    e.preventDefault();
    if (!form.descricao || !form.vencimento || form.valorPrevisto === "") return;
    const payload = { ...form, valorPrevisto: parseFloat(form.valorPrevisto)||0, valorPago: form.valorPago === "" ? null : parseFloat(form.valorPago) };
    if (editingId) { onUpdate(editingId, payload); } else { onAdd(payload); }
    setForm(blankCusto(mesFiltro));
    setEditingId(null);
    setShowForm(false);
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
                    {Object.keys(totals.porFonte).map((_, i) => <Cell key={i} fill={CHART_PALETTE[i % CHART_PALETTE.length]} />)}
                  </Pie>
                  <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 130 }}>
              {Object.entries(totals.porFonte).map(([f,v],i) => (
                <div key={f} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12 }}>
                  <span style={{ width: 8, height: 8, borderRadius: 3, background: CHART_PALETTE[i % CHART_PALETTE.length], flexShrink: 0 }} />
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
                      <button onClick={(e) => { e.stopPropagation(); onRemove(c.id); }} title="Remover" style={{
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
    if (editingId) { onUpdate(editingId, payload); } else { onAdd(payload); }
    setForm(blankReceita(mesFiltro));
    setEditingId(null);
    setShowForm(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <h2 style={{ margin: 0, fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 25, letterSpacing: -0.4, color: C.ink }}>Receitas</h2>
        <button className="tap" onClick={openNew} style={addBtnStyle}><Plus size={16} /> Nova</button>
      </div>

      <FilterBar>
        <Field label="Mês"><TextInput type="month" value={mesFiltro} onChange={e => setMesFiltro(e.target.value)} /></Field>
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
              <div style={{ fontWeight: 700, color: C.ink, fontSize: 15 }}>{r.descricao}</div>
              <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>
                Receb. {fmtDate(r.dataRecebimento)} · {r.categoria}{r.contaDestino ? ` · ${r.contaDestino}` : ""}
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
                  {porFonte.map((_, i) => <Cell key={i} fill={CHART_PALETTE[i % CHART_PALETTE.length]} />)}
                </Pie>
                <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ borderRadius: 10, border: "none", background: C.card, boxShadow: shadowMd, fontSize: 12, color: C.ink }} itemStyle={{ color: C.ink }} labelStyle={{ color: C.muted }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 120 }}>
            {porFonte.map((f, i) => (
              <div key={f.name} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12 }}>
                <span style={{ width: 8, height: 8, borderRadius: 3, background: CHART_PALETTE[i % CHART_PALETTE.length], flexShrink: 0 }} />
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

function DemandasTab({ demandas, onAdd, onRemove, onUpdate, onStatusChange }) {
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);

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

  const submit = (e) => {
    e.preventDefault();
    if (!form.nome) return;
    if (editingId) { onUpdate(editingId, form); } else { onAdd(form); }
    setForm(blank());
    setEditingId(null);
    setShowForm(false);
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
                      <div style={{ fontSize: 11.5, color: C.muted, marginBottom: 10 }}>
                        Prazo {fmtDate(d.prazo)} · {d.prioridade}{d.contato ? ` · ${d.contato}` : ""}{d.detalhes ? ` · ${d.detalhes}` : ""}
                      </div>
                      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        {col.next && (
                          <button onClick={(e) => { e.stopPropagation(); onStatusChange(d.id, col.next); }} style={{
                            display: "flex", alignItems: "center", gap: 5, border: "none", background: C.border,
                            color: C.ink, borderRadius: 9, padding: "6px 10px", fontSize: 11.5, fontWeight: 700, cursor: "pointer"
                          }}>
                            {col.next} <ArrowRight size={12} />
                          </button>
                        )}
                        <button onClick={(e) => { e.stopPropagation(); onRemove(d.id); }} title="Remover" style={{
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
        <Modal title={editingId ? "Editar demanda" : "Nova demanda"} onClose={() => { setShowForm(false); setEditingId(null); }}>
          <form onSubmit={submit} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <Field label="Nome da Demanda"><TextInput style={{gridColumn:"span 2"}} required value={form.nome} onChange={e=>setForm({...form,nome:e.target.value})} /></Field>
            <Field label="Contato"><TextInput value={form.contato} onChange={e=>setForm({...form,contato:e.target.value})} /></Field>
            <Field label="Prazo"><TextInput type="date" value={form.prazo} onChange={e=>setForm({...form,prazo:e.target.value})} /></Field>
            <Field label="Prioridade"><SelectInput options={PRIORIDADES} value={form.prioridade} onChange={e=>setForm({...form,prioridade:e.target.value})} /></Field>
            <Field label="Status"><SelectInput options={STATUS_DEMANDA} value={form.status} onChange={e=>setForm({...form,status:e.target.value})} /></Field>
            <Field label="Detalhes"><TextInput style={{gridColumn:"span 2"}} value={form.detalhes} onChange={e=>setForm({...form,detalhes:e.target.value})} /></Field>
            <Field label="Observações"><TextInput style={{gridColumn:"span 2"}} value={form.obs} onChange={e=>setForm({...form,obs:e.target.value})} /></Field>
            <button type="submit" style={{ ...addBtnStyle, gridColumn: "span 2", justifyContent: "center", marginTop: 4 }}>
              <Check size={16}/> {editingId ? "Salvar alterações" : "Salvar demanda"}
            </button>
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
