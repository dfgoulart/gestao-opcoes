import { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, ReferenceLine, Cell } from "recharts";
import { msalInstance, loginRequest, fetchPlanilha } from "./auth";

// ── Helpers ───────────────────────────────────────────────────
function fmt(val, decimals = 2) {
  if (val == null || val === "" || (typeof val === "string" && val.includes("N/A"))) return "—";
  const n = parseFloat(val);
  if (isNaN(n)) return "—";
  return n.toLocaleString("pt-BR", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}
function fmtR(val, decimals = 2) {
  const s = fmt(val, decimals);
  return s === "—" ? s : "R$ " + s;
}
function fmtPct(val) {
  if (val == null || isNaN(parseFloat(val))) return "—";
  return (parseFloat(val) * 100).toFixed(1) + "%";
}
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === "number") return new Date(Math.round((val - 25569) * 86400 * 1000));
  return null;
}
function fmtDate(val) {
  const d = parseDate(val);
  return d ? d.toLocaleDateString("pt-BR") : "—";
}
function isNA(val) {
  if (val == null) return true;
  if (typeof val === "string" && (val.includes("N/A") || val.includes("REF") || val === "")) return true;
  return false;
}

// Ordem cronológica dos meses
const MONTH_ORDER = {
  "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
  "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
  "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
};

function sortMonths(months) {
  return [...months].sort((a, b) => {
    if (a.ano !== b.ano) return a.ano - b.ano;
    return (MONTH_ORDER[a.mes] || 0) - (MONTH_ORDER[b.mes] || 0);
  });
}

// YTD = meses do ano corrente
function calcYTD(months) {
  const currentYear = new Date().getFullYear();
  return months
    .filter(m => m.ano === currentYear)
    .reduce((s, m) => s + (parseFloat(m.resultado) || 0), 0);
}

// ── Parser ────────────────────────────────────────────────────
function parseWorkbook(wb) {
  const result = { plAnual: [], positionsOpen: [], irpf: [], dre: null };

  const wsRes = wb.Sheets["P&L Anual (modelo resumido)"];
  if (wsRes) {
    const data = XLSX.utils.sheet_to_json(wsRes, { header: 1, defval: null });
    const months = [];
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (row[0] && typeof row[0] === "string" && row[0] !== "TOTAL")
        months.push({ mes: row[0], resultado: row[1], ano: 2025 });
      if (row[3] && typeof row[3] === "string" && row[3] !== "TOTAL")
        months.push({ mes: row[3], resultado: row[4], ano: 2026 });
    }
    // Ordena cronologicamente
    result.plAnual = sortMonths(
      months.filter(m => m.resultado != null && !isNaN(parseFloat(m.resultado)))
    );
  }

  const wsCtrl = wb.Sheets["Controle de posições"];
  if (wsCtrl) {
    const data = XLSX.utils.sheet_to_json(wsCtrl, { header: 1, defval: null });
    const positions = [];
    for (let r = 424; r < data.length; r++) {
      const row = data[r];
      if (!row[0] || typeof row[0] !== "string") continue;
      if (["TOTAL","Marcação","DRE","Receita","Custo","Resultado"].some(s => row[0].startsWith(s))) break;
      if (!row[1]) continue;
      positions.push({
        ticker: row[0], acao: row[1], tipo: row[2], qtd: row[3],
        precoVenda: row[5], strike: row[6], mktVsStrike: row[10], vencimento: row[11],
        mtm: row[12], plUnit: row[13], plTotal: row[14], plPct: row[15],
      });
    }
    result.positionsOpen = positions;

    let dreReceita = null, dreResultado = null, dreResultadoSemMtM = null;
    let recompras = 0, mtmTotal = 0, mtmDisponivel = false, inCustos = false;
    for (let r = 494; r < data.length; r++) {
      const row = data[r];
      if (!row[0]) continue;
      if (row[0] === "Receita") { dreReceita = row[2]; continue; }
      if (row[0] === "Custos e despesas") { inCustos = true; continue; }
      if (row[0] === "Resultado sem MtM") { dreResultadoSemMtM = row[2]; continue; }
      if (row[0] === "Resultado" && dreResultado === null) { dreResultado = row[2]; continue; }
      if (inCustos) {
        if (row[1] === "MtM") {
          const v = parseFloat(row[2]);
          if (!isNaN(v)) { mtmTotal += v; mtmDisponivel = true; }
        } else if (row[1] instanceof Date || (typeof row[1] === "number" && row[1] > 40000)) {
          const v = parseFloat(row[2]);
          if (!isNaN(v) && v < 0) recompras += v;
        }
      }
    }
    result.dre = {
      receita: dreReceita, recompras, mtmTotal: mtmDisponivel ? mtmTotal : null,
      resultado: !isNA(dreResultado) ? dreResultado : (dreReceita != null ? dreReceita + recompras + (mtmDisponivel ? mtmTotal : 0) : null),
      resultadoSemMtM: dreResultadoSemMtM, mtmDisponivel,
    };
  }

  // IRPF — nova estrutura: linha por mês
  // Colunas: Ano ref | Mês operação | Mês recolhimento | Valor devido (MyCapital) | Valor pago | OBS
  const wsIR = wb.Sheets["IRPF Ações + Opções"];
  if (wsIR) {
    const data = XLSX.utils.sheet_to_json(wsIR, { header: 1, defval: null });
    const rows = [];
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      // Pula linhas sem ano ou sem mês de operação
      if (row[0] == null || row[1] == null) continue;
      rows.push({
        ano: row[0],
        mesOperacao: row[1],
        mesRecolhimento: row[2],
        valorDevido: row[3],
        valorPago: row[4],
        obs: row[5],
      });
    }
    result.irpf = rows;
  }
  return result;
}

// ── Constants ─────────────────────────────────────────────────
const TABS = [
  { id: "overview", label: "Visão geral", icon: "ti-layout-dashboard" },
  { id: "positions", label: "Posições abertas", icon: "ti-table" },
  { id: "pl", label: "P&L", icon: "ti-chart-bar" },
  { id: "irpf", label: "IRPF", icon: "ti-receipt-tax" },
];
const C_POS = "#1D9E75", C_NEG = "#D85A30", C_WARN = "#C8881A";

// Cor para Mkt vs Strike: ITM (<0) vermelho, 0-10% amarelo, >10% normal
function mktColor(val) {
  if (isNA(val) || isNaN(parseFloat(val))) return "var(--color-text-secondary)";
  const v = parseFloat(val);
  if (v < 0) return C_NEG;
  if (v <= 0.10) return C_WARN;
  return "var(--color-text-primary)";
}
const th = (a) => ({ padding: "9px 12px", textAlign: a, fontSize: 11, fontWeight: 500, color: "var(--color-text-secondary)", borderBottom: "0.5px solid var(--color-border-tertiary)", whiteSpace: "nowrap", background: "var(--color-background-secondary)" });
const td = (a) => ({ padding: "8px 12px", textAlign: a, color: "var(--color-text-primary)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" });

// ── App ───────────────────────────────────────────────────────
export default function App() {
  const [authState, setAuthState] = useState("idle");
  const [data, setData] = useState(null);
  const [lastUpdated, setLastUpdated] = useState(null);
  const [tab, setTab] = useState("overview");
  const [loadingData, setLoadingData] = useState(false);
  const [error, setError] = useState(null);

  useEffect(() => {
    msalInstance.initialize().then(() => {
      msalInstance.handleRedirectPromise().then((response) => {
        if (response) {
          setAuthState("authenticated");
          loadFromOneDrive();
        } else {
          const accounts = msalInstance.getAllAccounts();
          if (accounts.length > 0) {
            setAuthState("authenticated");
            loadFromOneDrive();
          } else {
            setAuthState("idle");
          }
        }
      }).catch(() => setAuthState("idle"));
    });
  }, []);

  const login = async () => {
    setAuthState("loading");
    try {
      await msalInstance.initialize();
      await msalInstance.loginRedirect(loginRequest);
    } catch (e) {
      setError("Erro no login: " + e.message);
      setAuthState("error");
    }
  };

  const logout = async () => {
    await msalInstance.initialize();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) await msalInstance.logoutRedirect({ account: accounts[0] });
  };

  const loadFromOneDrive = async () => {
    setLoadingData(true);
    setError(null);
    try {
      const buffer = await fetchPlanilha();
      const wb = XLSX.read(buffer, { type: "array", cellDates: false });
      setData(parseWorkbook(wb));
      setLastUpdated(new Date());
    } catch (e) {
      setError(e.message);
    } finally {
      setLoadingData(false);
    }
  };

  if (authState === "idle" || authState === "error") return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: "2rem" }}>
      <div style={{ maxWidth: 400, width: "100%", textAlign: "center" }}>
        <div style={{ width: 56, height: 56, borderRadius: "var(--border-radius-lg)", background: "var(--color-background-secondary)", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 1.5rem" }}>
          <i className="ti ti-chart-candle" aria-hidden style={{ fontSize: 28, color: C_POS }} />
        </div>
        <h1 style={{ fontSize: 22, fontWeight: 500, margin: "0 0 8px" }}>Gestão de Opções</h1>
        <p style={{ fontSize: 14, color: "var(--color-text-secondary)", margin: "0 0 2rem" }}>
          Conecte sua conta Microsoft para carregar a planilha automaticamente do OneDrive
        </p>
        {error && (
          <div style={{ background: "var(--color-background-danger)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-md)", padding: "10px 14px", marginBottom: "1rem", fontSize: 13, color: "var(--color-text-danger)", textAlign: "left" }}>
            <i className="ti ti-alert-circle" aria-hidden style={{ marginRight: 6 }} />{error}
          </div>
        )}
        <button onClick={login} style={{ width: "100%", padding: "10px 20px", fontSize: 14, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
          <i className="ti ti-brand-windows" aria-hidden style={{ fontSize: 18 }} />
          Entrar com conta Microsoft
        </button>
      </div>
    </div>
  );

  if (authState === "loading" || (authState === "authenticated" && !data && loadingData)) return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ textAlign: "center" }}>
        <i className="ti ti-loader-2" aria-hidden style={{ fontSize: 32, color: "var(--color-text-secondary)", display: "block", marginBottom: 12, animation: "spin 1s linear infinite" }} />
        <p style={{ fontSize: 14, color: "var(--color-text-secondary)" }}>Carregando planilha do OneDrive…</p>
      </div>
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    </div>
  );

  if (authState === "authenticated" && !data && !loadingData) return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: "2rem" }}>
      <div style={{ maxWidth: 400, width: "100%", textAlign: "center" }}>
        <i className="ti ti-file-alert" aria-hidden style={{ fontSize: 36, color: C_NEG, display: "block", marginBottom: 12 }} />
        <h2 style={{ fontSize: 18, fontWeight: 500, margin: "0 0 8px" }}>Erro ao carregar planilha</h2>
        <p style={{ fontSize: 13, color: "var(--color-text-secondary)", margin: "0 0 1.5rem" }}>{error}</p>
        <div style={{ display: "flex", gap: 8, justifyContent: "center" }}>
          <button onClick={loadFromOneDrive} style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <i className="ti ti-refresh" aria-hidden /> Tentar novamente
          </button>
          <button onClick={logout} style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <i className="ti ti-logout" aria-hidden /> Sair
          </button>
        </div>
      </div>
    </div>
  );

  if (!data) return null;

  const { plAnual, positionsOpen, irpf, dre } = data;
  const totalAnual = plAnual.reduce((s, m) => s + (parseFloat(m.resultado) || 0), 0);
  const ytd = calcYTD(plAnual);
  const melhorMes = plAnual.reduce((b, m) => !b || parseFloat(m.resultado) > parseFloat(b.resultado) ? m : b, null);
  const piorMes = plAnual.reduce((b, m) => !b || parseFloat(m.resultado) < parseFloat(b.resultado) ? m : b, null);
  const mesesPos = plAnual.filter(m => parseFloat(m.resultado) > 0).length;

  // Somatórios posições abertas
  const mtmValidas = positionsOpen.filter(p => !isNA(p.mtm) && !isNaN(parseFloat(p.mtm)));
  const plValidas = positionsOpen.filter(p => !isNA(p.plTotal) && !isNaN(parseFloat(p.plTotal)));
  const totalMtM = mtmValidas.reduce((s, p) => s + parseFloat(p.mtm), 0);
  const totalPL = plValidas.reduce((s, p) => s + parseFloat(p.plTotal), 0);

  return (
    <div style={{ padding: "1.5rem 0" }}>
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>

      {/* Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "1.25rem" }}>
        <div>
          <h2 style={{ fontSize: 20, fontWeight: 500, margin: "0 0 3px" }}>Gestão de Opções</h2>
          <p style={{ fontSize: 12, color: "var(--color-text-secondary)", margin: 0 }}>
            <i className="ti ti-cloud" aria-hidden style={{ marginRight: 4 }} />
            {lastUpdated ? `Atualizado ${lastUpdated.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}` : "OneDrive"}
          </p>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={loadFromOneDrive} disabled={loadingData} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12 }}>
            <i className={`ti ${loadingData ? "ti-loader-2" : "ti-refresh"}`} aria-hidden style={loadingData ? { animation: "spin 1s linear infinite" } : {}} />
            {loadingData ? "Atualizando…" : "Atualizar"}
          </button>
          <button onClick={logout} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12 }}>
            <i className="ti ti-logout" aria-hidden /> Sair
          </button>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ display: "flex", gap: 2, marginBottom: "1.25rem", borderBottom: "0.5px solid var(--color-border-tertiary)" }}>
        {TABS.map(t => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{ fontSize: 13, padding: "8px 12px", border: "none", borderBottom: tab === t.id ? "2px solid var(--color-text-primary)" : "2px solid transparent", borderRadius: 0, background: "transparent", color: tab === t.id ? "var(--color-text-primary)" : "var(--color-text-secondary)", fontWeight: tab === t.id ? 500 : 400, cursor: "pointer", display: "flex", alignItems: "center", gap: 5 }}>
            <i className={`ti ${t.icon}`} aria-hidden style={{ fontSize: 14 }} />{t.label}
          </button>
        ))}
      </div>

      {/* ── Visão geral ── */}
      {tab === "overview" && (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginBottom: "1.25rem" }}>
            {[
              { label: "resultado total", value: fmtR(totalAnual), color: totalAnual >= 0 ? C_POS : C_NEG },
              { label: "meses registrados", value: plAnual.length },
              { label: "meses positivos", value: `${mesesPos} / ${plAnual.length}` },
              { label: "melhor mês", value: melhorMes ? `${melhorMes.mes}/${melhorMes.ano} (${fmtR(melhorMes.resultado, 0)})` : "—", color: C_POS },
              { label: "pior mês", value: piorMes ? `${piorMes.mes}/${piorMes.ano} (${fmtR(piorMes.resultado, 0)})` : "—", color: parseFloat(piorMes?.resultado) < 0 ? C_NEG : null },
              { label: "posições abertas", value: positionsOpen.length },
            ].map((c, i) => (
              <div key={i} style={{ background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", padding: "14px" }}>
                <p style={{ fontSize: 11, color: "var(--color-text-secondary)", margin: "0 0 5px" }}>{c.label}</p>
                <p style={{ fontSize: 14, fontWeight: 500, margin: 0, color: c.color || "var(--color-text-primary)", wordBreak: "break-word", lineHeight: 1.4 }}>{c.value}</p>
              </div>
            ))}
          </div>

          {dre && (
            <div style={{ marginBottom: "1.25rem" }}>
              <h3 style={{ fontSize: 15, fontWeight: 500, margin: "0 0 10px" }}>DRE — mês corrente</h3>
              <div style={{ border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", overflow: "hidden" }}>
                {[
                  { label: "Receita (prêmios vendidos)", value: dre.receita },
                  { label: "Recompras", value: dre.recompras },
                  { label: "MtM posições abertas", value: dre.mtmTotal, muted: !dre.mtmDisponivel, hint: !dre.mtmDisponivel ? "salve com o Profit aberto" : null },
                  { label: "Resultado sem MtM", value: dre.resultadoSemMtM, hi: true },
                  { label: "Resultado com MtM", value: dre.mtmDisponivel ? dre.resultado : null, hi: true, muted: !dre.mtmDisponivel },
                ].map((row, i) => {
                  const v = parseFloat(row.value);
                  const color = isNaN(v) ? "var(--color-text-secondary)" : v >= 0 ? C_POS : C_NEG;
                  return (
                    <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "11px 16px", gap: 12, borderTop: i > 0 ? "0.5px solid var(--color-border-tertiary)" : "none", background: row.hi ? "var(--color-background-secondary)" : "transparent" }}>
                      <span style={{ fontSize: 13, color: row.muted || !row.hi ? "var(--color-text-secondary)" : "var(--color-text-primary)", fontWeight: row.hi ? 500 : 400 }}>
                        {row.label}{row.hint && <span style={{ fontSize: 11, marginLeft: 6 }}>({row.hint})</span>}
                      </span>
                      <span style={{ fontSize: 13, fontWeight: row.hi ? 500 : 400, whiteSpace: "nowrap", color: row.muted ? "var(--color-text-secondary)" : row.hi ? color : "var(--color-text-primary)" }}>
                        {(row.muted && row.value == null) || isNaN(v) ? "—" : fmtR(row.value)}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          <h3 style={{ fontSize: 15, fontWeight: 500, margin: "0 0 10px" }}>Histórico mensal</h3>
          <MiniChart data={plAnual} />
        </div>
      )}

      {/* ── Posições abertas ── */}
      {tab === "positions" && (
        <div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <p style={{ margin: 0, fontSize: 13, color: "var(--color-text-secondary)" }}>
              {positionsOpen.length} posições short put em aberto
              <span style={{ marginLeft: 12 }}><i className="ti ti-info-circle" aria-hidden style={{ marginRight: 3 }} />MtM disponível quando Profit está ativo</span>
            </p>
          </div>

          {/* Somatórios */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
            <div style={{ background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", padding: "12px 16px" }}>
              <p style={{ fontSize: 11, color: "var(--color-text-secondary)", margin: "0 0 4px" }}>
                MtM total {mtmValidas.length < positionsOpen.length ? `(${mtmValidas.length}/${positionsOpen.length} posições)` : ""}
              </p>
              <p style={{ fontSize: 15, fontWeight: 500, margin: 0, color: mtmValidas.length === 0 ? "var(--color-text-secondary)" : totalMtM >= 0 ? C_POS : C_NEG }}>
                {mtmValidas.length === 0 ? "—" : fmtR(totalMtM)}
              </p>
            </div>
            <div style={{ background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", padding: "12px 16px" }}>
              <p style={{ fontSize: 11, color: "var(--color-text-secondary)", margin: "0 0 4px" }}>
                P&L total {plValidas.length < positionsOpen.length ? `(${plValidas.length}/${positionsOpen.length} posições)` : ""}
              </p>
              <p style={{ fontSize: 15, fontWeight: 500, margin: 0, color: plValidas.length === 0 ? "var(--color-text-secondary)" : totalPL >= 0 ? C_POS : C_NEG }}>
                {plValidas.length === 0 ? "—" : fmtR(totalPL)}
              </p>
            </div>
          </div>

          {/* Alerta de proximidade do strike */}
          <ProximityAlert positions={positionsOpen} threshold={0.10} />

          {/* Consolidado por empresa */}
          <h3 style={{ fontSize: 15, fontWeight: 500, margin: "4px 0 10px" }}>Consolidado por ativo</h3>
          <div style={{ marginBottom: "1.5rem" }}>
            <ByCompanyTable positions={positionsOpen} />
          </div>

          <h3 style={{ fontSize: 15, fontWeight: 500, margin: "4px 0 10px" }}>Todas as posições</h3>
          <PositionsTable positions={positionsOpen} />
        </div>
      )}

      {/* ── P&L ── */}
      {tab === "pl" && (
        <div>
          {/* YTD dashboard */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginBottom: "1.25rem" }}>
            {[
              { label: `resultado YTD (${new Date().getFullYear()})`, value: fmtR(ytd), color: ytd >= 0 ? C_POS : C_NEG },
              { label: "resultado total (todos os anos)", value: fmtR(totalAnual), color: totalAnual >= 0 ? C_POS : C_NEG },
              { label: "meses positivos", value: `${mesesPos} / ${plAnual.length}` },
            ].map((c, i) => (
              <div key={i} style={{ background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", padding: "14px" }}>
                <p style={{ fontSize: 11, color: "var(--color-text-secondary)", margin: "0 0 5px" }}>{c.label}</p>
                <p style={{ fontSize: 15, fontWeight: 500, margin: 0, color: c.color || "var(--color-text-primary)" }}>{c.value}</p>
              </div>
            ))}
          </div>
          <FullChart data={plAnual} />
          <div style={{ marginTop: "1rem" }}><PLTable data={plAnual} /></div>
        </div>
      )}

      {/* ── IRPF ── */}
      {tab === "irpf" && <IRPFTable data={irpf} />}
    </div>
  );
}

function MiniChart({ data }) {
  if (!data?.length) return null;
  const chartData = data.map(m => ({ name: `${m.mes.substring(0, 3)}/${String(m.ano).substring(2)}`, value: parseFloat(m.resultado) || 0 }));
  return (
    <div style={{ height: 150 }}>
      <ResponsiveContainer width="100%" height="100%">
        <BarChart data={chartData} margin={{ top: 4, right: 4, bottom: 0, left: 0 }}>
          <XAxis dataKey="name" tick={{ fontSize: 10, fill: "var(--color-text-secondary)" }} axisLine={false} tickLine={false} />
          <YAxis hide />
          <Tooltip formatter={(v) => ["R$ " + v.toLocaleString("pt-BR", { minimumFractionDigits: 2 }), "Resultado"]} contentStyle={{ fontSize: 12, borderRadius: 6 }} />
          <ReferenceLine y={0} stroke="var(--color-border-secondary)" strokeWidth={0.5} />
          <Bar dataKey="value" radius={[3, 3, 0, 0]}>
            {chartData.map((e, i) => <Cell key={i} fill={e.value >= 0 ? C_POS : C_NEG} />)}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

function FullChart({ data }) {
  if (!data?.length) return null;
  const chartData = data.map(m => ({ name: `${m.mes.substring(0, 3)}/${String(m.ano).substring(2)}`, value: parseFloat(m.resultado) || 0 }));
  return (
    <div style={{ height: 220 }}>
      <ResponsiveContainer width="100%" height="100%">
        <BarChart data={chartData} margin={{ top: 4, right: 8, bottom: 0, left: 8 }}>
          <XAxis dataKey="name" tick={{ fontSize: 10, fill: "var(--color-text-secondary)" }} axisLine={false} tickLine={false} />
          <YAxis tick={{ fontSize: 11, fill: "var(--color-text-secondary)" }} axisLine={false} tickLine={false} tickFormatter={v => "R$" + (Math.abs(v) >= 1000 ? (v / 1000).toFixed(0) + "k" : v)} />
          <Tooltip formatter={(v) => ["R$ " + v.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }), "Resultado"]} contentStyle={{ fontSize: 12, borderRadius: 6 }} />
          <ReferenceLine y={0} stroke="var(--color-border-primary)" strokeWidth={0.5} />
          <Bar dataKey="value" radius={[3, 3, 0, 0]}>
            {chartData.map((e, i) => <Cell key={i} fill={e.value >= 0 ? C_POS : C_NEG} />)}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

function PLTable({ data }) {
  const total = data.reduce((s, m) => s + (parseFloat(m.resultado) || 0), 0);
  return (
    <div style={{ border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", overflow: "hidden" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
        <thead><tr><th style={th("left")}>Mês</th><th style={th("left")}>Ano</th><th style={th("right")}>Resultado</th></tr></thead>
        <tbody>
          {data.map((m, i) => {
            const v = parseFloat(m.resultado);
            return (
              <tr key={i} style={{ borderTop: "0.5px solid var(--color-border-tertiary)" }}>
                <td style={td("left")}>{m.mes}</td>
                <td style={td("left")}>{m.ano}</td>
                <td style={{ ...td("right"), color: v >= 0 ? C_POS : C_NEG, fontWeight: 500 }}>{fmtR(m.resultado)}</td>
              </tr>
            );
          })}
        </tbody>
        <tfoot>
          <tr style={{ borderTop: "0.5px solid var(--color-border-primary)", background: "var(--color-background-secondary)" }}>
            <td colSpan={2} style={{ ...td("left"), fontWeight: 500 }}>Total</td>
            <td style={{ ...td("right"), fontWeight: 500, color: total >= 0 ? C_POS : C_NEG }}>{fmtR(total)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
  );
}

function ProximityAlert({ positions, threshold }) {
  // Filtra posições com Mkt vs Strike disponível e dentro do threshold
  const comDados = positions.filter(p => !isNA(p.mktVsStrike) && !isNaN(parseFloat(p.mktVsStrike)));

  // Se nenhum dado disponível (Profit fechado), mostra aviso
  if (comDados.length === 0) {
    return (
      <div style={{ background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", padding: "12px 16px", marginBottom: "1.5rem", fontSize: 13, color: "var(--color-text-secondary)" }}>
        <i className="ti ti-alert-triangle" aria-hidden style={{ marginRight: 6 }} />
        Alerta de proximidade do strike indisponível — salve a planilha com o Profit aberto para calcular a distância Mkt vs Strike.
      </div>
    );
  }

  const proximas = comDados
    .filter(p => parseFloat(p.mktVsStrike) <= threshold)
    .sort((a, b) => parseFloat(a.mktVsStrike) - parseFloat(b.mktVsStrike));

  return (
    <div style={{ marginBottom: "1.5rem" }}>
      <h3 style={{ fontSize: 15, fontWeight: 500, margin: "4px 0 10px", display: "flex", alignItems: "center", gap: 6 }}>
        <i className="ti ti-alert-triangle" aria-hidden style={{ color: proximas.length > 0 ? C_NEG : "var(--color-text-secondary)" }} />
        Proximidade do strike — Mkt vs Strike ≤ {(threshold * 100).toFixed(0)}%
      </h3>
      {proximas.length === 0 ? (
        <div style={{ background: "var(--color-background-success)", borderRadius: "var(--border-radius-md)", padding: "12px 16px", fontSize: 13, color: "var(--color-text-success)" }}>
          <i className="ti ti-circle-check" aria-hidden style={{ marginRight: 6 }} />
          Nenhuma posição com Mkt vs Strike de {(threshold * 100).toFixed(0)}% ou menos. Todas com folga.
        </div>
      ) : (
        <div style={{ border: `0.5px solid ${C_NEG}`, borderRadius: "var(--border-radius-lg)", overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 480 }}>
            <thead>
              <tr>
                <th style={th("left")}>Opção</th>
                <th style={th("left")}>Ação</th>
                <th style={th("right")}>Strike</th>
                <th style={th("right")}>Mkt vs Strike</th>
                <th style={th("right")}>Vencimento</th>
              </tr>
            </thead>
            <tbody>
              {proximas.map((p, i) => {
                const pct = parseFloat(p.mktVsStrike);
                return (
                  <tr key={i} style={{ borderTop: "0.5px solid var(--color-border-tertiary)" }}>
                    <td style={{ ...td("left"), fontFamily: "var(--font-mono)", fontSize: 11 }}>{p.ticker}</td>
                    <td style={td("left")}>{p.acao}</td>
                    <td style={td("right")}>R$ {fmt(p.strike, 2)}</td>
                    <td style={{ ...td("right"), fontWeight: 500, color: mktColor(pct) }}>{fmtPct(pct)}</td>
                    <td style={td("right")}>{fmtDate(p.vencimento)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function ByCompanyTable({ positions }) {
  const [sort, setSort] = useState({ col: "notional", dir: 1 });

  // Agrupa por ativo (acao)
  const grupos = {};
  for (const p of positions) {
    const key = p.acao || "—";
    if (!grupos[key]) grupos[key] = { acao: key, nPos: 0, totalQtd: 0, notional: 0, totalMtM: 0, totalPL: 0, mtmCount: 0, plCount: 0 };
    const g = grupos[key];
    g.nPos += 1;
    const q = parseFloat(p.qtd); if (!isNaN(q)) g.totalQtd += q;
    const strike = parseFloat(p.strike);
    if (!isNaN(q) && !isNaN(strike)) g.notional += q * strike;
    if (!isNA(p.mtm) && !isNaN(parseFloat(p.mtm))) { g.totalMtM += parseFloat(p.mtm); g.mtmCount += 1; }
    if (!isNA(p.plTotal) && !isNaN(parseFloat(p.plTotal))) { g.totalPL += parseFloat(p.plTotal); g.plCount += 1; }
  }
  const rows = Object.values(grupos);

  const sorted = [...rows].sort((a, b) => {
    let va = a[sort.col], vb = b[sort.col];
    if (sort.col === "acao") return va > vb ? sort.dir : va < vb ? -sort.dir : 0;
    return (va - vb) * sort.dir;
  });

  const toggleSort = (col) => setSort(s => ({ col, dir: s.col === col ? -s.dir : 1 }));
  const SI = ({ col }) => <i className={`ti ${sort.col === col ? (sort.dir > 0 ? "ti-arrow-up" : "ti-arrow-down") : "ti-arrows-sort"}`} aria-hidden style={{ fontSize: 10, marginLeft: 3, opacity: sort.col === col ? 1 : 0.3 }} />;

  const cols = [
    { k: "acao", l: "Ativo", a: "left" },
    { k: "nPos", l: "Nº posições", a: "right" },
    { k: "totalQtd", l: "Qtd total", a: "right" },
    { k: "notional", l: "Notional (R$)", a: "right" },
    { k: "totalMtM", l: "MtM total (R$)", a: "right" },
    { k: "totalPL", l: "P&L total (R$)", a: "right" },
  ];

  // Totais gerais
  const tMtM = rows.reduce((s, g) => s + g.totalMtM, 0);
  const tPL = rows.reduce((s, g) => s + g.totalPL, 0);
  const tQtd = rows.reduce((s, g) => s + g.totalQtd, 0);
  const tNotional = rows.reduce((s, g) => s + g.notional, 0);
  const tPos = rows.reduce((s, g) => s + g.nPos, 0);
  const anyMtM = rows.some(g => g.mtmCount > 0);
  const anyPL = rows.some(g => g.plCount > 0);

  return (
    <div style={{ border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 520 }}>
        <thead>
          <tr>{cols.map(c => <th key={c.k} onClick={() => toggleSort(c.k)} style={{ ...th(c.a), cursor: "pointer", userSelect: "none" }}>{c.l}<SI col={c.k} /></th>)}</tr>
        </thead>
        <tbody>
          {sorted.map((g, i) => (
            <tr key={i} style={{ borderTop: "0.5px solid var(--color-border-tertiary)" }}>
              <td style={{ ...td("left"), fontWeight: 500 }}>{g.acao}</td>
              <td style={td("right")}>{g.nPos}</td>
              <td style={td("right")}>{fmt(g.totalQtd, 0)}</td>
              <td style={{ ...td("right"), fontWeight: 500 }}>{g.notional > 0 ? fmtR(g.notional, 0) : "—"}</td>
              <td style={{ ...td("right"), color: g.mtmCount === 0 ? "var(--color-text-secondary)" : g.totalMtM >= 0 ? C_POS : C_NEG }}>
                {g.mtmCount === 0 ? "—" : fmtR(g.totalMtM)}
                {g.mtmCount > 0 && g.mtmCount < g.nPos && <span style={{ fontSize: 10, color: "var(--color-text-secondary)", marginLeft: 4 }}>({g.mtmCount}/{g.nPos})</span>}
              </td>
              <td style={{ ...td("right"), color: g.plCount === 0 ? "var(--color-text-secondary)" : g.totalPL >= 0 ? C_POS : C_NEG }}>
                {g.plCount === 0 ? "—" : fmtR(g.totalPL)}
                {g.plCount > 0 && g.plCount < g.nPos && <span style={{ fontSize: 10, color: "var(--color-text-secondary)", marginLeft: 4 }}>({g.plCount}/{g.nPos})</span>}
              </td>
            </tr>
          ))}
        </tbody>
        <tfoot>
          <tr style={{ borderTop: "0.5px solid var(--color-border-primary)", background: "var(--color-background-secondary)" }}>
            <td style={{ ...td("left"), fontWeight: 500 }}>Total</td>
            <td style={{ ...td("right"), fontWeight: 500 }}>{tPos}</td>
            <td style={{ ...td("right"), fontWeight: 500 }}>{fmt(tQtd, 0)}</td>
            <td style={{ ...td("right"), fontWeight: 500 }}>{tNotional > 0 ? fmtR(tNotional, 0) : "—"}</td>
            <td style={{ ...td("right"), fontWeight: 500, color: !anyMtM ? "var(--color-text-secondary)" : tMtM >= 0 ? C_POS : C_NEG }}>{!anyMtM ? "—" : fmtR(tMtM)}</td>
            <td style={{ ...td("right"), fontWeight: 500, color: !anyPL ? "var(--color-text-secondary)" : tPL >= 0 ? C_POS : C_NEG }}>{!anyPL ? "—" : fmtR(tPL)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
  );
}

function PositionsTable({ positions }) {
  const [sort, setSort] = useState({ col: "vencimento", dir: 1 });
  const sorted = [...positions].sort((a, b) => {
    let va = a[sort.col], vb = b[sort.col];
    if (sort.col === "vencimento") { va = parseDate(va)?.getTime() || 0; vb = parseDate(vb)?.getTime() || 0; }
    else if (!["ticker", "acao"].includes(sort.col)) { va = parseFloat(va) || 0; vb = parseFloat(vb) || 0; }
    return va > vb ? sort.dir : va < vb ? -sort.dir : 0;
  });
  const toggleSort = (col) => setSort(s => ({ col, dir: s.col === col ? -s.dir : 1 }));
  const SI = ({ col }) => <i className={`ti ${sort.col === col ? (sort.dir > 0 ? "ti-arrow-up" : "ti-arrow-down") : "ti-arrows-sort"}`} aria-hidden style={{ fontSize: 10, marginLeft: 3, opacity: sort.col === col ? 1 : 0.3 }} />;
  const cols = [
    { k: "ticker", l: "Opção", a: "left" }, { k: "acao", l: "Ação", a: "left" },
    { k: "qtd", l: "Qtd", a: "right" }, { k: "precoVenda", l: "Px venda", a: "right" },
    { k: "strike", l: "Strike", a: "right" }, { k: "mktVsStrike", l: "Mkt vs Strike", a: "right" },
    { k: "vencimento", l: "Vencimento", a: "right" },
    { k: "mtm", l: "MtM (R$)", a: "right" }, { k: "plTotal", l: "P&L (R$)", a: "right" },
    { k: "plPct", l: "P&L (%)", a: "right" },
  ];
  return (
    <div style={{ border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 640 }}>
        <thead>
          <tr>{cols.map(c => <th key={c.k} onClick={() => toggleSort(c.k)} style={{ ...th(c.a), cursor: "pointer", userSelect: "none" }}>{c.l}<SI col={c.k} /></th>)}</tr>
        </thead>
        <tbody>
          {sorted.map((p, i) => {
            const mtmN = parseFloat(p.mtm), plN = parseFloat(p.plTotal), plPN = parseFloat(p.plPct);
            return (
              <tr key={i} style={{ borderTop: "0.5px solid var(--color-border-tertiary)" }}>
                <td style={{ ...td("left"), fontFamily: "var(--font-mono)", fontSize: 11 }}>{p.ticker}</td>
                <td style={td("left")}>{p.acao}</td>
                <td style={td("right")}>{fmt(p.qtd, 0)}</td>
                <td style={td("right")}>R$ {fmt(p.precoVenda, 2)}</td>
                <td style={td("right")}>R$ {fmt(p.strike, 2)}</td>
                <td style={{ ...td("right"), color: mktColor(p.mktVsStrike), fontWeight: !isNA(p.mktVsStrike) && parseFloat(p.mktVsStrike) <= 0.10 ? 500 : 400 }}>{isNA(p.mktVsStrike) ? "—" : fmtPct(p.mktVsStrike)}</td>
                <td style={td("right")}>{fmtDate(p.vencimento)}</td>
                <td style={{ ...td("right"), color: isNA(p.mtm) ? "var(--color-text-secondary)" : mtmN >= 0 ? C_POS : C_NEG }}>{isNA(p.mtm) ? "—" : fmtR(p.mtm)}</td>
                <td style={{ ...td("right"), color: isNA(p.plTotal) ? "var(--color-text-secondary)" : plN >= 0 ? C_POS : C_NEG }}>{isNA(p.plTotal) ? "—" : fmtR(p.plTotal)}</td>
                <td style={{ ...td("right"), color: isNA(p.plPct) ? "var(--color-text-secondary)" : plPN >= 0 ? C_POS : C_NEG }}>{isNA(p.plPct) ? "—" : fmtPct(p.plPct)}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function IRPFTable({ data }) {
  if (!data?.length) return <p style={{ color: "var(--color-text-secondary)", fontSize: 14 }}>Nenhum dado de IRPF encontrado.</p>;

  const totalDevido = data.reduce((s, r) => s + (parseFloat(r.valorDevido) || 0), 0);
  const totalPago = data.reduce((s, r) => s + (parseFloat(r.valorPago) || 0), 0);

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginBottom: "1.25rem" }}>
        {[
          { label: "total devido (principal)", value: fmtR(totalDevido) },
          { label: "total pago", value: fmtR(totalPago) },
          { label: "meses registrados", value: data.length },
        ].map((c, i) => (
          <div key={i} style={{ background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", padding: "14px" }}>
            <p style={{ fontSize: 11, color: "var(--color-text-secondary)", margin: "0 0 5px" }}>{c.label}</p>
            <p style={{ fontSize: 15, fontWeight: 500, margin: 0 }}>{c.value}</p>
          </div>
        ))}
      </div>

      <div style={{ border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 680 }}>
          <thead>
            <tr>
              <th style={th("left")}>Ano ref.</th>
              <th style={th("left")}>Mês operação</th>
              <th style={th("left")}>Mês recolhimento</th>
              <th style={th("right")}>Valor devido (MyCapital)</th>
              <th style={th("right")}>Valor pago</th>
              <th style={th("left")}>Obs.</th>
            </tr>
          </thead>
          <tbody>
            {data.map((row, i) => (
              <tr key={i} style={{ borderTop: "0.5px solid var(--color-border-tertiary)" }}>
                <td style={td("left")}>{row.ano}</td>
                <td style={td("left")}>{row.mesOperacao}</td>
                <td style={td("left")}>{row.mesRecolhimento || "—"}</td>
                <td style={td("right")}>{fmtR(row.valorDevido)}</td>
                <td style={td("right")}>{fmtR(row.valorPago)}</td>
                <td style={{ ...td("left"), fontSize: 12, color: "var(--color-text-secondary)", whiteSpace: "normal" }}>{row.obs || "—"}</td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr style={{ borderTop: "0.5px solid var(--color-border-primary)", background: "var(--color-background-secondary)" }}>
              <td colSpan={3} style={{ ...td("left"), fontWeight: 500 }}>Total</td>
              <td style={{ ...td("right"), fontWeight: 500 }}>{fmtR(totalDevido)}</td>
              <td style={{ ...td("right"), fontWeight: 500 }}>{fmtR(totalPago)}</td>
              <td style={td("left")}></td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}
