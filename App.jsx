import { useState, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { AlertTriangle, CheckCircle2, Download, RotateCcw, Zap, Upload } from "lucide-react";

const INK = "#05070D";
const PANEL = "#0B0F1A";
const BLUE = "#2E6BFF";
const BLUE_DIM = "#173A8A";
const LINE = "#1E2536";
const TEXT_DIM = "#8592B0";
const ALERT = "#FF5D5D";
const OK = "#3ED598";

const NCM_PATTERN = /(\d{4}\.?\d{2}\.?\d{2})/;

const FIELD_OPTIONS = [
  { key: "codigo", label: "Código" },
  { key: "descricao", label: "Descrição" },
  { key: "ncm", label: "NCM" },
  { key: "cfop", label: "CST / CFOP / CSOSN" },
  { key: "ignore", label: "Ignorar coluna" },
];

// Palavras que costumam aparecer só na linha de cabeçalho, nunca em uma
// linha de produto de verdade — usadas para decidir sozinho se a primeira
// linha colada é cabeçalho, sem precisar perguntar ao usuário.
const HEADER_KEYWORDS = [
  "codigo", "cod", "descricao", "desc", "ncm", "cst", "cfop", "csosn",
  "produto", "item", "mercadoria", "unidade", "un", "quantidade",
];

function detectDelimiter(line) {
  if (line.includes("\t")) return "\t";
  if (line.includes(";")) return ";";
  if (line.includes(",")) return ",";
  // Fixed-width style paste: columns padded with 2+ spaces, single spaces
  // still used inside descriptions. Treat runs of 2+ spaces as the column break.
  if (/\s{2,}/.test(line)) return /\s{2,}/;
  // No reliable column separator (only single spaces, which also show up
  // inside descriptions like "Coca Cola"). Returning null here means the
  // caller falls back to pattern-based extraction instead of a naive split,
  // so a multi-word description never steals the NCM's column.
  return null;
}

// Normalize NCM robustly: strip everything but digits, then restore the
// leading zero Excel silently drops on codes like 0109.10.00.
function normalizeNCM(v) {
  const digits = (v || "").replace(/\D/g, "");
  if (!digits) return "";
  return digits.padStart(8, "0");
}

function splitLines(text) {
  return text.trim() ? text.trim().split("\n") : [];
}

function normalizeHeaderCell(s) {
  return (s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

// Decide sozinho se a primeira linha é cabeçalho, olhando se alguma célula
// bate com um termo típico de coluna (NCM, código, descrição...).
function looksLikeHeaderRow(row) {
  if (!row || row.length === 0) return false;
  return row.some((cell) => {
    const c = normalizeHeaderCell(cell);
    if (!c) return false;
    return HEADER_KEYWORDS.some((k) => c === k || c.includes(k));
  });
}

// Sem delimitador confiável: em vez de dividir por espaço (o que quebra
// descrições com mais de uma palavra, tipo "Coca Cola"), localiza o NCM
// pelo próprio padrão numérico (8 dígitos, com ou sem pontos) e usa isso
// como âncora — o que sobra antes dele vira código (se parecer um código
// curto) + descrição, sem nunca confundir uma palavra da descrição com o NCM.
function smartSplitLine(line) {
  const trimmed = line.trim();
  const m = trimmed.match(NCM_PATTERN);
  let ncm = "";
  let rest = trimmed;
  if (m) {
    ncm = m[1];
    rest = (trimmed.slice(0, m.index) + trimmed.slice(m.index + m[0].length))
      .replace(/\s{2,}/g, " ")
      .trim();
  }
  let codigo = "";
  const codeMatch = rest.match(/^(\S{1,12})\s+(?=\S)/);
  // Só trata o primeiro token como "código" separado se ele parecer mesmo
  // um código (tem dígito, é curto) — uma palavra comum da descrição
  // (como "Coca") não tem dígito e continua fazendo parte da descrição.
  if (codeMatch && /\d/.test(codeMatch[1]) && /^[A-Za-z0-9.\-\/]+$/.test(codeMatch[1])) {
    codigo = codeMatch[1];
    rest = rest.slice(codeMatch[0].length).trim();
  }
  return [codigo, rest, ncm, ""];
}

// Products: real column splitting when a delimiter exists (order can vary
// per client, so we map columns manually). When there's no reliable
// delimiter, fall back to smartSplitLine so multi-word descriptions never
// swallow the NCM column.
function parseProducts(text) {
  const lines = splitLines(text);
  if (lines.length === 0) return { rows: [], hasDelimiter: false, riskySpaceSplit: false };
  const delim = detectDelimiter(lines[0]);
  if (delim) {
    return { rows: lines.map((l) => l.split(delim).map((c) => c.trim())), hasDelimiter: true, riskySpaceSplit: false };
  }
  const rows = lines.map(smartSplitLine);
  return { rows, hasDelimiter: false, riskySpaceSplit: false };
}

// NCM base: always just [NCM, mensagem] — fixed, no mapping step needed.
function parseBase(text) {
  const lines = splitLines(text);
  if (lines.length === 0) return [];
  const delim = detectDelimiter(lines[0]);
  if (delim) return lines.map((l) => l.split(delim).map((c) => c.trim()));
  return lines.map((line) => {
    const m = line.match(NCM_PATTERN);
    if (!m) return [line.trim()];
    const rest = line.replace(m[1], "").trim();
    return [m[1], rest];
  });
}

function guessMapping(colCount) {
  const defaults = ["codigo", "descricao", "ncm", "cfop"];
  return Array.from({ length: colCount }, (_, i) => defaults[i] || "ignore");
}

function escapeHtml(s) {
  return (s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Aviso fixo que acompanha os dois relatórios: deixa claro que o diagnóstico
// depende do NCM informado e da base mantida pelo escritório, com a data e a
// fonte usadas, e recomenda a conferência individual antes de qualquer ação.
function buildAvisoLegal(dataBaseNCM, fonteBaseNCM, hoje) {
  const dataLabel = (dataBaseNCM || "").trim() || hoje;
  const fonteLabel = (fonteBaseNCM || "").trim() || "base de NCMs mantida pelo escritório";
  return (
    "Diagnóstico baseado no NCM informado pelo cliente e na base de NCMs com tratamento tributário diferenciado atualizada em " +
    dataLabel + " (fonte: " + fonteLabel + "). Este relatório não substitui a análise individualizada por profissional " +
    "tributário. Recomenda-se a conferência NCM a NCM na IOB (ou base equivalente) antes de qualquer alteração de " +
    "enquadramento fiscal, especialmente nos casos em que o mesmo NCM pode ter tratamentos tributários distintos " +
    "conforme a descrição do produto."
  );
}

// ---------------------------------------------------------------------------
// Login + limite de uso (MVP). Guardado em localStorage — funciona em
// qualquer navegador, mas é POR DISPOSITIVO: se o mesmo cliente acessar de
// outro computador, o contador não acompanha. Suficiente pra validar a
// ideia; pra produção de verdade (conta única multi-dispositivo), o certo é
// migrar pra um backend com banco de dados real.
// ---------------------------------------------------------------------------
const FREE_LIMIT = 4;
// Troque este código pra algo só seu. É o que você manda pro cliente depois
// que ele te paga (Pix, por exemplo), pra ele liberar o próprio acesso.
const UNLOCK_CODE = "LIBERAR2026";

function userKey(email) {
  return `dncm_users:${email.trim().toLowerCase()}`;
}

async function loadUser(email) {
  try {
    const raw = localStorage.getItem(userKey(email));
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

async function saveUser(email, data) {
  try {
    localStorage.setItem(userKey(email), JSON.stringify(data));
  } catch {
    // Falha silenciosa (ex: modo anônimo bloqueando storage) — o contador
    // local (React state) ainda funciona durante a sessão.
  }
}

function AuthShell({ children }) {
  return (
    <div style={{ background: INK, minHeight: "100%", color: "#EAEFFB", fontFamily: "'IBM Plex Sans', 'Segoe UI', sans-serif", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500&display=swap');
        .dncm-ta:focus, .dncm-input:focus { outline: none; border-color: ${BLUE} !important; box-shadow: 0 0 0 3px ${BLUE_DIM}55; }
        .dncm-pw-mask { -webkit-text-security: disc; text-security: disc; }
        .dncm-btn-primary:hover { background: #4A80FF !important; }
        .dncm-btn-primary:disabled { background:${LINE} !important; color:${TEXT_DIM} !important; cursor:not-allowed !important; }
      `}</style>
      <div style={{ width: "100%", maxWidth: 380, border: `1px solid ${LINE}`, borderRadius: 10, padding: 28, background: PANEL }}>
        {children}
      </div>
    </div>
  );
}

function LoginScreen({ email, setEmail, password, setPassword, onSubmit, loading, error }) {
  return (
    <AuthShell>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
        <div style={{ width: 34, height: 34, borderRadius: 8, background: `linear-gradient(135deg, ${BLUE}, ${BLUE_DIM})`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Zap size={17} color="#fff" strokeWidth={2.3} />
        </div>
        <div style={{ fontSize: 16, fontWeight: 700 }}>Diagnóstico NCM</div>
      </div>
      <p style={{ fontSize: 12.5, color: TEXT_DIM, margin: "0 0 18px" }}>
        Entre com e-mail e senha. Se for a primeira vez, sua conta é criada automaticamente.
      </p>
      <div style={{ marginBottom: 10 }}>
        <label style={{ fontSize: 12, color: TEXT_DIM, display: "block", marginBottom: 5 }}>E-mail</label>
        <input
          className="dncm-input"
          type="email"
          autoComplete="off"
          name="dncm-user-field"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          placeholder="voce@escritorio.com.br"
          style={{ width: "100%", boxSizing: "border-box", padding: "9px 12px", borderRadius: 6, border: `1px solid ${LINE}`, background: INK, color: "#EAEFFB", fontSize: 13.5 }}
        />
      </div>
      <div style={{ marginBottom: 16 }}>
        <label style={{ fontSize: 12, color: TEXT_DIM, display: "block", marginBottom: 5 }}>Senha</label>
        <input
          className="dncm-input dncm-pw-mask"
          type="text"
          autoComplete="off"
          autoCorrect="off"
          autoCapitalize="off"
          spellCheck="false"
          name="dncm-pass-field"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && onSubmit()}
          placeholder="sua senha"
          style={{ width: "100%", boxSizing: "border-box", padding: "9px 12px", borderRadius: 6, border: `1px solid ${LINE}`, background: INK, color: "#EAEFFB", fontSize: 13.5 }}
        />
      </div>
      {error && (
        <div style={{ marginBottom: 14, color: ALERT, fontSize: 12.5, display: "flex", gap: 7, alignItems: "center" }}>
          <AlertTriangle size={14} /> {error}
        </div>
      )}
      <button
        className="dncm-btn-primary"
        onClick={onSubmit}
        disabled={loading}
        style={{ width: "100%", padding: "10px 16px", borderRadius: 6, fontSize: 13.5, fontWeight: 700, border: "none", background: BLUE, color: "#fff", cursor: "pointer" }}
      >
        {loading ? "Entrando…" : "Entrar / Criar conta"}
      </button>
      <p style={{ fontSize: 11, color: "#5C6584", marginTop: 14, marginBottom: 0 }}>
        Teste grátis com {FREE_LIMIT} diagnósticos completos por conta.
      </p>
    </AuthShell>
  );
}

function PaywallScreen({ email, onLogout, unlockInput, setUnlockInput, onUnlock, unlockError }) {
  return (
    <AuthShell>
      <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 6 }}>Você usou seus {FREE_LIMIT} diagnósticos grátis</div>
      <p style={{ fontSize: 12.5, color: TEXT_DIM, margin: "0 0 18px" }}>
        Logado como <b style={{ color: "#C6CEE6" }}>{email}</b>. Pra continuar gerando diagnósticos ilimitados, assine o plano completo.
      </p>
      <div style={{ border: `1px solid ${BLUE}`, borderRadius: 8, padding: 16, marginBottom: 18, background: "#0E1730" }}>
        <div style={{ fontSize: 22, fontWeight: 700, color: "#EAEFFB" }}>R$ 79,90<span style={{ fontSize: 13, color: TEXT_DIM, fontWeight: 500 }}> /mês</span></div>
        <div style={{ fontSize: 12, color: TEXT_DIM, marginTop: 4 }}>Diagnósticos ilimitados, Excel e Word completos.</div>
      </div>
      <p style={{ fontSize: 12, color: TEXT_DIM, marginBottom: 8 }}>
        Fale com a gente pra assinar e receber seu código de liberação.
      </p>
      <a
        href="https://wa.me/5518997599504?text=Ol%C3%A1!%20Quero%20assinar%20o%20Diagn%C3%B3stico%20NCM."
        target="_blank"
        rel="noopener noreferrer"
        style={{
          display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
          width: "100%", boxSizing: "border-box", padding: "10px 16px", borderRadius: 6,
          fontSize: 13.5, fontWeight: 700, textDecoration: "none",
          background: "#25D366", color: "#05070D", marginBottom: 18,
        }}
      >
        Falar no WhatsApp para assinar
      </a>
      <div style={{ marginBottom: 10 }}>
        <label style={{ fontSize: 12, color: TEXT_DIM, display: "block", marginBottom: 5 }}>Código de liberação</label>
        <input
          className="dncm-input"
          value={unlockInput}
          onChange={(e) => setUnlockInput(e.target.value)}
          placeholder="Cole o código recebido"
          style={{ width: "100%", boxSizing: "border-box", padding: "9px 12px", borderRadius: 6, border: `1px solid ${LINE}`, background: INK, color: "#EAEFFB", fontSize: 13.5 }}
        />
      </div>
      {unlockError && (
        <div style={{ marginBottom: 14, color: ALERT, fontSize: 12.5, display: "flex", gap: 7, alignItems: "center" }}>
          <AlertTriangle size={14} /> {unlockError}
        </div>
      )}
      <button
        className="dncm-btn-primary"
        onClick={onUnlock}
        style={{ width: "100%", padding: "10px 16px", borderRadius: 6, fontSize: 13.5, fontWeight: 700, border: "none", background: BLUE, color: "#fff", cursor: "pointer", marginBottom: 10 }}
      >
        Liberar acesso
      </button>
      <button
        onClick={onLogout}
        style={{ width: "100%", padding: "8px 16px", borderRadius: 6, fontSize: 12.5, fontWeight: 600, border: `1px solid ${LINE}`, background: "transparent", color: TEXT_DIM, cursor: "pointer" }}
      >
        Sair
      </button>
    </AuthShell>
  );
}

export default function DiagnosticoNCM() {
  // --- Login / limite de uso ---
  const [user, setUser] = useState(null); // { email, password, reportsUsed, plan }
  const [authEmail, setAuthEmail] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authLoading, setAuthLoading] = useState(false);
  const [authError, setAuthError] = useState("");
  const [unlockInput, setUnlockInput] = useState("");
  const [unlockError, setUnlockError] = useState("");

  async function handleAuthSubmit() {
    setAuthError("");
    const email = authEmail.trim().toLowerCase();
    if (!email || !authPassword) {
      setAuthError("Preencha e-mail e senha.");
      return;
    }
    setAuthLoading(true);
    const existing = await loadUser(email);
    if (existing) {
      if (existing.password !== authPassword) {
        setAuthError("Senha incorreta.");
        setAuthLoading(false);
        return;
      }
      setUser({ email, ...existing });
    } else {
      const fresh = { password: authPassword, reportsUsed: 0, plan: "free" };
      await saveUser(email, fresh);
      setUser({ email, ...fresh });
    }
    setAuthLoading(false);
  }

  function handleLogout() {
    setUser(null);
    setAuthEmail("");
    setAuthPassword("");
    setUnlockInput("");
    setUnlockError("");
  }

  async function handleUnlock() {
    setUnlockError("");
    if (unlockInput.trim().toUpperCase() === UNLOCK_CODE) {
      const updated = { ...user, plan: "pago" };
      setUser(updated);
      await saveUser(user.email, { password: updated.password, reportsUsed: updated.reportsUsed, plan: "pago" });
    } else {
      setUnlockError("Código inválido.");
    }
  }

  async function registerUsage() {
    if (!user || user.plan === "pago") return;
    const updated = { ...user, reportsUsed: (user.reportsUsed || 0) + 1 };
    setUser(updated);
    await saveUser(user.email, { password: updated.password, reportsUsed: updated.reportsUsed, plan: updated.plan });
  }

  const [productsRaw, setProductsRaw] = useState("");
  const [columnMapping, setColumnMapping] = useState([]);
  const fileInputRef = useRef(null);

  const [ncmRaw, setNcmRaw] = useState("");

  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [viewFilter, setViewFilter] = useState("impactados"); // impactados | naoImpactados | todos
  const [empresa, setEmpresa] = useState("");
  const [escritorio, setEscritorio] = useState("");
  const [dataBaseNCM, setDataBaseNCM] = useState("");
  const [fonteBaseNCM, setFonteBaseNCM] = useState("LC 214/2025 e tabela cClassTrib");
  const [mensagem, setMensagem] = useState(
    "Segue o diagnóstico de impacto da Reforma Tributária sobre o catálogo de produtos. Os itens abaixo têm NCM classificada como impactada e precisam de atenção no enquadramento fiscal. Recomendamos encaminhar esta lista para o setor responsável pela atualização no sistema."
  );

  const productParse = useMemo(() => parseProducts(productsRaw), [productsRaw]);
  const ncmRowsAll = useMemo(() => parseBase(ncmRaw), [ncmRaw]);

  const productsHeaderAuto = useMemo(() => looksLikeHeaderRow(productParse.rows[0]), [productParse]);
  const ncmHeaderAuto = useMemo(() => looksLikeHeaderRow(ncmRowsAll[0]), [ncmRowsAll]);

  const productDataRows = useMemo(() => {
    const rows = productParse.rows;
    const body = productsHeaderAuto && rows.length > 1 ? rows.slice(1) : rows;
    return body.filter((r) => r.some((c) => c && c.length));
  }, [productParse, productsHeaderAuto]);

  const ncmDataRows = useMemo(() => {
    const body = ncmHeaderAuto && ncmRowsAll.length > 1 ? ncmRowsAll.slice(1) : ncmRowsAll;
    return body.filter((r) => r.some((c) => c && c.length));
  }, [ncmRowsAll, ncmHeaderAuto]);

  // Assim que uma coluna nova aparece (ou o número de colunas muda), propõe
  // um mapeamento automático — o usuário só ajusta o que estiver errado.
  useEffect(() => {
    if (!productParse.hasDelimiter) {
      setColumnMapping([]);
      return;
    }
    const colCount = productDataRows[0]?.length || 0;
    if (!colCount) return;
    setColumnMapping((prev) => (prev.length === colCount ? prev : guessMapping(colCount)));
  }, [productParse.hasDelimiter, productDataRows]);

  function updateMapping(index, value) {
    setColumnMapping((prev) => {
      const next = [...prev];
      next[index] = value;
      return next;
    });
  }

  function handleProductsFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
        const text = aoa
          .filter((row) => row.some((c) => String(c).trim().length))
          .map((row) => row.map((c) => String(c).trim()).join("\t"))
          .join("\n");
        if (!text) {
          setError("Essa planilha parece estar vazia.");
          return;
        }
        setProductsRaw(text);
        setError("");
      } catch (err) {
        setError("Não consegui ler esse arquivo. Confira se é uma planilha válida (.xlsx, .xls ou .csv).");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  const canRun = productDataRows.length > 0 && ncmDataRows.length > 0;

  function runAnalysis() {
    setError("");

    if (productDataRows.length === 0) {
      setError("Cole a lista de produtos ou importe a planilha do cliente antes de rodar.");
      return;
    }
    if (productParse.hasDelimiter && !columnMapping.includes("ncm")) {
      setError("Marque qual coluna é o NCM antes de rodar o comparativo.");
      return;
    }
    if (ncmDataRows.length === 0) {
      setError("Cole a base de NCMs impactadas antes de rodar.");
      return;
    }

    const ncmMap = new Map();
    ncmDataRows.forEach((r) => {
      const ncm = normalizeNCM(r[0]);
      if (!ncm) return;
      ncmMap.set(ncm, r[1] || "Impactado pela Reforma Tributária");
    });

    const idx = (key) => columnMapping.indexOf(key);
    const useMapping = productParse.hasDelimiter;

    const detail = productDataRows.map((r) => {
      const codigo = useMapping ? (idx("codigo") >= 0 ? r[idx("codigo")] : "") : r[0];
      const descricao = useMapping ? (idx("descricao") >= 0 ? r[idx("descricao")] : "") : r[1];
      const ncm = useMapping ? (idx("ncm") >= 0 ? r[idx("ncm")] : "") : r[2];
      const ncmNorm = normalizeNCM(ncm);
      const impactada = ncmMap.has(ncmNorm);
      return {
        codigo: codigo || "",
        descricao: descricao || "",
        ncm: ncm || "",
        impactado: impactada,
        mensagem: impactada ? ncmMap.get(ncmNorm) : "",
      };
    });

    const totalAnalisado = detail.length;
    const totalImpactado = detail.filter((d) => d.impactado).length;
    const ncmsDistintasImpactadas = new Set(
      detail.filter((d) => d.impactado).map((d) => normalizeNCM(d.ncm))
    ).size;
    const naoEncontrados = detail.filter((d) => !normalizeNCM(d.ncm)).length;

    setResult({ detail, totalAnalisado, totalImpactado, ncmsDistintasImpactadas, naoEncontrados });
    setViewFilter("impactados");
    registerUsage();
  }

  function downloadExcel() {
    if (!result) return;
    const hoje = new Date().toLocaleDateString("pt-BR");
    const empresaLabel = empresa.trim() || "—";
    const escritorioLabel = escritorio.trim() || "—";

    // Gera um workbook Excel de verdade (SpreadsheetML), com abas reais e
    // nomeadas — mais confiável que o truque de HTML disfarçado de .xls,
    // que às vezes agrupa tudo numa única aba.
    function esc(s) {
      return (s == null ? "" : String(s))
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    }

    function cell(styleId, type, value, mergeAcross) {
      const merge = mergeAcross ? " ss:MergeAcross=\"" + mergeAcross + "\"" : "";
      return "<Cell ss:StyleID=\"" + styleId + "\"" + merge + "><Data ss:Type=\"" + type + "\">" + esc(value) + "</Data></Cell>";
    }
    function row(cellsXml, height) {
      const h = height ? " ss:Height=\"" + height + "\"" : "";
      return "<Row" + h + ">" + cellsXml + "</Row>";
    }

    const styles =
      "<Styles>" +
      "<Style ss:ID=\"titleBanner\"><Font ss:Size=\"18\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#173A8A\" ss:Pattern=\"Solid\"/><Alignment ss:Vertical=\"Center\"/></Style>" +
      "<Style ss:ID=\"sub\"><Font ss:Size=\"9\" ss:Color=\"#5C5648\"/></Style>" +
      "<Style ss:ID=\"label\"><Font ss:Size=\"11\" ss:Bold=\"1\"/></Style>" +
      "<Style ss:ID=\"value\"><Font ss:Size=\"11\"/></Style>" +
      "<Style ss:ID=\"indicHeader\"><Font ss:Size=\"12\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#173A8A\" ss:Pattern=\"Solid\"/></Style>" +
      "<Style ss:ID=\"indicLabel\"><Font ss:Size=\"11\" ss:Bold=\"1\"/><Interior ss:Color=\"#F2F5FC\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#B9C3DA\"/></Borders></Style>" +
      "<Style ss:ID=\"vDark\"><Font ss:Size=\"16\" ss:Bold=\"1\" ss:Color=\"#1C2B33\"/><Interior ss:Color=\"#F2F5FC\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#B9C3DA\"/></Borders></Style>" +
      "<Style ss:ID=\"vRed\"><Font ss:Size=\"16\" ss:Bold=\"1\" ss:Color=\"#D1425A\"/><Interior ss:Color=\"#F2F5FC\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#B9C3DA\"/></Borders></Style>" +
      "<Style ss:ID=\"vBlue\"><Font ss:Size=\"16\" ss:Bold=\"1\" ss:Color=\"#2E6BFF\"/><Interior ss:Color=\"#F2F5FC\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#B9C3DA\"/></Borders></Style>" +
      "<Style ss:ID=\"vOrange\"><Font ss:Size=\"16\" ss:Bold=\"1\" ss:Color=\"#B5790A\"/><Interior ss:Color=\"#F2F5FC\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#B9C3DA\"/></Borders></Style>" +
      "<Style ss:ID=\"bannerRed\"><Font ss:Size=\"20\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#D1425A\" ss:Pattern=\"Solid\"/><Alignment ss:Vertical=\"Center\"/></Style>" +
      "<Style ss:ID=\"bannerGreen\"><Font ss:Size=\"20\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#2FAE7A\" ss:Pattern=\"Solid\"/><Alignment ss:Vertical=\"Center\"/></Style>" +
      "<Style ss:ID=\"tableHeader\"><Font ss:Size=\"11\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#173A8A\" ss:Pattern=\"Solid\"/></Style>" +
      "<Style ss:ID=\"rowEven\"><Font ss:Size=\"10.5\"/><Interior ss:Color=\"#DCE6FF\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#B9C3DA\"/></Borders></Style>" +
      "<Style ss:ID=\"rowOdd\"><Font ss:Size=\"10.5\"/><Interior ss:Color=\"#FFFFFF\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#B9C3DA\"/></Borders></Style>" +
      "<Style ss:ID=\"aviso\"><Font ss:Size=\"9\" ss:Italic=\"1\" ss:Color=\"#5C5648\"/><Alignment ss:WrapText=\"1\" ss:Vertical=\"Top\"/></Style>" +
      "</Styles>";

    // --- Aba Resumo ---
    const avisoLegal = buildAvisoLegal(dataBaseNCM, fonteBaseNCM, hoje);
    const resumoRows =
      row(cell("titleBanner", "String", "Diagnóstico de Impacto — Reforma Tributária", 1), 30) +
      row(cell("sub", "String", "Gerado em " + hoje, 1)) +
      row("") +
      row(cell("label", "String", "Cliente") + cell("value", "String", empresaLabel)) +
      row(cell("label", "String", "Escritório responsável") + cell("value", "String", escritorioLabel)) +
      row("") +
      row(cell("indicHeader", "String", "Indicador") + cell("indicHeader", "String", "Valor")) +
      row(cell("indicLabel", "String", "Produtos analisados") + cell("vDark", "Number", result.totalAnalisado)) +
      row(cell("indicLabel", "String", "Produtos impactados") + cell("vRed", "Number", result.totalImpactado)) +
      row(cell("indicLabel", "String", "NCMs distintas impactadas") + cell("vBlue", "Number", result.ncmsDistintasImpactadas)) +
      row(cell("indicLabel", "String", "NCM ilegível / não identificado") + cell("vOrange", "Number", result.naoEncontrados)) +
      row("") +
      row(cell("aviso", "String", avisoLegal, 1), 60);

    const resumoSheet =
      "<Worksheet ss:Name=\"Resumo\"><Table ss:DefaultColumnWidth=\"140\">" +
      "<Column ss:Width=\"230\"/><Column ss:Width=\"170\"/>" +
      resumoRows +
      "</Table></Worksheet>";

    // --- Abas Impactados / Não impactados ---
    function sectionSheet(sheetName, title, bannerStyle, rows) {
      const headerRow = row(
        ["Código", "Descrição", "NCM", "Impactado", "Motivo / mensagem"]
          .map((h) => cell("tableHeader", "String", h))
          .join("")
      );
      const bodyRows = rows
        .map((d, i) => {
          const style = i % 2 === 0 ? "rowEven" : "rowOdd";
          return row(
            cell(style, "String", d.codigo) +
            cell(style, "String", d.descricao) +
            cell(style, "String", d.ncm) +
            cell(style, "String", d.impactado ? "Sim" : "Não") +
            cell(style, "String", d.mensagem)
          );
        })
        .join("");
      const bannerRow = row(cell(bannerStyle, "String", title + " (" + rows.length + ")", 4), 26);
      return (
        "<Worksheet ss:Name=\"" + esc(sheetName) + "\"><Table ss:DefaultColumnWidth=\"120\">" +
        "<Column ss:Width=\"90\"/><Column ss:Width=\"260\"/><Column ss:Width=\"90\"/><Column ss:Width=\"80\"/><Column ss:Width=\"260\"/>" +
        bannerRow +
        row("") +
        headerRow +
        bodyRows +
        "</Table></Worksheet>"
      );
    }

    const impactadosSheet = sectionSheet(
      "Impactados", "Produtos Impactados", "bannerRed",
      result.detail.filter((d) => d.impactado)
    );
    const naoImpactadosSheet = sectionSheet(
      "Nao impactados", "Produtos Não Impactados", "bannerGreen",
      result.detail.filter((d) => !d.impactado)
    );

    const workbook =
      "<?xml version=\"1.0\"?>\n" +
      "<?mso-application progid=\"Excel.Sheet\"?>\n" +
      "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" " +
      "xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
      "xmlns:x=\"urn:schemas-microsoft-com:office:excel\" " +
      "xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" " +
      "xmlns:html=\"http://www.w3.org/TR/REC-html40\">" +
      styles +
      resumoSheet +
      impactadosSheet +
      naoImpactadosSheet +
      "</Workbook>";

    try {
      const blob = new Blob([workbook], { type: "application/vnd.ms-excel" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const nomeArquivo = (empresa.trim() || "cliente").toLowerCase().replace(/\s+/g, "-");
      a.href = url;
      a.download = "diagnostico-ncm-" + nomeArquivo + ".xls";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      setError("Não consegui gerar a planilha Excel. Tente novamente.");
    }
  }

  function downloadWord() {
    if (!result) return;
    const impacted = result.detail.filter((d) => d.impactado);
    const empresaLabel = empresa.trim() || "Cliente";
    const escritorioLabel = escritorio.trim();
    const hoje = new Date().toLocaleDateString("pt-BR");

    // Só NCMs distintas: agrupa pelo código normalizado e formata como
    // dddd.dd.dd (padrão de 8 dígitos), ordenado.
    const distinctMap = new Map();
    for (const d of impacted) {
      const norm = normalizeNCM(d.ncm);
      if (!norm || distinctMap.has(norm)) continue;
      const formatted = norm.slice(0, 4) + "." + norm.slice(4, 6) + "." + norm.slice(6, 8);
      distinctMap.set(norm, formatted);
    }
    const distinctNCMs = Array.from(distinctMap.values()).sort();

    const rowsHtmlParts = distinctNCMs.map(
      (ncm) => "<tr><td>" + escapeHtml(ncm) + "</td></tr>"
    );
    const rowsHtml = rowsHtmlParts.join("");

    const letterheadHtml = escritorioLabel
      ? "<p class=\"letterhead\">" + escapeHtml(escritorioLabel) + "</p>"
      : "";
    const footerHtml = escritorioLabel
      ? "<p class=\"footer\">Relatório gerado por " + escapeHtml(escritorioLabel) + ".</p>"
      : "";
    const avisoLegal = buildAvisoLegal(dataBaseNCM, fonteBaseNCM, hoje);

    const htmlParts = [
      "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns=\"http://www.w3.org/TR/REC-html40\">",
      "<head><meta charset=\"utf-8\"><title>Diagnóstico - " + escapeHtml(empresaLabel) + "</title>",
      "<style>",
      "body { font-family: Calibri, Arial, sans-serif; color:#1C2B33; font-size: 12px; }",
      ".letterhead { font-size: 11px; color:#5C5648; margin: 0 0 6px; }",
      "h1 { font-size: 24px; color:#173A8A; margin-bottom: 2px; }",
      ".sub { color:#5C5648; font-size: 12px; margin: 2px 0; }",
      ".meta { margin: 14px 0; }",
      ".meta b { color:#173A8A; }",
      "table { border-collapse: collapse; width:260px; margin-top:16px; }",
      "th, td { border:1px solid #B9C3DA; padding:7px 12px; font-size:12px; text-align:left; }",
      "th { background:#173A8A; color:#fff; font-size:13px; }",
      "tr:nth-child(even) td { background:#F2F5FC; }",
      ".aviso { margin-top: 16px; padding: 10px 12px; border-left: 3px solid #B5790A; background:#FFF8ED; font-size: 10.5px; color:#5C5648; font-style: italic; }",
      ".footer { margin-top: 18px; font-size: 10.5px; color:#8592B0; }",
      "</style></head>",
      "<body>",
      letterheadHtml,
      "<h1>Diagnóstico de Impacto — Reforma Tributária</h1>",
      "<p class=\"sub\">Cliente: <b>" + escapeHtml(empresaLabel) + "</b></p>",
      "<p class=\"sub\">Data: " + hoje + "</p>",
      "<div class=\"meta\">",
      "<p>" + escapeHtml(mensagem) + "</p>",
      "<p><b>Produtos analisados:</b> " + result.totalAnalisado + " &nbsp;&nbsp; " +
        "<b>Produtos impactados:</b> " + result.totalImpactado + " &nbsp;&nbsp; " +
        "<b>NCMs distintas impactadas:</b> " + distinctNCMs.length + "</p>",
      "</div>",
      "<table>",
      "<tr><th>NCM impactada</th></tr>",
      rowsHtml,
      "</table>",
      "<p class=\"aviso\">" + escapeHtml(avisoLegal) + "</p>",
      footerHtml,
      "</body></html>",
    ];
    const html = htmlParts.join("\n");

    try {
      const blob = new Blob(["\ufeff", html], { type: "application/msword" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "ncm-impactados-" + empresaLabel.replace(/\s+/g, "-").toLowerCase() + ".doc";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      setError("Não consegui gerar o relatório em Word. Tente novamente.");
    }
  }

  function reset() {
    setProductsRaw("");
    setNcmRaw("");
    setResult(null);
    setError("");
    setColumnMapping([]);
  }

  const pct = result && result.totalAnalisado ? Math.round((result.totalImpactado / result.totalAnalisado) * 100) : 0;

  if (!user) {
    return (
      <LoginScreen
        email={authEmail}
        setEmail={setAuthEmail}
        password={authPassword}
        setPassword={setAuthPassword}
        onSubmit={handleAuthSubmit}
        loading={authLoading}
        error={authError}
      />
    );
  }

  if (user.plan !== "pago" && (user.reportsUsed || 0) >= FREE_LIMIT) {
    return (
      <PaywallScreen
        email={user.email}
        onLogout={handleLogout}
        unlockInput={unlockInput}
        setUnlockInput={setUnlockInput}
        onUnlock={handleUnlock}
        unlockError={unlockError}
      />
    );
  }

  return (
    <div style={{ background: INK, minHeight: "100%", color: "#EAEFFB", fontFamily: "'IBM Plex Sans', 'Segoe UI', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500&display=swap');
        .dncm-ta { background:${PANEL}; }
        .dncm-ta:focus { outline: none; border-color: ${BLUE} !important; box-shadow: 0 0 0 3px ${BLUE_DIM}55; }
        .dncm-btn-primary:hover { background: #4A80FF !important; }
        .dncm-btn-primary:disabled { background:${LINE} !important; color:${TEXT_DIM} !important; cursor:not-allowed !important; }
        .dncm-btn-ghost:hover { border-color: ${BLUE} !important; color: ${BLUE} !important; }
        .dncm-row:hover { background: #10152488; }
        select.dncm-select { background: ${INK}; color: #EAEFFB; border: 1px solid ${LINE}; border-radius: 5px; padding: 5px 8px; font-size: 12.5px; font-family: 'IBM Plex Sans', sans-serif; }
        select.dncm-select:focus { outline: none; border-color: ${BLUE}; }
      `}</style>

      <div style={{ maxWidth: 980, margin: "0 auto", padding: "36px 24px 64px" }}>
        <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 12, marginBottom: 14 }}>
          <span style={{ fontSize: 12, color: TEXT_DIM }}>
            {user.email}
            {user.plan === "pago" ? (
              <span style={{ color: OK, marginLeft: 8 }}>· plano ativo</span>
            ) : (
              <span style={{ marginLeft: 8 }}>· {(user.reportsUsed || 0)}/{FREE_LIMIT} diagnósticos grátis usados</span>
            )}
          </span>
          <button
            onClick={handleLogout}
            style={{ fontSize: 12, color: TEXT_DIM, background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}
          >
            Sair
          </button>
        </div>
        <header style={{ marginBottom: 30, display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{
            width: 38, height: 38, borderRadius: 8, background: `linear-gradient(135deg, ${BLUE}, ${BLUE_DIM})`,
            display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
          }}>
            <Zap size={19} color="#fff" strokeWidth={2.3} />
          </div>
          <div>
            <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0, letterSpacing: -0.3 }}>
              Diagnóstico de impacto — Reforma Tributária
            </h1>
            <p style={{ margin: "3px 0 0", color: TEXT_DIM, fontSize: 13.5 }}>
              Compare o catálogo do cliente com a base de NCMs impactadas e gere o relatório em segundos.
            </p>
          </div>
        </header>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 20, alignItems: "start" }}>
          <div style={{ border: `1px solid ${LINE}`, borderRadius: 8, padding: 16, background: PANEL }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 2 }}>
              <div style={{ fontSize: 14, fontWeight: 700 }}>Produtos do cliente</div>
              <button
                className="dncm-btn-ghost"
                onClick={() => fileInputRef.current?.click()}
                style={{
                  padding: "5px 10px", borderRadius: 5, fontSize: 11.5, fontWeight: 600,
                  border: `1px solid ${LINE}`, background: "transparent", color: TEXT_DIM, cursor: "pointer",
                  display: "flex", alignItems: "center", gap: 6, transition: "all .15s",
                }}
              >
                <Upload size={12} /> Importar planilha
              </button>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                style={{ display: "none" }}
                onChange={handleProductsFile}
              />
            </div>
            <div style={{ fontSize: 12, color: TEXT_DIM, marginBottom: 10 }}>
              Cole a tabela como veio do cliente, ou importe o arquivo Excel/CSV direto
            </div>

            <textarea
              className="dncm-ta"
              value={productsRaw}
              onChange={(e) => setProductsRaw(e.target.value)}
              placeholder="Cole aqui (Ctrl+V)…"
              rows={8}
              style={{
                width: "100%", boxSizing: "border-box", padding: "10px 12px",
                border: `1px solid ${LINE}`, borderRadius: 6, color: "#EAEFFB",
                fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, lineHeight: 1.6,
                resize: "vertical", transition: "border-color .15s",
              }}
            />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
              <span style={{ fontSize: 11.5, color: TEXT_DIM }}>
                {productDataRows.length > 0 ? (productsHeaderAuto ? "cabeçalho identificado e ignorado" : "sem cabeçalho detectado") : ""}
              </span>
              <span style={{ fontSize: 12, color: productsRaw.trim() ? BLUE : "#3A4258", fontFamily: "'IBM Plex Mono', monospace" }}>
                {productDataRows.length} linha{productDataRows.length === 1 ? "" : "s"}
              </span>
            </div>

            {productParse.hasDelimiter && productDataRows.length > 0 && (
              <div style={{ marginTop: 14, paddingTop: 14, borderTop: `1px solid ${LINE}` }}>
                <div style={{ fontSize: 12, color: TEXT_DIM, marginBottom: 8 }}>
                  Confirme o que cada coluna representa — o que não precisar, marque como "Ignorar coluna":
                </div>
                {productParse.riskySpaceSplit && (
                  <div style={{ fontSize: 11.5, color: "#F5B942", marginBottom: 10 }}>
                    Não encontrei tabulação, ";" ou "," para separar as colunas.
                  </div>
                )}
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {columnMapping.map((mapped, i) => (
                    <div key={i} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <span style={{
                        fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: TEXT_DIM,
                        width: 26, flexShrink: 0,
                      }}>
                        C{i + 1}
                      </span>
                      <select
                        className="dncm-select"
                        value={mapped}
                        onChange={(e) => updateMapping(i, e.target.value)}
                        style={{ width: 150, flexShrink: 0 }}
                      >
                        {FIELD_OPTIONS.map((opt) => (
                          <option key={opt.key} value={opt.key}>{opt.label}</option>
                        ))}
                      </select>
                      <span style={{
                        fontSize: 12, color: "#C6CEE6", fontFamily: "'IBM Plex Mono', monospace",
                        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1,
                      }}>
                        {productDataRows[0]?.[i] || "—"}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {!productParse.hasDelimiter && productDataRows.length > 0 && (
              <div style={{ marginTop: 14, paddingTop: 14, borderTop: `1px solid ${LINE}`, fontSize: 12, color: TEXT_DIM }}>
                Não encontrei colunas separadas — o NCM é identificado direto em cada linha, sem precisar mapear nada.
              </div>
            )}
          </div>

          <div style={{ border: `1px solid ${LINE}`, borderRadius: 8, padding: 16, background: PANEL }}>
            <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 2 }}>Base de NCMs impactadas</div>
            <div style={{ fontSize: 12, color: TEXT_DIM, marginBottom: 10 }}>
              Uma linha por NCM: o código e, se quiser, o motivo do impacto
            </div>
            <textarea
              className="dncm-ta"
              value={ncmRaw}
              onChange={(e) => setNcmRaw(e.target.value)}
              placeholder="Cole aqui (Ctrl+V)…"
              rows={8}
              style={{
                width: "100%", boxSizing: "border-box", padding: "10px 12px",
                border: `1px solid ${LINE}`, borderRadius: 6, color: "#EAEFFB",
                fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, lineHeight: 1.6,
                resize: "vertical", transition: "border-color .15s",
              }}
            />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
              <span style={{ fontSize: 11.5, color: TEXT_DIM }}>
                {ncmDataRows.length > 0 ? (ncmHeaderAuto ? "cabeçalho identificado e ignorado" : "sem cabeçalho detectado") : ""}
              </span>
              <span style={{ fontSize: 12, color: ncmRaw.trim() ? BLUE : "#3A4258", fontFamily: "'IBM Plex Mono', monospace" }}>
                {ncmDataRows.length} linha{ncmDataRows.length === 1 ? "" : "s"}
              </span>
            </div>
          </div>
        </div>

        {error && (
          <div style={{ marginBottom: 16, color: ALERT, fontSize: 13.5, display: "flex", gap: 8, alignItems: "center" }}>
            <AlertTriangle size={15} /> {error}
          </div>
        )}

        <div style={{ display: "flex", gap: 10, marginBottom: 30 }}>
          <button
            className="dncm-btn-primary"
            onClick={runAnalysis}
            disabled={!canRun}
            style={{
              padding: "11px 22px", borderRadius: 6, fontSize: 14, fontWeight: 700,
              border: "none", background: BLUE, color: "#fff", cursor: "pointer",
              display: "flex", alignItems: "center", gap: 8, transition: "background .15s",
            }}
          >
            <Zap size={15} /> Rodar comparativo
          </button>
          {result && (
            <button
              className="dncm-btn-ghost"
              onClick={reset}
              style={{
                padding: "11px 18px", borderRadius: 6, fontSize: 13.5, fontWeight: 600,
                border: `1px solid ${LINE}`, background: "transparent", color: TEXT_DIM, cursor: "pointer",
                display: "flex", alignItems: "center", gap: 8, transition: "all .15s",
              }}
            >
              <RotateCcw size={14} /> Limpar tudo
            </button>
          )}
        </div>

        {result && (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 22 }}>
              <StatCard label="Analisados" value={result.totalAnalisado} />
              <StatCard label="Impactados" value={result.totalImpactado} accent={ALERT} sub={`${pct}% do total`} />
              <StatCard label="NCMs distintas impactadas" value={result.ncmsDistintasImpactadas} accent={BLUE} />
              <StatCard label="NCM ilegível" value={result.naoEncontrados} accent={result.naoEncontrados ? "#F5B942" : TEXT_DIM} />
            </div>

            <div style={{ border: `1px solid ${LINE}`, borderRadius: 8, overflow: "hidden", marginBottom: 22, background: PANEL }}>
              <div style={{
                padding: "10px 16px", background: "#10152A", borderBottom: `1px solid ${LINE}`,
                display: "flex", alignItems: "center", gap: 6,
              }}>
                {[
                  { key: "impactados", label: `Impactados (${result.totalImpactado})` },
                  { key: "naoImpactados", label: `Não impactados (${result.totalAnalisado - result.totalImpactado})` },
                  { key: "todos", label: `Todos (${result.totalAnalisado})` },
                ].map((tab) => (
                  <button
                    key={tab.key}
                    onClick={() => setViewFilter(tab.key)}
                    style={{
                      padding: "5px 12px", borderRadius: 5, fontSize: 12, fontWeight: 600, cursor: "pointer",
                      border: `1px solid ${viewFilter === tab.key ? BLUE : LINE}`,
                      background: viewFilter === tab.key ? BLUE_DIM : "transparent",
                      color: viewFilter === tab.key ? "#EAEFFB" : TEXT_DIM,
                      transition: "all .15s",
                    }}
                  >
                    {tab.label}
                  </button>
                ))}
              </div>
              {(() => {
                const rows =
                  viewFilter === "impactados"
                    ? result.detail.filter((d) => d.impactado)
                    : viewFilter === "naoImpactados"
                    ? result.detail.filter((d) => !d.impactado)
                    : result.detail;
                if (rows.length === 0) {
                  return (
                    <div style={{ padding: 28, textAlign: "center", color: OK, fontSize: 13.5, display: "flex", gap: 8, justifyContent: "center", alignItems: "center" }}>
                      <CheckCircle2 size={16} /> Nenhum produto nessa categoria.
                    </div>
                  );
                }
                return (
                  <div style={{ maxHeight: 360, overflowY: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                      <thead>
                        <tr>
                          {["Código", "Descrição", "NCM", "Impactado", "Mensagem"].map((h) => (
                            <th key={h} style={{
                              textAlign: "left", padding: "9px 16px", position: "sticky", top: 0,
                              background: PANEL, borderBottom: `1px solid ${LINE}`, fontWeight: 600, color: TEXT_DIM,
                            }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((d, i) => (
                          <tr key={i} className="dncm-row" style={{ borderTop: `1px solid ${LINE}` }}>
                            <td style={{ padding: "9px 16px", fontFamily: "'IBM Plex Mono', monospace", color: "#C6CEE6" }}>{d.codigo}</td>
                            <td style={{ padding: "9px 16px", maxWidth: 260, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{d.descricao}</td>
                            <td style={{ padding: "9px 16px", fontFamily: "'IBM Plex Mono', monospace", color: "#C6CEE6" }}>{d.ncm}</td>
                            <td style={{ padding: "9px 16px", color: d.impactado ? ALERT : OK, fontWeight: 600 }}>{d.impactado ? "Sim" : "Não"}</td>
                            <td style={{ padding: "9px 16px", color: ALERT }}>{d.mensagem}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                );
              })()}
            </div>

            <div style={{ border: `1px solid ${LINE}`, borderRadius: 8, padding: 16, background: PANEL, marginBottom: 22 }}>
              <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 12 }}>Dados do relatório</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 10 }}>
                <div>
                  <label style={{ fontSize: 12, color: TEXT_DIM, display: "block", marginBottom: 5 }}>Nome do seu escritório (opcional)</label>
                  <input
                    value={escritorio}
                    onChange={(e) => setEscritorio(e.target.value)}
                    placeholder="Ex: Contábil Silva & Associados"
                    style={{
                      width: "100%", boxSizing: "border-box", padding: "8px 12px", borderRadius: 6,
                      border: `1px solid ${LINE}`, background: INK, color: "#EAEFFB", fontSize: 13,
                    }}
                  />
                </div>
                <div>
                  <label style={{ fontSize: 12, color: TEXT_DIM, display: "block", marginBottom: 5 }}>Nome do cliente</label>
                  <input
                    value={empresa}
                    onChange={(e) => setEmpresa(e.target.value)}
                    placeholder="Ex: Supermercado Boa Compra"
                    style={{
                      width: "100%", boxSizing: "border-box", padding: "8px 12px", borderRadius: 6,
                      border: `1px solid ${LINE}`, background: INK, color: "#EAEFFB", fontSize: 13,
                    }}
                  />
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 10 }}>
                <div>
                  <label style={{ fontSize: 12, color: TEXT_DIM, display: "block", marginBottom: 5 }}>Data de referência da base de NCMs</label>
                  <input
                    value={dataBaseNCM}
                    onChange={(e) => setDataBaseNCM(e.target.value)}
                    placeholder="Ex: 05/09/2026 (em branco = data de hoje)"
                    style={{
                      width: "100%", boxSizing: "border-box", padding: "8px 12px", borderRadius: 6,
                      border: `1px solid ${LINE}`, background: INK, color: "#EAEFFB", fontSize: 13,
                    }}
                  />
                </div>
                <div>
                  <label style={{ fontSize: 12, color: TEXT_DIM, display: "block", marginBottom: 5 }}>Fonte da base de NCMs</label>
                  <input
                    value={fonteBaseNCM}
                    onChange={(e) => setFonteBaseNCM(e.target.value)}
                    placeholder="Ex: LC 214/2025 e tabela cClassTrib"
                    style={{
                      width: "100%", boxSizing: "border-box", padding: "8px 12px", borderRadius: 6,
                      border: `1px solid ${LINE}`, background: INK, color: "#EAEFFB", fontSize: 13,
                    }}
                  />
                </div>
              </div>
              <div>
                <label style={{ fontSize: 12, color: TEXT_DIM, display: "block", marginBottom: 5 }}>Mensagem explicativa</label>
                <textarea
                  value={mensagem}
                  onChange={(e) => setMensagem(e.target.value)}
                  rows={3}
                  style={{
                    width: "100%", boxSizing: "border-box", padding: "8px 12px", borderRadius: 6,
                    border: `1px solid ${LINE}`, background: INK, color: "#EAEFFB", fontSize: 12.5,
                    fontFamily: "'IBM Plex Sans', sans-serif", resize: "vertical",
                  }}
                />
              </div>
            </div>

            <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button
                className="dncm-btn-primary"
                onClick={downloadExcel}
                style={{
                  padding: "11px 22px", borderRadius: 6, fontSize: 14, fontWeight: 700,
                  border: "none", background: BLUE, color: "#fff", cursor: "pointer",
                  display: "flex", alignItems: "center", gap: 8, transition: "background .15s",
                }}
              >
                <Download size={15} /> Baixar planilha Excel
              </button>
              <button
                className="dncm-btn-ghost"
                onClick={downloadWord}
                style={{
                  padding: "11px 22px", borderRadius: 6, fontSize: 14, fontWeight: 700,
                  border: `1px solid ${BLUE}`, background: "transparent", color: BLUE, cursor: "pointer",
                  display: "flex", alignItems: "center", gap: 8, transition: "all .15s",
                }}
              >
                <Download size={15} /> Gerar relatório em Word
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function StatCard({ label, value, accent, sub }) {
  return (
    <div style={{ border: `1px solid ${LINE}`, borderRadius: 8, padding: "14px 16px", background: PANEL }}>
      <div style={{ fontSize: 11.5, color: TEXT_DIM, marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 26, fontWeight: 700, color: accent || "#EAEFFB", fontFamily: "'IBM Plex Mono', monospace" }}>
        {value.toLocaleString("pt-BR")}
      </div>
      {sub && <div style={{ fontSize: 11.5, color: accent || TEXT_DIM, marginTop: 3 }}>{sub}</div>}
    </div>
  );
}
