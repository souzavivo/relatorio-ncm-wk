import { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, VerticalAlign,
  Footer, PageNumber,
} from "docx";
import { saveAs } from "file-saver";
import { AlertTriangle, CheckCircle2, Download, RotateCcw, Zap } from "lucide-react";

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
  { key: "cst", label: "CST" },
  { key: "cfop", label: "CFOP" },
  { key: "csosn", label: "CSOSN" },
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
  const defaults = ["codigo", "descricao", "ncm", "cst", "cfop", "csosn"];
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
    <div style={{ minHeight: "100vh", width: "100%", background: INK, color: "#EAEFFB", fontFamily: "'IBM Plex Sans', 'Segoe UI', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500&display=swap');
        html, body, #root { height: 100%; margin: 0; }
        .dncm-ta:focus, .dncm-input:focus { outline: none; border-color: ${BLUE} !important; box-shadow: 0 0 0 3px ${BLUE_DIM}55; }
        .dncm-pw-mask { -webkit-text-security: disc; text-security: disc; }
        .dncm-btn-primary:hover { background: #4A80FF !important; }
        .dncm-btn-primary:disabled { background:${LINE} !important; color:${TEXT_DIM} !important; cursor:not-allowed !important; }
        .dncm-auth-wrap { display: flex; min-height: 100vh; width: 100%; }
        .dncm-auth-brand {
          flex: 1.1; display: flex; flex-direction: column; justify-content: center;
          padding: 64px; position: relative; overflow: hidden;
          background: radial-gradient(circle at 20% 20%, ${BLUE_DIM} 0%, ${INK} 65%);
        }
        .dncm-auth-brand::after {
          content: ""; position: absolute; inset: 0;
          background: radial-gradient(circle at 85% 80%, ${BLUE}33 0%, transparent 55%);
        }
        .dncm-auth-brand > * { position: relative; z-index: 1; }
        .dncm-auth-form-col { flex: 1; display: flex; align-items: center; justify-content: center; padding: 32px; }
        .dncm-auth-card { width: 100%; max-width: 400px; }
        @media (max-width: 860px) {
          .dncm-auth-wrap { flex-direction: column; }
          .dncm-auth-brand { flex: none; padding: 40px 28px; }
          .dncm-auth-brand h2 { font-size: 22px !important; }
        }
      `}</style>
      <div className="dncm-auth-wrap">
        <div className="dncm-auth-brand">
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 28 }}>
            <div style={{ width: 42, height: 42, borderRadius: 10, background: `linear-gradient(135deg, ${BLUE}, ${BLUE_DIM})`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              <Zap size={21} color="#fff" strokeWidth={2.3} />
            </div>
            <div style={{ fontSize: 18, fontWeight: 700 }}>Diagnóstico NCM</div>
          </div>
          <h2 style={{ fontSize: 32, fontWeight: 700, maxWidth: 460, lineHeight: 1.28, margin: "0 0 16px" }}>
            Descubra em segundos quais produtos serão impactados pela Reforma Tributária.
          </h2>
          <p style={{ fontSize: 14.5, color: "#B7C2E0", maxWidth: 420, lineHeight: 1.6 }}>
            Cole a tabela do cliente, cruze com a base de NCMs impactadas e gere o relatório pronto pra entregar — Excel e Word em minutos.
          </p>
        </div>
        <div className="dncm-auth-form-col">
          <div className="dncm-auth-card">{children}</div>
        </div>
      </div>
    </div>
  );
}

function LoginScreen({
  mode, setMode,
  email, setEmail, password, setPassword, confirmPassword, setConfirmPassword,
  onSubmit, loading, error,
}) {
  const isSignup = mode === "signup";
  return (
    <AuthShell>
      <div style={{ display: "flex", gap: 4, marginBottom: 20, background: INK, border: `1px solid ${LINE}`, borderRadius: 8, padding: 3 }}>
        {[{ key: "login", label: "Entrar" }, { key: "signup", label: "Criar conta" }].map((tab) => (
          <button
            key={tab.key}
            onClick={() => setMode(tab.key)}
            style={{
              flex: 1, padding: "8px 0", borderRadius: 6, border: "none", cursor: "pointer",
              fontSize: 13, fontWeight: 600,
              background: mode === tab.key ? BLUE : "transparent",
              color: mode === tab.key ? "#fff" : TEXT_DIM,
              transition: "all .15s",
            }}
          >
            {tab.label}
          </button>
        ))}
      </div>
      <p style={{ fontSize: 12.5, color: TEXT_DIM, margin: "0 0 20px" }}>
        {isSignup ? "Crie sua conta com e-mail e senha." : "Entre com o e-mail e senha da sua conta."}
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
      <div style={{ marginBottom: isSignup ? 10 : 16 }}>
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
          onKeyDown={(e) => e.key === "Enter" && !isSignup && onSubmit()}
          placeholder="sua senha"
          style={{ width: "100%", boxSizing: "border-box", padding: "9px 12px", borderRadius: 6, border: `1px solid ${LINE}`, background: INK, color: "#EAEFFB", fontSize: 13.5 }}
        />
      </div>
      {isSignup && (
        <div style={{ marginBottom: 16 }}>
          <label style={{ fontSize: 12, color: TEXT_DIM, display: "block", marginBottom: 5 }}>Confirmar senha</label>
          <input
            className="dncm-input dncm-pw-mask"
            type="text"
            autoComplete="off"
            autoCorrect="off"
            autoCapitalize="off"
            spellCheck="false"
            name="dncm-pass-confirm-field"
            value={confirmPassword}
            onChange={(e) => setConfirmPassword(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && onSubmit()}
            placeholder="digite a senha de novo"
            style={{ width: "100%", boxSizing: "border-box", padding: "9px 12px", borderRadius: 6, border: `1px solid ${LINE}`, background: INK, color: "#EAEFFB", fontSize: 13.5 }}
          />
        </div>
      )}
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
        {loading ? "Aguarde…" : isSignup ? "Criar conta" : "Entrar"}
      </button>
      <p style={{ fontSize: 11, color: "#5C6584", marginTop: 14, marginBottom: 0 }}>
        Teste grátis com {FREE_LIMIT} diagnósticos completos por conta.
      </p>
    </AuthShell>
  );
}

function PaywallScreen({ email, wasSubscriber, onLogout, unlockInput, setUnlockInput, onUnlock, unlockError }) {
  return (
    <AuthShell>
      <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 6 }}>
        {wasSubscriber ? "Sua assinatura expirou" : `Você usou seus ${FREE_LIMIT} diagnósticos grátis`}
      </div>
      <p style={{ fontSize: 12.5, color: TEXT_DIM, margin: "0 0 18px" }}>
        Logado como <b style={{ color: "#C6CEE6" }}>{email}</b>.{" "}
        {wasSubscriber
          ? "Renove pra continuar gerando diagnósticos ilimitados."
          : "Pra continuar gerando diagnósticos ilimitados, assine o plano completo."}
      </p>
      <div style={{ border: `1px solid ${BLUE}`, borderRadius: 8, padding: 16, marginBottom: 18, background: "#0E1730" }}>
        <div style={{ fontSize: 22, fontWeight: 700, color: "#EAEFFB" }}>R$ 79,90<span style={{ fontSize: 13, color: TEXT_DIM, fontWeight: 500 }}> /mês</span></div>
        <div style={{ fontSize: 12, color: TEXT_DIM, marginTop: 4 }}>Diagnósticos ilimitados, Excel e Word completos.</div>
      </div>
      <p style={{ fontSize: 12, color: TEXT_DIM, marginBottom: 8 }}>
        {wasSubscriber ? "Fale com a gente pra renovar e receber um novo código." : "Fale com a gente pra assinar e receber seu código de liberação."}
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
        Falar no WhatsApp para {wasSubscriber ? "renovar" : "assinar"}
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
  const [authMode, setAuthMode] = useState("login"); // login | signup
  const [authEmail, setAuthEmail] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authConfirmPassword, setAuthConfirmPassword] = useState("");
  const [authLoading, setAuthLoading] = useState(false);
  const [authError, setAuthError] = useState("");
  const [unlockInput, setUnlockInput] = useState("");
  const [unlockError, setUnlockError] = useState("");

  const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

async function handleAuthSubmit() {
    setAuthError("");
    const email = authEmail.trim().toLowerCase();
    if (!email || !authPassword) {
      setAuthError("Preencha e-mail e senha.");
      return;
    }
    if (!EMAIL_PATTERN.test(email)) {
      setAuthError("Digite um e-mail válido (ex: nome@empresa.com).");
      return;
    }
    if (authPassword.length < 4) {
      setAuthError("A senha precisa ter pelo menos 4 caracteres.");
      return;
    }

    setAuthLoading(true);
    const existing = await loadUser(email);

    if (authMode === "signup") {
      if (authPassword !== authConfirmPassword) {
        setAuthError("As senhas não coincidem.");
        setAuthLoading(false);
        return;
      }
      if (existing) {
        setAuthError("Já existe uma conta com esse e-mail. Use \"Entrar\".");
        setAuthLoading(false);
        return;
      }
      const fresh = { password: authPassword, reportsUsed: 0, plan: "free" };
      await saveUser(email, fresh);
      setUser({ email, ...fresh });
      setAuthLoading(false);
      return;
    }

    // modo login
    if (!existing) {
      setAuthError("Não existe conta com esse e-mail. Use \"Criar conta\".");
      setAuthLoading(false);
      return;
    }
    if (existing.password !== authPassword) {
      setAuthError("Senha incorreta.");
      setAuthLoading(false);
      return;
    }
    setUser({ email, ...existing });
    setAuthLoading(false);
  }

  function handleLogout() {
    setUser(null);
    setAuthEmail("");
    setAuthPassword("");
    setAuthConfirmPassword("");
    setUnlockInput("");
    setUnlockError("");
  }

  async function handleUnlock() {
    setUnlockError("");
    if (unlockInput.trim().toUpperCase() === UNLOCK_CODE) {
      const expiresAt = Date.now() + 30 * 24 * 60 * 60 * 1000; // 30 dias a partir de agora
      const updated = { ...user, plan: "pago", expiresAt };
      setUser(updated);
      await saveUser(user.email, {
        password: updated.password,
        reportsUsed: updated.reportsUsed,
        plan: "pago",
        expiresAt,
      });
      setUnlockInput("");
    } else {
      setUnlockError("Código inválido.");
    }
  }

  function isPlanActive(u) {
    return !!u && u.plan === "pago" && !!u.expiresAt && Date.now() < u.expiresAt;
  }

  async function registerUsage() {
    if (!user || isPlanActive(user)) return; // assinatura ativa: uso ilimitado, não conta
    const updated = { ...user, reportsUsed: (user.reportsUsed || 0) + 1 };
    setUser(updated);
    await saveUser(user.email, {
      password: updated.password,
      reportsUsed: updated.reportsUsed,
      plan: updated.plan,
      expiresAt: updated.expiresAt,
    });
  }

  const [productsRaw, setProductsRaw] = useState("");
  const [columnMapping, setColumnMapping] = useState([]);

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

  const canRun = productDataRows.length > 0 && ncmDataRows.length > 0;

  function runAnalysis() {
    setError("");

    if (productDataRows.length === 0) {
      setError("Cole a lista de produtos antes de rodar.");
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
      const cst = useMapping ? (idx("cst") >= 0 ? r[idx("cst")] : "") : "";
      const cfop = useMapping ? (idx("cfop") >= 0 ? r[idx("cfop")] : "") : r[3];
      const csosn = useMapping ? (idx("csosn") >= 0 ? r[idx("csosn")] : "") : "";
      const ncmNorm = normalizeNCM(ncm);
      const impactada = ncmMap.has(ncmNorm);
      return {
        codigo: codigo || "",
        descricao: descricao || "",
        ncm: ncm || "",
        cst: cst || "",
        cfop: cfop || "",
        csosn: csosn || "",
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
      row(cell("titleBanner", "String", "DIAGNÓSTICO DE IMPACTO — REFORMA TRIBUTÁRIA", 1), 30) +
      row(cell("sub", "String", "Gerado em " + hoje, 1)) +
      row("") +
      row(cell("label", "String", "CLIENTE") + cell("value", "String", empresaLabel)) +
      row(cell("label", "String", "ESCRITÓRIO RESPONSÁVEL") + cell("value", "String", escritorioLabel)) +
      row("") +
      row(cell("indicHeader", "String", "INDICADOR") + cell("indicHeader", "String", "VALOR")) +
      row(cell("indicLabel", "String", "PRODUTOS ANALISADOS") + cell("vDark", "Number", result.totalAnalisado)) +
      row(cell("indicLabel", "String", "PRODUTOS IMPACTADOS") + cell("vRed", "Number", result.totalImpactado)) +
      row(cell("indicLabel", "String", "NCMS DISTINTAS IMPACTADAS") + cell("vBlue", "Number", result.ncmsDistintasImpactadas)) +
      row(cell("indicLabel", "String", "NCM ILEGÍVEL / NÃO IDENTIFICADO") + cell("vOrange", "Number", result.naoEncontrados)) +
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
        ["Código", "Descrição", "NCM", "CST", "CFOP", "CSOSN", "Impactado", "Motivo / mensagem"]
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
            cell(style, "String", d.cst) +
            cell(style, "String", d.cfop) +
            cell(style, "String", d.csosn) +
            cell(style, "String", d.impactado ? "Sim" : "Não") +
            cell(style, "String", d.mensagem)
          );
        })
        .join("");
      const bannerRow = row(cell(bannerStyle, "String", title.toUpperCase() + " (" + rows.length + ")", 8), 26);
      return (
        "<Worksheet ss:Name=\"" + esc(sheetName) + "\"><Table ss:DefaultColumnWidth=\"120\">" +
        "<Column ss:Width=\"90\"/><Column ss:Width=\"260\"/><Column ss:Width=\"90\"/><Column ss:Width=\"70\"/><Column ss:Width=\"70\"/><Column ss:Width=\"80\"/><Column ss:Width=\"80\"/><Column ss:Width=\"260\"/>" +
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

  async function downloadWord() {
    if (!result) return;
    const impacted = result.detail.filter((d) => d.impactado);
    const empresaLabel = empresa.trim() || "Cliente";
    const escritorioLabel = escritorio.trim();
    const hoje = new Date().toLocaleDateString("pt-BR");
    const avisoLegal = buildAvisoLegal(dataBaseNCM, fonteBaseNCM, hoje);

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

    const NAVY = "173A8A";
    const DARK = "1C2B33";
    const GRAY = "5C6584";
    const LIGHT_BLUE = "F2F5FC";
    const BORDER = "B9C3DA";

    const cellBorder = {
      top: { style: BorderStyle.SINGLE, size: 4, color: BORDER },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER },
      left: { style: BorderStyle.SINGLE, size: 4, color: BORDER },
      right: { style: BorderStyle.SINGLE, size: 4, color: BORDER },
    };

    const headerCell = new TableCell({
      width: { size: 3000, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: NAVY },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 90, bottom: 90, left: 120, right: 120 },
      borders: cellBorder,
      children: [
        new Paragraph({
          children: [new TextRun({ text: "NCM impactada", bold: true, color: "FFFFFF", size: 20, font: "Calibri" })],
        }),
      ],
    });

    const bodyRows = distinctNCMs.map(
      (ncm, i) =>
        new TableRow({
          children: [
            new TableCell({
              width: { size: 3000, type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: i % 2 === 0 ? "FFFFFF" : LIGHT_BLUE },
              margins: { top: 70, bottom: 70, left: 120, right: 120 },
              borders: cellBorder,
              children: [new Paragraph({ children: [new TextRun({ text: ncm, size: 20, font: "Consolas" })] })],
            }),
          ],
        })
    );

    const table = new Table({
      columnWidths: [3000],
      rows: [new TableRow({ tableHeader: true, children: [headerCell] }), ...bodyRows],
    });

    const letterheadParagraphs = escritorioLabel
      ? [
          new Paragraph({
            children: [new TextRun({ text: escritorioLabel.toUpperCase(), size: 18, color: GRAY, font: "Calibri" })],
            spacing: { after: 120 },
          }),
        ]
      : [];

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: { top: 900, bottom: 900, left: 1000, right: 1000 },
            },
          },
          footers: {
            default: new Footer({
              children: [
                new Paragraph({
                  border: { top: { style: BorderStyle.SINGLE, size: 4, color: BORDER, space: 6 } },
                  spacing: { before: 60 },
                  children: [
                    new TextRun({ text: avisoLegal, italics: true, size: 15, color: GRAY, font: "Calibri" }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 100 },
                  children: [
                    new TextRun({ text: "Página ", size: 15, color: GRAY, font: "Calibri" }),
                    new TextRun({ children: [PageNumber.CURRENT], size: 15, color: GRAY, font: "Calibri" }),
                    new TextRun({ text: " de ", size: 15, color: GRAY, font: "Calibri" }),
                    new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 15, color: GRAY, font: "Calibri" }),
                  ],
                }),
              ],
            }),
          },
          children: [
            ...letterheadParagraphs,
            new Paragraph({
              children: [new TextRun({ text: "Diagnóstico de Impacto — Reforma Tributária", bold: true, size: 40, color: NAVY, font: "Calibri" })],
              spacing: { after: 80 },
            }),
            new Paragraph({
              border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } },
              spacing: { after: 200 },
              children: [new TextRun({ text: " ", size: 4 })],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Cliente: ", bold: true, size: 21, color: DARK, font: "Calibri" }),
                new TextRun({ text: empresaLabel, size: 21, color: DARK, font: "Calibri" }),
              ],
              spacing: { after: 40 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Data: ", bold: true, size: 21, color: DARK, font: "Calibri" }),
                new TextRun({ text: hoje, size: 21, color: DARK, font: "Calibri" }),
              ],
              spacing: { after: 220 },
            }),
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              children: [new TextRun({ text: mensagem, size: 21, color: DARK, font: "Calibri" })],
              spacing: { after: 240 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Produtos analisados: ", bold: true, size: 20, color: NAVY, font: "Calibri" }),
                new TextRun({ text: `${result.totalAnalisado}    `, size: 20, color: DARK, font: "Calibri" }),
                new TextRun({ text: "Produtos impactados: ", bold: true, size: 20, color: NAVY, font: "Calibri" }),
                new TextRun({ text: `${result.totalImpactado}    `, size: 20, color: DARK, font: "Calibri" }),
                new TextRun({ text: "NCMs distintas impactadas: ", bold: true, size: 20, color: NAVY, font: "Calibri" }),
                new TextRun({ text: `${distinctNCMs.length}`, size: 20, color: DARK, font: "Calibri" }),
              ],
              spacing: { after: 260 },
            }),
            table,
          ],
        },
      ],
    });

    try {
      const blob = await Packer.toBlob(doc);
      const nomeArquivo = "ncm-impactados-" + empresaLabel.replace(/\s+/g, "-").toLowerCase() + ".docx";
      saveAs(blob, nomeArquivo);
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
        mode={authMode}
        setMode={(m) => { setAuthMode(m); setAuthError(""); }}
        email={authEmail}
        setEmail={setAuthEmail}
        password={authPassword}
        setPassword={setAuthPassword}
        confirmPassword={authConfirmPassword}
        setConfirmPassword={setAuthConfirmPassword}
        onSubmit={handleAuthSubmit}
        loading={authLoading}
        error={authError}
      />
    );
  }

  if (!isPlanActive(user) && (user.reportsUsed || 0) >= FREE_LIMIT) {
    return (
      <PaywallScreen
        email={user.email}
        wasSubscriber={!!user.expiresAt}
        onLogout={handleLogout}
        unlockInput={unlockInput}
        setUnlockInput={setUnlockInput}
        onUnlock={handleUnlock}
        unlockError={unlockError}
      />
    );
  }

  return (
    <div style={{ background: INK, minHeight: "100vh", color: "#EAEFFB", fontFamily: "'IBM Plex Sans', 'Segoe UI', sans-serif" }}>
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

      <div style={{ width: "100%", boxSizing: "border-box", padding: "36px 5vw 64px" }}>
        <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 12, marginBottom: 14 }}>
          <span style={{ fontSize: 12, color: TEXT_DIM }}>
            {user.email}
            {isPlanActive(user) ? (
              <span style={{ color: OK, marginLeft: 8 }}>
                · plano ativo (renova em {Math.max(1, Math.ceil((user.expiresAt - Date.now()) / (24 * 60 * 60 * 1000)))} dia{Math.ceil((user.expiresAt - Date.now()) / (24 * 60 * 60 * 1000)) === 1 ? "" : "s"})
              </span>
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
            <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 2 }}>Produtos do cliente</div>
            <div style={{ fontSize: 12, color: TEXT_DIM, marginBottom: 10 }}>
              Cole a tabela já limpa (código, descrição, NCM, CST/CFOP) — do jeito que vocês já preparam hoje
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
