"use client";
import { useState, useMemo, useRef, useCallback, useEffect } from "react";

/* ═══════════════════════════════════════════════
   UTILITÁRIOS
═══════════════════════════════════════════════ */
const fmt = v => Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const uid = () => `id_${Date.now()}_${Math.random().toString(36).slice(2,6)}`;

const parseValorBR = raw => {
  if (raw == null) return NaN;
  const s = String(raw).trim().replace(/[R$\s]/g,"").replace(/\.(?=\d{3})/g,"").replace(",",".");
  return parseFloat(s);
};

const parseDateBR = raw => {
  if (!raw) return null;
  if (raw instanceof Date) return raw;
  const s = String(raw).trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) { const [d,m,y]=s.split("/"); return new Date(+y,+m-1,+d); }
  if (/^\d{4}-\d{2}-\d{2}/.test(s))     { const [y,m,d]=s.split("-"); return new Date(+y,+m-1,+d); }
  const n = Number(s);
  if (!isNaN(n) && n > 40000) return new Date(Date.UTC(1899,11,30) + n*86400000);
  return null;
};

const fmtDate = d => {
  const dt = parseDateBR(d);
  if (!dt) return d ? String(d) : "";
  return `${String(dt.getDate()).padStart(2,"0")}/${String(dt.getMonth()+1).padStart(2,"0")}/${dt.getFullYear()}`;
};

const diffDias = (a, b) => {
  if (!a || !b) return 9999;
  return Math.round(Math.abs(a - b) / 86400000);
};

/* ═══════════════════════════════════════════════
   PARSER PDF (fiel ao parse_pdf do Python)
═══════════════════════════════════════════════ */
function parsePdfText(texto, nome) {
  const rows = [];
  for (const linha of texto.split("\n")) {
    const datas   = [...linha.matchAll(/\d{2}\/\d{2}\/\d{4}/g)].map(m => m[0]);
    const valores = [...linha.matchAll(/-?\d{1,3}(?:\.\d{3})*,\d{2}/g)].map(m => m[0]);
    if (!datas.length || !valores.length) continue;
    const dt    = parseDateBR(datas[0]);
    const valor = parseValorBR(valores[valores.length - 1]);
    if (!dt || isNaN(valor)) continue;
    const hist = linha
      .replace(/\d{2}\/\d{2}\/\d{4}/g, "")
      .replace(/-?\d{1,3}(?:\.\d{3})*,\d{2}/g, "")
      .replace(/\s{2,}/g, " ").trim();
    rows.push({ id: uid(), data: fmtDate(dt), _date: dt, valor, historico: hist, origem: nome });
  }
  if (!rows.length)
    throw new Error("Nenhum lançamento extraído do PDF. Verifique se o arquivo contém texto selecionável (não é escaneado).");
  return rows;
}

/* ═══════════════════════════════════════════════
   PARSERS DE EXTRATO
═══════════════════════════════════════════════ */
function normalizarExtrato(rows, colData, colValor, colHist, colCred, colDeb, origem) {
  return rows.map(r => {
    let valor;
    if (colCred && colDeb) {
      const c = parseValorBR(r[colCred]) || 0;
      const d = Math.abs(parseValorBR(r[colDeb]) || 0);
      valor = c - d;
    } else {
      valor = parseValorBR(r[colValor]);
    }
    const dt = parseDateBR(r[colData]);
    return { id: uid(), data: fmtDate(dt), _date: dt, valor, historico: String(r[colHist] || ""), origem };
  }).filter(r => r._date && !isNaN(r.valor));
}

function parseXlsxExtrato(wb, nome) {
  const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: true });
  if (!rows.length) throw new Error("Planilha vazia.");
  const cols = Object.keys(rows[0]);
  const find = kws => cols.find(c => kws.some(k => c.toLowerCase().includes(k))) || null;
  const dateCol  = find(["data","date"]);
  const valorCol = find(["valor","value","montante"]);
  const histCol  = find(["hist","descri","memo","lançamento","lancamento"]);
  const credCol  = find(["crédit","credit","entrada"]);
  const debCol   = find(["débit","debit","saída","saida"]);
  if (!dateCol) throw new Error(`Coluna de data não encontrada. Colunas: ${cols.join(", ")}`);
  if (!valorCol && !(credCol && debCol)) throw new Error(`Coluna de valor não encontrada. Colunas: ${cols.join(", ")}`);
  return normalizarExtrato(rows, dateCol, valorCol, histCol || "", credCol, debCol, nome);
}

function parseXlsxCpj(wb, config) {
  const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: true });
  if (!rows.length) throw new Error("Planilha CPJ vazia.");
  const cols = Object.keys(rows[0]);
  const { data: cData, valor: cValor, historico: cHist, conta: cConta, tipo: cTipo } = config.cpjColunas;
  for (const [campo, nome] of [["data",cData],["valor",cValor],["historico",cHist]]) {
    if (nome && !cols.includes(nome))
      throw new Error(`Coluna "${nome}" não encontrada no CPJ-3C.\nColunas disponíveis: ${cols.join(", ")}\nAjuste o mapeamento em Configurações.`);
  }
  return rows.map(r => {
    const dt = parseDateBR(r[cData]);
    return {
      id: uid(), data: fmtDate(dt), _date: dt,
      valor:     parseValorBR(r[cValor]),
      historico: String(r[cHist]  || ""),
      conta:     cConta ? String(r[cConta] || "") : "",
      tipo:      cTipo  ? String(r[cTipo]  || "") : "",
      origem:    "CPJ-3C",
    };
  }).filter(r => r._date && !isNaN(r.valor));
}

function parseOFX(text, nome) {
  return [...text.matchAll(/<STMTTRN>([\s\S]*?)<\/STMTTRN>/gi)].map(m => {
    const b = m[1];
    const get = tag => { const x = b.match(new RegExp(`<${tag}>([^<\r\n]+)`,"i")); return x ? x[1].trim() : ""; };
    const dtRaw = get("DTPOSTED") || get("DTUSER");
    const valor = parseFloat(get("TRNAMT"));
    if (!dtRaw || isNaN(valor)) return null;
    const dt = new Date(+dtRaw.slice(0,4), +dtRaw.slice(4,6)-1, +dtRaw.slice(6,8));
    return { id: uid(), data: fmtDate(dt), _date: dt, valor, historico: get("MEMO") || get("NAME") || "", origem: nome };
  }).filter(Boolean);
}

function parseCsvBB(text, nome) {
  return text.split("\n").slice(3).filter(l => l.trim()).map(l => {
    const p = l.split(";").map(s => s.replace(/"/g,"").trim());
    const dt = parseDateBR(p[0]); const valor = parseValorBR(p[2]);
    if (!dt || isNaN(valor)) return null;
    return { id: uid(), data: fmtDate(dt), _date: dt, valor, historico: p[1]||"", origem: nome };
  }).filter(Boolean);
}

function parseCsvBradesco(text, nome) {
  return text.split("\n").slice(2).filter(l => l.trim()).map(l => {
    const p = l.split(";").map(s => s.replace(/"/g,"").trim());
    const cred = parseValorBR(p[3])||0, deb = Math.abs(parseValorBR(p[4])||0);
    const valor = (cred||deb) ? cred-deb : parseValorBR(p[2]);
    const dt = parseDateBR(p[0]);
    if (!dt || isNaN(valor)) return null;
    return { id: uid(), data: fmtDate(dt), _date: dt, valor, historico: p[1]||"", origem: nome };
  }).filter(Boolean);
}

function parseCsvCEF(text, nome) {
  for (const skip of [8,9,10,7]) {
    const linhas = text.split("\n").slice(skip).filter(l => l.trim());
    if (!linhas.length) continue;
    const header = linhas[0].split(";").map(s => s.replace(/"/g,"").trim());
    if (!header.some(h => h.toLowerCase().includes("data"))) continue;
    const iData = header.findIndex(h => h.toLowerCase().includes("data"));
    const iVal  = header.findIndex(h => h.toLowerCase().includes("valor"));
    const iHist = header.findIndex(h => h.toLowerCase().includes("descri") || h.toLowerCase().includes("lança"));
    return linhas.slice(1).map(l => {
      const p = l.split(";").map(s => s.replace(/"/g,"").trim());
      const dt = parseDateBR(p[iData]); const valor = parseValorBR(p[iVal >= 0 ? iVal : 2]);
      if (!dt || isNaN(valor)) return null;
      return { id: uid(), data: fmtDate(dt), _date: dt, valor, historico: p[iHist >= 0 ? iHist : 1]||"", origem: nome };
    }).filter(Boolean);
  }
  throw new Error("Não foi possível detectar o cabeçalho no CSV da Caixa.");
}

/* ═══════════════════════════════════════════════
   MOTOR DE CONCILIAÇÃO
═══════════════════════════════════════════════ */
function conciliar(cpj, extrato, toleranciaDias) {
  const cpjItems = cpj.map(r => ({ ...r, _ok: false }));
  const extItems = extrato.map(r => ({ ...r, _ok: false }));
  const pares = [];
  for (const rc of cpjItems) {
    const cand = extItems.find(re =>
      !re._ok &&
      Math.abs(re.valor - rc.valor) < 0.01 &&
      diffDias(re._date, rc._date) <= toleranciaDias
    );
    if (cand) { pares.push({ cpj: rc, ext: cand }); rc._ok = true; cand._ok = true; }
  }
  return {
    conciliados: pares,
    soCpj:       cpjItems.filter(r => !r._ok),
    soExtrato:   extItems.filter(r => !r._ok),
  };
}

/* ═══════════════════════════════════════════════
   CORES POR TIPO
═══════════════════════════════════════════════ */
const TIPO_COR = t => ({
  "honorários":             {bg:"#E6F1FB",tx:"#0C447C",bd:"#B5D4F4"},
  "custas processuais":     {bg:"#FAEEDA",tx:"#854F0B",bd:"#FAC775"},
  "custas reembolsáveis":   {bg:"#FAEEDA",tx:"#854F0B",bd:"#FAC775"},
  "acordo judicial":        {bg:"#EAF3DE",tx:"#3B6D11",bd:"#C0DD97"},
  "levantamento de alvará": {bg:"#EEEDFE",tx:"#3C3489",bd:"#CECBF6"},
  "receita":                {bg:"#EAF3DE",tx:"#3B6D11",bd:"#C0DD97"},
}[(t||"").toLowerCase()] || {bg:"#F3F3F3",tx:"#555",bd:"#D5D5D5"});

/* ═══════════════════════════════════════════════
   COMPONENTE PRINCIPAL
═══════════════════════════════════════════════ */
const DEFAULTS = {
  cpjColunas: { data:"Data", valor:"Valor", historico:"Descrição da Movimentação", conta:"Conta", tipo:"Tipo" },
  tolerancia: 1,
  banco: "auto",
};

export default function App() {
  const [cpjRows,   setCpjRows]   = useState([]);
  const [extRows,   setExtRows]   = useState([]);
  const [selCpj,    setSelCpj]    = useState([]);
  const [selExt,    setSelExt]    = useState([]);
  const [matches,   setMatches]   = useState([]);
  const [resultado, setResultado] = useState(null);
  const [filtroExt, setFiltroExt] = useState("");
  const [filtroCpj, setFiltroCpj] = useState("");
  const [config,    setConfig]    = useState(DEFAULTS);
  const [showCfg,   setShowCfg]   = useState(false);
  const [logItems,  setLogItems]  = useState([]);
  const [loading,   setLoading]   = useState(false);
  const [tab,       setTab]       = useState("manual");
  const [xlsxReady, setXlsxReady] = useState(false);
  const [pdfReady,  setPdfReady]  = useState(false);

  const cpjFileRef = useRef();
  const extFileRef = useRef();

  /* ── carregar libs externas ── */
  useEffect(() => {
    if (window.XLSX) { setXlsxReady(true); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload = () => setXlsxReady(true);
    document.head.appendChild(s);
  }, []);

  useEffect(() => {
    if (window.pdfjsLib) { setPdfReady(true); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    s.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      setPdfReady(true);
    };
    document.head.appendChild(s);
  }, []);

  const addLog = useCallback((msg, tipo = "info") =>
    setLogItems(l => [...l, { msg, tipo, id: uid() }]), []);

  /* ── extrai texto do PDF ── */
  const extractPdfText = useCallback(async (buf) => {
    if (!pdfReady) throw new Error("PDF.js ainda carregando, aguarde.");
    const pdf = await window.pdfjsLib.getDocument({ data: buf }).promise;
    let full = "";
    for (let i = 1; i <= pdf.numPages; i++) {
      const page    = await pdf.getPage(i);
      const content = await page.getTextContent();
      const byY = {};
      for (const item of content.items) {
        const y = Math.round(item.transform[5]);
        (byY[y] = byY[y] || []).push(item.str);
      }
      for (const y of Object.keys(byY).sort((a, b) => b - a))
        full += byY[y].join(" ") + "\n";
    }
    return full;
  }, [pdfReady]);

  /* ── leitura genérica de arquivo ── */
  const lerArquivo = useCallback(async (file) => {
    const ext  = file.name.split(".").pop().toLowerCase();
    const nome = file.name.toLowerCase();
    const buf  = await file.arrayBuffer();

    if (ext === "ofx" || ext === "ofc") {
      const text = new TextDecoder("latin1").decode(buf);
      return { tipo: "extrato", rows: parseOFX(text, file.name) };
    }
    if (ext === "csv") {
      const text  = new TextDecoder("latin1").decode(buf);
      const banco = config.banco !== "auto" ? config.banco
        : nome.includes("bradesco") ? "bradesco"
        : nome.includes("caixa") || nome.includes("cef") ? "cef" : "bb";
      const rows = banco === "bradesco" ? parseCsvBradesco(text, file.name)
                 : banco === "cef"      ? parseCsvCEF(text, file.name)
                 :                        parseCsvBB(text, file.name);
      return { tipo: "extrato", rows };
    }
    if (ext === "pdf") {
      const texto = await extractPdfText(buf);
      return { tipo: "extrato", rows: parsePdfText(texto, file.name) };
    }
    if (ext === "xlsx" || ext === "xls") {
      if (!xlsxReady) throw new Error("SheetJS ainda carregando, aguarde.");
      const wb = window.XLSX.read(buf, { type: "array", cellDates: false });
      return { tipo: "xlsx", wb };
    }
    throw new Error(`Formato não suportado: .${ext}`);
  }, [config, xlsxReady, pdfReady, extractPdfText]);

  /* ── carregar CPJ ── */
  const onPickCpj = async e => {
    const file = e.target.files[0]; if (!file) return;
    try {
      setLoading(true);
      const res = await lerArquivo(file);
      if (res.tipo !== "xlsx") throw new Error("CPJ-3C deve ser um arquivo .xlsx ou .xls");
      const rows = parseXlsxCpj(res.wb, config);
      setCpjRows(rows);
      addLog(`✅ CPJ-3C: ${rows.length} lançamentos carregados de "${file.name}"`, "ok");
    } catch (err) {
      addLog("❌ " + err.message, "erro");
    } finally {
      setLoading(false); e.target.value = "";
    }
  };

  /* ── carregar extrato ── */
  const onPickExt = async e => {
    const file = e.target.files[0]; if (!file) return;
    try {
      setLoading(true);
      const res  = await lerArquivo(file);
      const rows = res.tipo === "xlsx" ? parseXlsxExtrato(res.wb, file.name) : res.rows;
      setExtRows(rows);
      const isPdf = file.name.toLowerCase().endsWith(".pdf");
      addLog(`✅ Extrato: ${rows.length} lançamentos de "${file.name}"`, "ok");
      if (isPdf) addLog("⚠️  PDF: confira os valores no resultado.", "warn");
    } catch (err) {
      addLog("❌ " + err.message, "erro");
      setTab("log");
    } finally {
      setLoading(false); e.target.value = "";
    }
  };

  /* ── conciliação automática ── */
  const rodarAuto = () => {
    if (!cpjRows.length) { addLog("❌ Carregue o CPJ-3C primeiro.", "erro"); return; }
    if (!extRows.length) { addLog("❌ Carregue o extrato primeiro.", "erro"); return; }
    setLoading(true);
    setTimeout(() => {
      try {
        addLog("═".repeat(44), "titulo");
        addLog("  CONCILIAÇÃO AUTOMÁTICA",  "titulo");
        addLog("═".repeat(44), "titulo");
        const r = conciliar(cpjRows, extRows, config.tolerancia);
        setResultado(r);
        addLog(`  Conciliados              : ${r.conciliados.length}`, r.conciliados.length > 0 ? "ok"   : "info");
        addLog(`  Só no extrato (pendente) : ${r.soExtrato.length}`,   r.soExtrato.length  > 0 ? "warn" : "ok");
        addLog(`  Só no CPJ    (pendente)  : ${r.soCpj.length}`,       r.soCpj.length      > 0 ? "warn" : "ok");
        addLog("─".repeat(44), "titulo");
        setTab("resultado");
      } catch (err) {
        addLog("❌ " + err.message, "erro");
      } finally {
        setLoading(false);
      }
    }, 50);
  };

  /* ── seleção manual ── */
  const togExt = id => setSelExt(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
  const togCpj = id => setSelCpj(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);

  const totSelExt = selExt.reduce((s, id) => s + (extRows.find(e => e.id === id)?.valor || 0), 0);
  const totSelCpj = selCpj.reduce((s, id) => s + (cpjRows.find(c => c.id === id)?.valor || 0), 0);
  const diff      = parseFloat((totSelExt - totSelCpj).toFixed(2));
  const diffOk    = diff === 0 && selExt.length > 0 && selCpj.length > 0;

  const confirmarMatch = () => {
    if (!selExt.length || !selCpj.length) return;
    const eItems = extRows.filter(e => selExt.includes(e.id));
    const cItems = cpjRows.filter(c => selCpj.includes(c.id));
    setMatches(m => [{ id: uid(), eItems, cItems, diff, ts: new Date().toLocaleTimeString("pt-BR") }, ...m]);
    setExtRows(r => r.filter(x => !selExt.includes(x.id)));
    setCpjRows(r => r.filter(x => !selCpj.includes(x.id)));
    setSelExt([]); setSelCpj([]);
  };

  const extFilt = useMemo(() => extRows.filter(e =>
    !filtroExt ||
    e.historico.toLowerCase().includes(filtroExt.toLowerCase()) ||
    String(e.valor).includes(filtroExt) ||
    e.data.includes(filtroExt)
  ), [extRows, filtroExt]);

  const cpjFilt = useMemo(() => cpjRows.filter(c =>
    !filtroCpj ||
    c.historico.toLowerCase().includes(filtroCpj.toLowerCase()) ||
    String(c.valor).includes(filtroCpj) ||
    c.data.includes(filtroCpj) ||
    (c.tipo || "").toLowerCase().includes(filtroCpj.toLowerCase())
  ), [cpjRows, filtroCpj]);

  /* ── helpers de estilo ── */
  const card   = { background:"#fff", border:"1px solid #E0E0E0", borderRadius:6, padding:"16px 18px", marginBottom:14 };
  const input  = { width:"100%", padding:"6px 8px", border:"1px solid #E0E0E0", borderRadius:4, fontSize:12, fontFamily:"inherit", background:"#FAFAFA", boxSizing:"border-box" };
  const btnP   = (bg="#1A56A0") => ({ padding:"7px 16px", background:bg, color:"#fff", border:"none", borderRadius:4, fontSize:13, fontWeight:600, cursor:"pointer", fontFamily:"inherit" });
  const btnO   = { padding:"5px 12px", background:"#fff", color:"#1A56A0", border:"1px solid #1A56A0", borderRadius:4, fontSize:12, cursor:"pointer", fontFamily:"inherit" };
  const tabBtn = active => ({ padding:"8px 16px", background: active ? "#1A56A0" : "#fff", color: active ? "#fff" : "#757575", border:"1px solid #E0E0E0", borderBottom: active ? "1px solid #fff" : "1px solid #E0E0E0", fontSize:12, cursor:"pointer", fontFamily:"inherit", marginBottom:-1 });
  const logCor = t => ({ ok:"#2E7D32", erro:"#C62828", warn:"#E65100", titulo:"#1A56A0" }[t] || "#757575");

  const Badge = ({ n, bg }) => n > 0
    ? <span style={{ marginLeft:6, background:bg, color:"#fff", borderRadius:10, padding:"1px 7px", fontSize:11 }}>{n}</span>
    : null;

  /* ── tabela de resultado genérica ── */
  const TabelaResultado = ({ label, cor, rows, colunas, celulas, zebraColor }) => (
    rows.length > 0 ? (
      <div style={{ marginBottom:16 }}>
        <div style={{ fontSize:12, fontWeight:700, color:cor, textTransform:"uppercase", letterSpacing:1, marginBottom:8 }}>{label}</div>
        <div style={{ border:"1px solid #E0E0E0", borderRadius:6, overflow:"hidden" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
            <thead>
              <tr style={{ background:"#F5F5F5" }}>
                {colunas.map(h => <th key={h} style={{ padding:"8px 10px", textAlign:"left", fontWeight:600, borderBottom:"1px solid #E0E0E0", whiteSpace:"nowrap" }}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={i} style={{ background: i%2===0 ? "#fff" : zebraColor }}>
                  {celulas(r).map((v, j) => (
                    <td key={j} style={{ padding:"7px 10px", borderBottom:"1px solid #F0F0F0", maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{v}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    ) : null
  );

  return (
    <div style={{ fontFamily:"'Segoe UI',system-ui,sans-serif", background:"#F5F5F5", minHeight:"100vh", fontSize:14 }}>

      {/* HEADER */}
      <div style={{ background:"#1A56A0", padding:"14px 20px", color:"#fff" }}>
        <div style={{ fontSize:17, fontWeight:700 }}>Conciliação Bancária</div>
        <div style={{ fontSize:12, color:"#BBDEFB", marginTop:3 }}>CPJ-3C · BB / Caixa / Bradesco · OFX / CSV / PDF / XLSX</div>
      </div>

      <div style={{ padding:"16px 18px" }}>

        {/* ENTRADAS */}
        <div style={card}>
          <div style={{ fontSize:13, fontWeight:700, marginBottom:12 }}>Arquivos de entrada</div>
          <div style={{ display:"flex", gap:12, flexWrap:"wrap" }}>
            {[
              { label:"Export CPJ-3C (.xlsx)",                          ref:cpjFileRef, accept:".xlsx,.xls",                    onChange:onPickCpj, count:cpjRows.length, cor:"#3B6D11" },
              { label:"Extrato bancário (.ofx / .csv / .pdf / .xlsx)",  ref:extFileRef, accept:".ofx,.ofc,.csv,.pdf,.xlsx,.xls", onChange:onPickExt, count:extRows.length, cor:"#185FA5" },
            ].map(f => (
              <div key={f.label} style={{ flex:1, minWidth:200 }}>
                <div style={{ fontSize:12, color:"#757575", marginBottom:5 }}>{f.label}</div>
                <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                  <button style={btnO} onClick={() => f.ref.current.click()}>Procurar…</button>
                  {f.count > 0 && <span style={{ fontSize:12, color:f.cor, fontWeight:500 }}>{f.count} lançamentos</span>}
                  <input type="file" accept={f.accept} ref={f.ref} onChange={f.onChange} style={{ display:"none" }}/>
                </div>
              </div>
            ))}
          </div>

          {/* config avançada */}
          <div style={{ marginTop:14 }}>
            <button onClick={() => setShowCfg(s => !s)}
              style={{ background:"none", border:"none", color:"#757575", fontSize:12, cursor:"pointer", padding:0 }}>
              {showCfg ? "▾" : "▸"} Configurações avançadas
            </button>
            {showCfg && (
              <div style={{ marginTop:10, display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"8px 14px" }}>
                {[
                  { label:"Coluna Data (CPJ)",      key:"data" },
                  { label:"Coluna Valor (CPJ)",     key:"valor" },
                  { label:"Coluna Descrição (CPJ)", key:"historico" },
                  { label:"Coluna Conta (CPJ)",     key:"conta" },
                  { label:"Coluna Tipo (CPJ)",      key:"tipo" },
                ].map(f => (
                  <div key={f.key}>
                    <div style={{ fontSize:11, color:"#757575", marginBottom:3 }}>{f.label}</div>
                    <input style={input} value={config.cpjColunas[f.key]}
                      onChange={e => setConfig(c => ({ ...c, cpjColunas: { ...c.cpjColunas, [f.key]: e.target.value }}))}/>
                  </div>
                ))}
                <div>
                  <div style={{ fontSize:11, color:"#757575", marginBottom:3 }}>Tolerância de datas (dias)</div>
                  <input type="number" min={0} max={7} style={{ ...input, width:72 }} value={config.tolerancia}
                    onChange={e => setConfig(c => ({ ...c, tolerancia: parseInt(e.target.value) || 0 }))}/>
                </div>
                <div>
                  <div style={{ fontSize:11, color:"#757575", marginBottom:3 }}>Banco (CSV)</div>
                  <select style={{ ...input, width:"auto" }} value={config.banco}
                    onChange={e => setConfig(c => ({ ...c, banco: e.target.value }))}>
                    <option value="auto">Auto (pelo nome do arquivo)</option>
                    <option value="bb">Banco do Brasil</option>
                    <option value="cef">Caixa Econômica Federal</option>
                    <option value="bradesco">Bradesco</option>
                  </select>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* ABAS */}
        <div style={{ display:"flex", borderBottom:"1px solid #E0E0E0", marginBottom:14 }}>
          {[
            ["manual",    "Conciliação Manual"],
            ["auto",      "Automática"],
            ["resultado", "Resultado"],
            ["log",       "Log"],
          ].map(([k, l]) => (
            <button key={k} style={tabBtn(tab === k)} onClick={() => setTab(k)}>
              {l}
              {k === "resultado" && resultado && <Badge n={resultado.conciliados.length} bg="#1A56A0"/>}
              {k === "manual"    && <Badge n={matches.length} bg="#2E7D32"/>}
            </button>
          ))}
          <div style={{ marginLeft:"auto", display:"flex", alignItems:"center", paddingBottom:1 }}>
            <button style={{ ...btnP(), fontSize:12, opacity: loading || !cpjRows.length || !extRows.length ? .5 : 1 }}
              onClick={rodarAuto} disabled={loading || !cpjRows.length || !extRows.length}>
              {loading ? "Processando…" : "⚡ Conciliar automaticamente"}
            </button>
          </div>
        </div>

        {/* ══════════ ABA MANUAL ══════════ */}
        {tab === "manual" && (
          <>
            {/* totais */}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:8, marginBottom:12 }}>
              {[
                { l:"Total extrato",    v:`R$ ${fmt(extRows.reduce((s,e)=>s+e.valor,0))}`, c:"#212121" },
                { l:"Selecionado ext.", v:`R$ ${fmt(totSelExt)}`, c:"#185FA5" },
                { l:"Selecionado CPJ",  v:`R$ ${fmt(totSelCpj)}`, c:"#3B6D11" },
                { l:"Diferença",
                  v: (selExt.length||selCpj.length) ? (diff===0 ? "R$ 0,00 ✓" : `R$ ${fmt(Math.abs(diff))}`) : "—",
                  c: diffOk ? "#2E7D32" : (selExt.length||selCpj.length)&&diff!==0 ? "#C62828" : "#212121" },
              ].map(x => (
                <div key={x.l} style={{ background:"#fff", border:"1px solid #E0E0E0", borderRadius:6, padding:"9px 12px" }}>
                  <div style={{ fontSize:11, color:"#757575", marginBottom:3 }}>{x.l}</div>
                  <div style={{ fontSize:15, fontWeight:500, color:x.c }}>{x.v}</div>
                </div>
              ))}
            </div>

            <div style={{ display:"flex", gap:8, marginBottom:12, alignItems:"center" }}>
              <span style={{ fontSize:12, color:"#757575", background:"#fff", border:"1px solid #E0E0E0", borderRadius:20, padding:"3px 10px" }}>
                {extRows.length} pendente{extRows.length !== 1 ? "s" : ""}
              </span>
              <button onClick={confirmarMatch} disabled={!selExt.length || !selCpj.length}
                style={{ marginLeft:"auto", ...btnP(diffOk ? "#2E7D32" : selExt.length&&selCpj.length ? "#1A56A0" : "#9E9E9E") }}>
                {diffOk ? "✓ Confirmar Match (R$ 0,00)" : `Confirmar Match${diff!==0&&selExt.length&&selCpj.length ? ` (Δ R$ ${fmt(Math.abs(diff))})` : ""}`}
              </button>
              <button onClick={() => { setSelExt([]); setSelCpj([]); }} style={btnO}>Limpar seleção</button>
            </div>

            {/* colunas */}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
              {/* extrato */}
              <div>
                <div style={{ fontSize:12, fontWeight:600, color:"#757575", textTransform:"uppercase", letterSpacing:1, marginBottom:8 }}>Extrato bancário</div>
                <input value={filtroExt} onChange={e => setFiltroExt(e.target.value)} placeholder="Filtrar…" style={{ ...input, marginBottom:8 }}/>
                <div style={{ display:"flex", flexDirection:"column", gap:5, maxHeight:420, overflowY:"auto" }}>
                  {extFilt.length === 0 && (
                    <div style={{ fontSize:12, color:"#9E9E9E", padding:"1rem 0", textAlign:"center" }}>
                      {extRows.length ? "Nenhum resultado" : "Nenhum lançamento carregado"}
                    </div>
                  )}
                  {extFilt.map(e => {
                    const sel = selExt.includes(e.id);
                    return (
                      <div key={e.id} onClick={() => togExt(e.id)}
                        style={{ background: sel?"#E6F1FB":"#fff", border:`1.5px solid ${sel?"#378ADD":"#E0E0E0"}`, borderRadius:6, padding:"9px 11px", cursor:"pointer" }}>
                        <div style={{ display:"flex", justifyContent:"space-between", gap:8 }}>
                          <div>
                            <div style={{ fontSize:11, color:"#9E9E9E", marginBottom:2 }}>{e.data} · {e.origem}</div>
                            <div style={{ fontSize:12, fontWeight:500, color:"#212121", lineHeight:1.35 }}>{e.historico || "—"}</div>
                          </div>
                          <div style={{ fontSize:14, fontWeight:600, color:"#185FA5", whiteSpace:"nowrap" }}>R$ {fmt(e.valor)}</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* CPJ */}
              <div>
                <div style={{ fontSize:12, fontWeight:600, color:"#757575", textTransform:"uppercase", letterSpacing:1, marginBottom:8 }}>Lançamentos CPJ-3C</div>
                <input value={filtroCpj} onChange={e => setFiltroCpj(e.target.value)} placeholder="Filtrar…" style={{ ...input, marginBottom:8 }}/>
                <div style={{ display:"flex", flexDirection:"column", gap:5, maxHeight:420, overflowY:"auto" }}>
                  {cpjFilt.length === 0 && (
                    <div style={{ fontSize:12, color:"#9E9E9E", padding:"1rem 0", textAlign:"center" }}>
                      {cpjRows.length ? "Nenhum resultado" : "Nenhum lançamento carregado"}
                    </div>
                  )}
                  {cpjFilt.map(c => {
                    const sel = selCpj.includes(c.id);
                    const tc  = TIPO_COR(c.tipo);
                    return (
                      <div key={c.id} onClick={() => togCpj(c.id)}
                        style={{ background: sel?"#EAF3DE":"#fff", border:`1.5px solid ${sel?"#639922":"#E0E0E0"}`, borderRadius:6, padding:"9px 11px", cursor:"pointer" }}>
                        <div style={{ display:"flex", justifyContent:"space-between", gap:8 }}>
                          <div style={{ flex:1, minWidth:0 }}>
                            <div style={{ fontSize:11, color:"#9E9E9E", marginBottom:2 }}>Venc. {c.data}</div>
                            <div style={{ fontSize:12, fontWeight:500, color:"#212121", marginBottom:4 }}>{c.historico || "—"}</div>
                            {c.tipo && <span style={{ fontSize:10, background:tc.bg, color:tc.tx, border:`0.5px solid ${tc.bd}`, borderRadius:4, padding:"2px 7px" }}>{c.tipo}</span>}
                          </div>
                          <div style={{ fontSize:14, fontWeight:600, color:"#3B6D11", whiteSpace:"nowrap" }}>R$ {fmt(c.valor)}</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* matches confirmados */}
            {matches.length > 0 && (
              <div style={{ marginTop:18 }}>
                <div style={{ fontSize:12, fontWeight:600, color:"#757575", textTransform:"uppercase", letterSpacing:1, marginBottom:10 }}>
                  Matches confirmados ({matches.length})
                </div>
                {matches.map(m => (
                  <div key={m.id} style={{ background: m.diff!==0?"#FAEEDA":"#EAF3DE", border:`1px solid ${m.diff!==0?"#FAC775":"#C0DD97"}`, borderRadius:6, padding:"11px 14px", marginBottom:8 }}>
                    <div style={{ fontSize:11, fontWeight:600, color: m.diff!==0?"#854F0B":"#3B6D11", marginBottom:7 }}>
                      {m.diff===0 ? "✓ Match exato" : `⚠ Δ R$ ${fmt(Math.abs(m.diff))}`} — {m.ts}
                    </div>
                    {m.eItems.map(e => <div key={e.id} style={{ fontSize:12, color:"#185FA5", marginBottom:2 }}>📥 {e.data} · {e.historico} · R$ {fmt(e.valor)}</div>)}
                    <div style={{ borderTop:`1px solid ${m.diff!==0?"#FAC775":"#C0DD97"}`, margin:"6px 0" }}/>
                    {m.cItems.map(c => <div key={c.id} style={{ fontSize:12, color:"#3B6D11", marginBottom:2 }}>📋 <strong>Baixar no CPJ:</strong> {c.historico} · {c.tipo} · R$ {fmt(c.valor)}</div>)}
                  </div>
                ))}
              </div>
            )}
          </>
        )}

        {/* ══════════ ABA RESULTADO ══════════ */}
        {tab === "resultado" && (
          <div>
            {!resultado ? (
              <div style={{ textAlign:"center", color:"#9E9E9E", padding:"3rem 0", fontSize:13 }}>
                Carregue os arquivos e clique em <strong>Conciliar automaticamente</strong>.
              </div>
            ) : (
              <>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:10, marginBottom:16 }}>
                  {[
                    { l:"Conciliados",   v:resultado.conciliados.length, c:"#2E7D32", bg:"#EAF3DE" },
                    { l:"Só no extrato", v:resultado.soExtrato.length,   c:"#E65100", bg:"#FFF3E0" },
                    { l:"Só no CPJ",     v:resultado.soCpj.length,       c:"#C62828", bg:"#FFEBEE" },
                  ].map(x => (
                    <div key={x.l} style={{ background:x.bg, border:"1px solid #E0E0E0", borderRadius:6, padding:"14px 16px", textAlign:"center" }}>
                      <div style={{ fontSize:28, fontWeight:700, color:x.c }}>{x.v}</div>
                      <div style={{ fontSize:12, color:"#757575", marginTop:3 }}>{x.l}</div>
                    </div>
                  ))}
                </div>

                <TabelaResultado
                  label="✓ Conciliados" cor="#2E7D32" rows={resultado.conciliados} zebraColor="#F9F9F9"
                  colunas={["Data CPJ","Data Extrato","Valor (R$)","Histórico CPJ","Histórico Extrato"]}
                  celulas={p => [p.cpj.data, p.ext.data, `R$ ${fmt(p.cpj.valor)}`, p.cpj.historico, p.ext.historico]}
                />
                <TabelaResultado
                  label="⚠ Só no CPJ (não encontrado no extrato)" cor="#C62828" rows={resultado.soCpj} zebraColor="#FFF8F8"
                  colunas={["Data","Valor (R$)","Histórico","Tipo"]}
                  celulas={r => [r.data, `R$ ${fmt(r.valor)}`, r.historico, r.tipo]}
                />
                <TabelaResultado
                  label="⚠ Só no extrato (não lançado no CPJ)" cor="#E65100" rows={resultado.soExtrato} zebraColor="#FFFAF5"
                  colunas={["Data","Valor (R$)","Histórico","Origem"]}
                  celulas={r => [r.data, `R$ ${fmt(r.valor)}`, r.historico, r.origem]}
                />
              </>
            )}
          </div>
        )}

        {/* ══════════ ABA LOG ══════════ */}
        {tab === "log" && (
          <div style={card}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}>
              <span style={{ fontSize:13, fontWeight:700 }}>Log de execução</span>
              <button onClick={() => setLogItems([])} style={{ background:"none", border:"none", color:"#757575", fontSize:12, cursor:"pointer" }}>Limpar</button>
            </div>
            <div style={{ background:"#FAFAFA", border:"1px solid #E0E0E0", borderRadius:4, padding:"10px 12px", height:280, overflowY:"auto", fontFamily:"Consolas,'Courier New',monospace", fontSize:12, lineHeight:1.7 }}>
              {logItems.length === 0 && <span style={{ color:"#9E9E9E" }}>Nenhum evento ainda.</span>}
              {logItems.map(l => <div key={l.id} style={{ color: logCor(l.tipo) }}>{l.msg}</div>)}
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
