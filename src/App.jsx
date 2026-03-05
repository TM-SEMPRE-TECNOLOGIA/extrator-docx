import React, { useCallback, useMemo, useRef, useState } from "react";
import { AnimatePresence, motion } from "framer-motion";
import {
  CloudUpload,
  FileText,
  Loader2,
  Download,
  CheckCircle2,
  AlertTriangle,
  Info,
  Sigma,
  Moon,
  Sun,
} from "lucide-react";
import * as XLSX from "xlsx";
import JSZip from "jszip";

/**
 * Extrator de Itens DOCX - Ocean Breeze Design
 * Redesigned com foco em simplicidade e hierarquia visual
 */

const CODE_RE = /^\s*\d+(?:\.\d+)?\s*$/;

/** @typedef {{ codigo: string; descricao: string; quantidade_raw: string; quantidade: number; origem?: string }} Item */

function cn(...xs) {
  return xs.filter(Boolean).join(" ");
}

function norm(s) {
  return (s ?? "").replace(/\u00A0/g, " ").trim();
}

function fmtInt(n) {
  try {
    return new Intl.NumberFormat("pt-BR").format(n);
  } catch {
    return String(n);
  }
}

function fmtQty(q) {
  if (!Number.isFinite(q)) return "";
  try {
    return new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 6 }).format(q);
  } catch {
    return String(q);
  }
}

function parsePtNumber(s) {
  const t = norm(s);
  if (!t) return NaN;
  const cleaned = t.replace(/\./g, "").replace(/,/g, ".");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : NaN;
}

function safeBaseName(name) {
  const base = String(name || "documento")
    .replace(/\.docx$/i, "")
    .replace(/[^a-zA-Z0-9\-_. ]+/g, "_")
    .trim();
  return base || "documento";
}

function makeName(docxName, prefix, ext) {
  return `${prefix}_${safeBaseName(docxName)}.${ext}`;
}

function isInIframe() {
  try {
    return window.self !== window.top;
  } catch {
    return true;
  }
}

function fallbackDownload({ filename, mime, data }) {
  const blob = data instanceof Blob ? data : new Blob([data], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

async function saveFile({ filename, mime, data, hint }) {
  const inIframe = typeof window !== "undefined" && isInIframe();

  if (!inIframe && typeof window !== "undefined" && window.showSaveFilePicker) {
    try {
      const ext = filename.split(".").pop() || "";
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [
          {
            description: hint?.toUpperCase() || "Arquivo",
            accept: { [mime]: [`.${ext}`] },
          },
        ],
      });
      const writable = await handle.createWritable();
      const blob = data instanceof Blob ? data : new Blob([data], { type: mime });
      await writable.write(blob);
      await writable.close();
      return;
    } catch (e) {
      console.warn("showSaveFilePicker falhou; usando fallback.", e);
    }
  }

  fallbackDownload({ filename, mime, data });
}

function buildXlsx(items, filenameBase) {
  const rows = items.map((it) => ({
    Codigo: it.codigo,
    Descricao: it.descricao,
    Quantidade: Number.isFinite(it.quantidade) ? it.quantidade : it.quantidade_raw,
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Itens");
  ws["!cols"] = [{ wch: 16 }, { wch: 56 }, { wch: 14 }];

  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  void saveFile({
    filename: `${filenameBase}.xlsx`,
    mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    data: blob,
    hint: "xlsx",
  });
}

function buildLogText({ fileName, statusLines, meta, items, aggregated }) {
  const now = new Date();
  const header = [
    "TM Sempre Tecnologia - Extrator de Itens DOCX",
    `Data: ${now.toLocaleString("pt-BR")}`,
    `Arquivo: ${fileName || "(nenhum)"}`,
    "",
    "--- Status ---",
    ...statusLines.map((s) => `- ${s}`),
    "",
    "--- Metricas ---",
    `Tabelas totais no DOCX: ${meta?.tables_total ?? 0}`,
    `Tabelas identificadas como 'Itens': ${meta?.itens_tables ?? 0}`,
    `Linhas extraidas (total): ${meta?.rows_extracted ?? 0}`,
    `Linhas ignoradas: ${meta?.rows_ignored ?? 0}`,
    aggregated ? `Itens unicos (somados): ${aggregated.length}` : "",
    "",
  ].filter(Boolean);

  const ignored = (meta?.ignored_details ?? []).slice(0, 300);
  const ignoredBlock = ignored.length
    ? [
      "--- Detalhes ignorados (amostra) ---",
      ...ignored.map((d) => `- ${d}`),
      ignored.length < (meta?.ignored_details ?? []).length
        ? `... (${(meta?.ignored_details ?? []).length - ignored.length} a mais)`
        : "",
      "",
    ].filter(Boolean)
    : [];

  const sample = (items ?? []).slice(0, 20).map(
    (it, i) =>
      `${String(i + 1).padStart(2, "0")}. ${it.codigo} | ${it.descricao} | qtd=${it.quantidade_raw}`
  );

  return [...header, ...ignoredBlock, "--- Saida (amostra) ---", ...sample].join("\n");
}

function xmlTextOf(node) {
  const ts = node.getElementsByTagName("w:t");
  let out = "";
  for (let i = 0; i < ts.length; i++) out += ts[i].textContent ?? "";
  return norm(out);
}

function pickQuantityFromRow(cellsText) {
  if (cellsText.length >= 3) {
    const q = norm(cellsText[2]);
    if (q && q.toUpperCase() !== "#N/D") return q;
  }
  const joined = cellsText.join(" ");
  const m = joined.match(/(\d{1,3}(?:\.\d{3})*,\d+|\d+,\d+|\d+)/);
  return m?.[1] ?? "";
}

async function extractItemsFromDocx(file) {
  const buf = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);

  const docXml = await zip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("Nao foi possivel ler word/document.xml do DOCX.");

  const parser = new DOMParser();
  const xml = parser.parseFromString(docXml, "application/xml");
  const perr = xml.getElementsByTagName("parsererror");
  if (perr?.length) throw new Error("Falha ao interpretar o XML do DOCX.");

  const tables = Array.from(xml.getElementsByTagName("w:tbl"));

  /** @type {Item[]} */
  const results = [];
  /** @type {string[]} */
  const ignored = [];
  let itensTables = 0;

  tables.forEach((tbl, tIndex) => {
    const rows = Array.from(tbl.getElementsByTagName("w:tr"));
    if (!rows.length) return;

    const headerCells = Array.from(rows[0].getElementsByTagName("w:tc"));
    const headerTexts = headerCells.map((tc) => xmlTextOf(tc));
    const isItens = headerTexts.some((t) => norm(t).toLowerCase() === "itens");
    if (!isItens) return;

    itensTables += 1;

    rows.slice(1).forEach((tr, rOffset) => {
      const rNumber = rOffset + 2;
      const tNumber = tIndex + 1;

      const tcs = Array.from(tr.getElementsByTagName("w:tc"));
      if (!tcs.length) {
        ignored.push(`T${tNumber} L${rNumber}: skip_empty_row`);
        return;
      }

      const cellsText = tcs.map((tc) => xmlTextOf(tc));
      const code = norm(cellsText[0] ?? "");
      const desc = norm(cellsText[1] ?? "");

      if (!code || code.toUpperCase() === "#N/D") {
        ignored.push(`T${tNumber} L${rNumber}: skip_code_empty_or_ND`);
        return;
      }
      if (!CODE_RE.test(code)) {
        ignored.push(`T${tNumber} L${rNumber}: skip_code_invalid ${code}`);
        return;
      }

      const qtyRaw = pickQuantityFromRow(cellsText);
      if (!qtyRaw || qtyRaw.toUpperCase() === "#N/D") {
        ignored.push(`T${tNumber} L${rNumber}: skip_qty_empty_or_ND ${code}`);
        return;
      }

      const qty = parsePtNumber(qtyRaw);

      results.push({
        codigo: code,
        descricao: desc,
        quantidade_raw: qtyRaw,
        quantidade: qty,
        origem: `T${tNumber}/L${rNumber}`,
      });
    });
  });

  const meta = {
    tables_total: tables.length,
    itens_tables: itensTables,
    rows_extracted: results.length,
    rows_ignored: ignored.length,
    ignored_details: ignored,
  };

  return { items: results, meta };
}

function aggregateItems(items, rule) {
  /** @type {Map<string, {codigo:string, descricao:string, quantidade:number}>} */
  const map = new Map();

  const keyOf = (it) => {
    if (rule === "code_only") return norm(it.codigo).toLowerCase();
    if (rule === "desc_only") return norm(it.descricao).toLowerCase();
    return `${norm(it.codigo).toLowerCase()}|${norm(it.descricao).toLowerCase()}`;
  };

  items.forEach((it) => {
    const key = keyOf(it);
    const prev = map.get(key);
    const q = Number.isFinite(it.quantidade) ? it.quantidade : parsePtNumber(it.quantidade_raw);
    const safeQ = Number.isFinite(q) ? q : 0;

    if (!prev) {
      map.set(key, {
        codigo: rule === "desc_only" ? "" : it.codigo,
        descricao: rule === "code_only" ? "" : it.descricao,
        quantidade: safeQ,
      });
    } else {
      prev.quantidade += safeQ;
    }
  });

  return Array.from(map.values());
}

function Badge({ kind, icon, children }) {
  const cls =
    kind === "idle"
      ? "tm-badge tm-badge--idle"
      : kind === "work"
        ? "tm-badge tm-badge--work"
        : kind === "ok"
          ? "tm-badge tm-badge--ok"
          : "tm-badge tm-badge--err";

  return (
    <span className={cls}>
      <span className="tm-badge__icon">{icon}</span>
      <span>{children}</span>
    </span>
  );
}

function StatCard({ label, value, sub }) {
  return (
    <div className="tm-stat">
      <div className="tm-stat__label">{label}</div>
      <div className="tm-stat__value">{value}</div>
      {sub ? <div className="tm-stat__sub">{sub}</div> : null}
    </div>
  );
}

function Section({ title, desc, right, children }) {
  return (
    <section className="tm-panel">
      <div className="tm-panel__header">
        <div className="tm-panel__left">
          <h2 className="tm-panel__title">{title}</h2>
          {desc ? <div className="tm-panel__desc">{desc}</div> : null}
        </div>
        {right ? <div className="tm-panel__right">{right}</div> : null}
      </div>
      <div className="tm-panel__body">{children}</div>
    </section>
  );
}

export default function AppExtratorDocx() {
  const inputRef = useRef(null);
  const [drag, setDrag] = useState(false);
  const [theme, setTheme] = useState("light");

  const [file, setFile] = useState(null);
  const [phase, setPhase] = useState("idle");
  const [statusText, setStatusText] = useState("Envie um .docx para iniciar.");
  const [lines, setLines] = useState(["Pronto para receber arquivo."]);

  const [items, setItems] = useState(/** @type {Item[]} */([]));
  const [meta, setMeta] = useState(null);
  const [logText, setLogText] = useState("");

  const [aggRule, setAggRule] = useState("code_desc");
  const [aggPhase, setAggPhase] = useState("idle");
  const [aggText, setAggText] = useState("Escolha a regra e gere a planilha consolidada.");
  const [aggLines, setAggLines] = useState(["Aguardando acao."]);
  const [aggItems, setAggItems] = useState([]);

  const canProcess = !!file && phase !== "work";
  const canAggregate = phase === "ok" && items.length > 0 && aggPhase !== "work";

  const onPick = useCallback(() => inputRef.current?.click(), []);

  const onFileSelected = useCallback((f) => {
    if (!f) return;

    if (!String(f.name).toLowerCase().endsWith(".docx")) {
      setPhase("err");
      setStatusText("Arquivo invalido. Envie um .docx.");
      setLines(["O arquivo selecionado nao e .docx."]);
      return;
    }

    setFile(f);
    setPhase("idle");
    setStatusText("Arquivo carregado. Pronto para processar.");
    setLines(["Arquivo selecionado", "Clique em EXTRAIR ITENS"]);

    setItems([]);
    setMeta(null);
    setLogText("");

    setAggPhase("idle");
    setAggText("Escolha a regra e gere a planilha consolidada.");
    setAggLines(["Aguardando acao."]);
    setAggItems([]);
  }, []);

  const onInputChange = useCallback(
    (e) => {
      const f = e.target.files?.[0];
      onFileSelected(f);
    },
    [onFileSelected]
  );

  const onDrop = useCallback(
    (e) => {
      e.preventDefault();
      e.stopPropagation();
      setDrag(false);
      const f = e.dataTransfer?.files?.[0];
      onFileSelected(f);
    },
    [onFileSelected]
  );

  const processDoc = useCallback(async () => {
    if (!file) return;

    setPhase("work");
    setStatusText("Processando documento...");
    setLines(["Lendo tabelas", "Extraindo codigos e quantidades", "Preparando saida"]);

    const step = async (ms, msg) => {
      await new Promise((r) => setTimeout(r, ms));
      setLines((prev) => [...prev.slice(0, 1), msg]);
    };

    try {
      await step(200, "Abrindo arquivo e lendo XML...");
      const { items: extracted, meta: m } = await extractItemsFromDocx(file);

      await step(200, `Tabelas: ${m.tables_total} | Itens: ${m.itens_tables}`);
      await step(200, `Validas: ${m.rows_extracted} | Ignoradas: ${m.rows_ignored}`);

      setMeta(m);

      if (!extracted.length) {
        setPhase("err");
        setStatusText("Nenhuma linha valida foi encontrada nas tabelas 'Itens'.");
        setLines([
          "Nenhum item extraido.",
          "Verifique se existe uma tabela com cabecalho 'Itens' (1a linha).",
        ]);

        const t = buildLogText({
          fileName: file.name,
          statusLines: ["Sem dados"],
          meta: m,
          items: [],
        });
        setLogText(t);
        return;
      }

      setItems(extracted);

      const t = buildLogText({
        fileName: file.name,
        statusLines: ["Extracao concluida"],
        meta: m,
        items: extracted,
      });
      setLogText(t);

      setPhase("ok");
      setStatusText("Extracao concluida!");
      setLines([
        `Itens encontrados: ${fmtInt(extracted.length)}`,
        "Downloads disponiveis abaixo",
      ]);
    } catch (err) {
      setPhase("err");
      setStatusText("Erro ao processar o DOCX.");
      setLines([
        String(err?.message ?? err),
        "Dica: tente exportar o Word novamente (DOCX padrao).",
      ]);

      const m = meta ?? {
        tables_total: 0,
        itens_tables: 0,
        rows_extracted: 0,
        rows_ignored: 0,
        ignored_details: [],
      };
      const t = buildLogText({
        fileName: file?.name,
        statusLines: ["Erro"],
        meta: m,
        items: [],
      });
      setLogText(t);
    }
  }, [file, meta]);

  const downloadBruto = useCallback(() => {
    if (!items.length) return;
    buildXlsx(items, `itens_bruto_${safeBaseName(file?.name)}`);
  }, [items, file]);

  const downloadSomado = useCallback(() => {
    if (!aggItems.length) return;
    const rows = aggItems.map((x) => ({
      codigo: x.codigo,
      descricao: x.descricao,
      quantidade_raw: fmtQty(x.quantidade),
      quantidade: x.quantidade,
    }));
    buildXlsx(rows, `itens_somados_${safeBaseName(file?.name)}`);
  }, [aggItems, file]);

  const downloadLog = useCallback(() => {
    if (!logText) return;
    void saveFile({
      filename: makeName(file?.name, "itens_log", "txt"),
      mime: "text/plain;charset=utf-8",
      data: new Blob([logText], { type: "text/plain;charset=utf-8" }),
      hint: "txt",
    });
  }, [logText, file]);

  const doAggregate = useCallback(async () => {
    if (!canAggregate) return;

    setAggPhase("work");
    setAggText("Consolidando itens...");
    setAggLines(["Definindo chave", "Somando quantidades", "Preparando saida"]);

    try {
      await new Promise((r) => setTimeout(r, 200));
      const ag = aggregateItems(items, aggRule);
      setAggItems(ag);

      setAggPhase("ok");
      const keyLabel =
        aggRule === "code_only"
          ? "Apenas Codigo"
          : aggRule === "desc_only"
            ? "Apenas Descricao"
            : "Codigo + Descricao";
      setAggText("Consolidacao concluida!");
      setAggLines([
        `Regra: ${keyLabel}`,
        `Itens unicos: ${fmtInt(ag.length)}`,
        "Planilha pronta para download",
      ]);

      if (file) {
        const extra = `\n\n--- Consolidado (amostra) ---\n${ag
          .slice(0, 10)
          .map(
            (x, i) =>
              `${String(i + 1).padStart(2, "0")}. ${x.codigo} | ${x.descricao} | qtd=${fmtQty(
                x.quantidade
              )}`
          )
          .join("\n")}`;
        setLogText((prev) => (prev ? prev + extra : extra));
      }
    } catch (err) {
      setAggPhase("err");
      setAggText("Erro na consolidacao.");
      setAggLines([String(err?.message ?? err)]);
    }
  }, [aggRule, canAggregate, items, file]);

  const badge = useMemo(() => {
    if (phase === "work") return { kind: "work", icon: <Loader2 size={16} className="tm-spin" /> };
    if (phase === "ok") return { kind: "ok", icon: <CheckCircle2 size={16} /> };
    if (phase === "err") return { kind: "err", icon: <AlertTriangle size={16} /> };
    return { kind: "idle", icon: <Info size={16} /> };
  }, [phase]);

  const aggBadge = useMemo(() => {
    if (aggPhase === "work") return { kind: "work", icon: <Loader2 size={16} className="tm-spin" /> };
    if (aggPhase === "ok") return { kind: "ok", icon: <Sigma size={16} /> };
    if (aggPhase === "err") return { kind: "err", icon: <AlertTriangle size={16} /> };
    return { kind: "idle", icon: <Info size={16} /> };
  }, [aggPhase]);

  const toggleTheme = useCallback(() => {
    setTheme((t) => (t === "dark" ? "light" : "dark"));
  }, []);

  return (
    <div className={cn("tm-root", theme === "dark" && "dark")}>
      <header className="tm-header">
        <div className="tm-brand">
          <div className="tm-brand__mark">TM</div>
          <div>
            <div className="tm-brand__name">TM Sempre Tecnologia</div>
            <div className="tm-brand__sub">Extrator de Itens DOCX</div>
          </div>
        </div>

        <div className="tm-header__right">
          <span className="tm-pill tm-pill--online">
            <span className="tm-dot" />
            Online
          </span>
          <button type="button" className="tm-btn tm-btn--outline" onClick={toggleTheme} aria-label="Alternar tema">
            {theme === "dark" ? <Sun size={16} /> : <Moon size={16} />}
          </button>
        </div>
      </header>

      <main className="tm-main">
        {/* HERO - Só aparece no idle */}
        {phase === "idle" && !file && (
          <div style={{ textAlign: "center", padding: "48px 20px 32px", marginBottom: "24px" }}>
            <h1 style={{
              fontSize: "clamp(28px, 5vw, 40px)",
              fontWeight: "800",
              marginBottom: "16px",
              letterSpacing: "-0.5px",
              lineHeight: "1.1"
            }}>
              Extraia itens do seu DOCX em segundos
            </h1>
            <p style={{
              fontSize: "16px",
              color: "var(--TM-muted-foreground)",
              maxWidth: "560px",
              margin: "0 auto",
              lineHeight: "1.6"
            }}>
              Arraste seu relatório DOCX aqui. Identificamos as tabelas de itens e geramos o Excel para você.
            </p>
          </div>
        )}

        {/* UPLOAD SECTION */}
        <Section
          title="Upload do Arquivo"
          desc="Arraste seu DOCX aqui ou clique para selecionar"
          right={
            <Badge kind={badge.kind} icon={badge.icon}>
              {phase === "idle" ? "Aguardando" : phase === "work" ? "Processando" : phase === "ok" ? "Concluído" : "Erro"}
            </Badge>
          }
        >
          <div
            onDragEnter={(e) => {
              e.preventDefault();
              e.stopPropagation();
              setDrag(true);
            }}
            onDragOver={(e) => {
              e.preventDefault();
              e.stopPropagation();
              setDrag(true);
            }}
            onDragLeave={(e) => {
              e.preventDefault();
              e.stopPropagation();
              setDrag(false);
            }}
            onDrop={onDrop}
            className={cn("tm-drop", drag && "tm-drop--active")}
            role="button"
            tabIndex={0}
            onClick={onPick}
          >
            <div className="tm-drop__row">
              <div className="tm-drop__left">
                <div className="tm-drop__icon">DOCX</div>
                <div>
                  <div className="tm-drop__title">Arraste seu DOCX aqui</div>
                  <div className="tm-drop__hint">ou clique para selecionar · Max: 20MB</div>
                </div>
              </div>
            </div>
          </div>

          <input ref={inputRef} type="file" accept=".docx" hidden onChange={onInputChange} />

          <div className="tm-actions" style={{ marginTop: "16px" }}>
            <button type="button" onClick={processDoc} disabled={!canProcess} className="tm-btn tm-btn--primary">
              {phase === "work" ? <Loader2 size={16} className="tm-spin" /> : <FileText size={16} />}
              Extrair Itens
            </button>
            <button type="button" onClick={onPick} disabled={phase === "work"} className="tm-btn tm-btn--outline">
              <CloudUpload size={16} />
              Novo Arquivo
            </button>
          </div>

          {file && (
            <div className="tm-status" style={{ marginTop: "14px" }}>
              <div className="tm-status__top">
                <span>{statusText}</span>
                <span className="tm-status__file">{file.name}</span>
              </div>
              <div className="tm-status__lines">
                {lines.map((l, i) => (
                  <span key={i}>{l}</span>
                ))}
              </div>
            </div>
          )}
        </Section>

        {/* RESULTADO - Só aparece quando concluído */}
        <AnimatePresence>
          {phase === "ok" && (
            <motion.div
              initial={{ opacity: 0, y: 8 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 8 }}
              transition={{ duration: 0.3 }}
            >
              <Section
                title="Resumo da Extração"
                desc={`${fmtInt(items.length)} itens encontrados e prontos para download`}
              >
                <div className="tm-stats">
                  <StatCard label="Itens extraídos" value={fmtInt(items.length)} />
                  <StatCard label="Tabelas processadas" value={meta ? `${meta.itens_tables}/${meta.tables_total}` : "-"} />
                </div>

                <div className="tm-actions" style={{ marginTop: "16px" }}>
                  <button type="button" onClick={downloadBruto} className="tm-btn tm-btn--primary">
                    <Download size={16} />
                    Baixar Excel
                  </button>
                  <button type="button" onClick={downloadLog} disabled={!logText} className="tm-btn tm-btn--outline">
                    <Download size={16} />
                    Baixar Log
                  </button>
                </div>
              </Section>

              {/* CONSOLIDAÇÃO */}
              {items.length > 0 && (
                <Section
                  title="Consolidar Itens Repetidos"
                  desc="Agrupe itens iguais e some as quantidades"
                  right={
                    <Badge kind={aggBadge.kind} icon={aggBadge.icon}>
                      {aggPhase === "idle" ? "Pronto" : aggPhase === "work" ? "Somando" : aggPhase === "ok" ? "Concluído" : "Erro"}
                    </Badge>
                  }
                >
                  <div className="tm-panel__desc" style={{ marginBottom: "12px" }}>Regra de agrupamento:</div>
                  <div className="tm-rule-grid">
                    {[
                      { v: "code_desc", label: "Código + Descrição" },
                      { v: "code_only", label: "Apenas Código" },
                      { v: "desc_only", label: "Apenas Descrição" },
                    ].map((opt) => (
                      <label key={opt.v} className={cn("tm-rule", aggRule === opt.v && "tm-rule--active")}>
                        <input
                          type="radio"
                          name="rule"
                          value={opt.v}
                          checked={aggRule === opt.v}
                          onChange={() => setAggRule(opt.v)}
                        />
                        {opt.label}
                      </label>
                    ))}
                  </div>

                  <div className="tm-actions" style={{ marginTop: "16px" }}>
                    <button type="button" onClick={doAggregate} disabled={!canAggregate} className="tm-btn tm-btn--primary">
                      {aggPhase === "work" ? <Loader2 size={16} className="tm-spin" /> : <Sigma size={16} />}
                      Agrupar e Baixar
                    </button>
                    {aggPhase === "ok" && (
                      <button
                        type="button"
                        onClick={downloadSomado}
                        disabled={aggItems.length === 0}
                        className="tm-btn tm-btn--outline"
                      >
                        <Download size={16} />
                        Baixar Consolidado
                      </button>
                    )}
                  </div>

                  {aggPhase !== "idle" && (
                    <div className="tm-status" style={{ marginTop: "14px" }}>
                      <div className="tm-status__top">
                        <span>{aggText}</span>
                      </div>
                      <div className="tm-status__lines">
                        {aggLines.map((l, i) => (
                          <span key={i}>{l}</span>
                        ))}
                      </div>
                    </div>
                  )}

                  {aggPhase === "ok" && aggItems.length > 0 && (
                    <div style={{ marginTop: "16px" }}>
                      <div className="tm-mini-title" style={{ marginBottom: "8px" }}>
                        Prévia dos Resultados (primeiros 5)
                      </div>
                      <div className="tm-preview">
                        {aggItems.slice(0, 5).map((x, idx) => (
                          <div key={idx} className="tm-preview__item">
                            <div className="tm-preview__meta">
                              <div className="tm-preview__code">{x.codigo || "(sem código)"}</div>
                              <div className="tm-preview__desc">{x.descricao || "(sem descrição)"}</div>
                            </div>
                            <div className="tm-preview__qty">{fmtQty(x.quantidade)}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </Section>
              )}
            </motion.div>
          )}
        </AnimatePresence>

        <footer className="tm-footer">
          TM Sempre Tecnologia · Extrator DOCX v1.4 · Ocean Breeze Design
        </footer>
      </main>
    </div>
  );
}
