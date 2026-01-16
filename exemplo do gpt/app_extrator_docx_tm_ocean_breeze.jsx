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
 * Extrator de Itens DOCX (Online)
 * Tema aplicado: Design System TM — Ocean Breeze (tokens CSS + light/dark)
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
    } catch {
      // fallback
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
    (it, i) => `${String(i + 1).padStart(2, "0")}. ${it.codigo} | ${it.descricao} | qtd=${it.quantidade_raw}`
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

  return Array.from(map.values()).sort((a, b) => {
    const ak = `${a.codigo} ${a.descricao}`.trim().toLowerCase();
    const bk = `${b.codigo} ${b.descricao}`.trim().toLowerCase();
    return ak.localeCompare(bk, "pt-BR");
  });
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

  const [items, setItems] = useState(/** @type {Item[]} */ ([]));
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
    setLines(["Arquivo selecionado", "Clique em PROCESSAR DOCUMENTO"]);

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
        "Gere Excel bruto e (opcional) consolidado",
        "Log disponivel para auditoria",
      ]);
    } catch (err) {
      setPhase("err");
      setStatusText("Erro ao processar o DOCX.");
      setLines([String(err?.message ?? err), "Dica: tente exportar o Word novamente (DOCX padrao)."]);

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
      const keyLabel2 =
        aggRule === "code_only"
          ? "Apenas Codigo"
          : aggRule === "desc_only"
          ? "Apenas Descricao"
          : "Codigo + Descricao";
      setAggText("Soma concluida!");
      setAggLines([`Regra: ${keyLabel2}`, `Itens unicos: ${fmtInt(ag.length)}`, "Excel consolidado pronto"]);

      if (file) {
        const extra = `\n\n--- Consolidado (amostra) ---\n${ag
          .slice(0, 10)
          .map(
            (x, i) =>
              `${String(i + 1).padStart(2, "0")}. ${x.codigo} | ${x.descricao} | qtd=${fmtQty(x.quantidade)}`
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

  const keyLabel =
    aggRule === "code_only" ? "Apenas Codigo" : aggRule === "desc_only" ? "Apenas Descricao" : "Codigo + Descricao";

  const toggleTheme = useCallback(() => {
    setTheme((t) => (t === "dark" ? "light" : "dark"));
  }, []);

  return (
    <div className={cn("tm-root", theme === "dark" && "dark")}>
      <style>{styles}</style>

      <header className="tm-header">
        <div className="tm-brand">
          <div className="tm-brand__mark">TM</div>
          <div>
            <div className="tm-brand__name">TM Sempre Tecnologia</div>
            <div className="tm-brand__sub">Extrator de Itens DOCX · Ocean Breeze</div>
          </div>
        </div>

        <div className="tm-header__right">
          <span className="tm-pill tm-pill--online">
            <span className="tm-dot" />
            Online
          </span>
          <span className="tm-pill">v1.3</span>
          <button type="button" className="tm-btn tm-btn--outline" onClick={toggleTheme} aria-label="Alternar tema">
            {theme === "dark" ? <Sun size={16} /> : <Moon size={16} />}
            {theme === "dark" ? "Claro" : "Escuro"}
          </button>
        </div>
      </header>

      <main className="tm-main">
        <Section
          title="1) Enviar e processar"
          desc={
            <>
              Envie um arquivo <b>.docx</b>. Depois gere o <b>Excel bruto</b> e (opcional) o <b>consolidado</b>.
            </>
          }
          right={
            <Badge kind={badge.kind} icon={badge.icon}>
              {phase === "idle" ? "Aguardando" : phase === "work" ? "Processando" : phase === "ok" ? "Sucesso" : "Erro"}
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
                  <div className="tm-drop__title">Arraste o arquivo aqui</div>
                  <div className="tm-drop__hint">ou clique para selecionar</div>
                </div>
              </div>
              <div className="tm-drop__hint">Max. recomendado: 20MB</div>
            </div>
          </div>

          <input ref={inputRef} type="file" accept=".docx" hidden onChange={onInputChange} />

          <div className="tm-actions" style={{ marginTop: 14 }}>
            <button type="button" onClick={processDoc} disabled={!canProcess} className="tm-btn tm-btn--primary">
              {phase === "work" ? <Loader2 size={16} className="tm-spin" /> : <FileText size={16} />}
              Processar documento
            </button>
            <button type="button" onClick={onPick} disabled={phase === "work"} className="tm-btn tm-btn--outline">
              <CloudUpload size={16} />
              Selecionar outro
            </button>
          </div>

          <div className="tm-status" style={{ marginTop: 14 }}>
            <div className="tm-status__top">
              <span>{statusText}</span>
              <span className="tm-status__file">{file?.name ?? "(nenhum)"}</span>
            </div>
            <div className="tm-status__lines">
              {lines.map((l, i) => (
                <span key={i}>{l}</span>
              ))}
            </div>
          </div>
        </Section>

        {phase === "ok" ? (
          <Section title="Resultado" desc="Pronto para download.">
            <div className="tm-status__top" style={{ marginBottom: 10 }}>
              <span>Extracao concluida com sucesso.</span>
              <span className="tm-status__file">{fmtInt(items.length)} itens</span>
            </div>
            <div className="tm-actions">
              <button type="button" onClick={downloadBruto} disabled={!items.length} className="tm-btn tm-btn--primary">
                <Download size={16} />
                Baixar Excel bruto
              </button>
              <button type="button" onClick={downloadLog} disabled={!logText} className="tm-btn tm-btn--outline">
                <Download size={16} />
                Baixar Log
              </button>
            </div>
          </Section>
        ) : null}

        <AnimatePresence>
          {phase === "ok" && items.length > 0 ? (
            <motion.div
              initial={{ opacity: 0, y: 8 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 8 }}
              transition={{ duration: 0.2 }}
            >
              <Section
                title="3) Somar itens iguais"
                desc={
                  <>
                    Gera uma planilha consolidada somando <b>Quantidade</b> para itens repetidos.
                  </>
                }
                right={
                  <Badge kind={aggBadge.kind} icon={aggBadge.icon}>
                    {aggPhase === "idle" ? "Pronto" : aggPhase === "work" ? "Somando" : aggPhase === "ok" ? "Concluido" : "Erro"}
                  </Badge>
                }
              >
                <div className="tm-subtle">Regra de chave:</div>
                <div className="tm-rule-grid">
                  {[
                    { v: "code_desc", label: "Codigo + Descricao" },
                    { v: "code_only", label: "Apenas Codigo" },
                    { v: "desc_only", label: "Apenas Descricao" },
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

                <div className="tm-actions" style={{ marginTop: 14 }}>
                  <button type="button" onClick={doAggregate} disabled={!canAggregate} className="tm-btn tm-btn--primary">
                    {aggPhase === "work" ? <Loader2 size={16} className="tm-spin" /> : <Sigma size={16} />}
                    Gerar planilha somada
                  </button>
                  <button
                    type="button"
                    onClick={downloadSomado}
                    disabled={aggPhase !== "ok" || aggItems.length === 0}
                    className="tm-btn tm-btn--outline"
                  >
                    <Download size={16} />
                    Baixar Excel consolidado
                  </button>
                </div>

                <div className="tm-status" style={{ marginTop: 14 }}>
                  <div className="tm-status__top">
                    <span>{aggText}</span>
                  </div>
                  <div className="tm-status__lines">
                    {aggLines.map((l, i) => (
                      <span key={i}>{l}</span>
                    ))}
                  </div>

                  {aggPhase === "ok" && aggItems.length ? (
                    <div className="tm-panel tm-panel--inner" style={{ marginTop: 12 }}>
                      <div className="tm-panel__title" style={{ fontSize: 12 }}>
                        Previa do consolidado (8 primeiros)
                      </div>
                      <div className="tm-preview" style={{ marginTop: 10 }}>
                        {aggItems.slice(0, 8).map((x, idx) => (
                          <div key={idx} className="tm-preview__item">
                            <div className="tm-preview__meta">
                              <div className="tm-preview__code">{x.codigo || "(sem codigo)"}</div>
                              <div className="tm-preview__desc">{x.descricao || "(sem descricao)"}</div>
                            </div>
                            <div className="tm-preview__qty">{fmtQty(x.quantidade)}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : null}
                </div>
              </Section>
            </motion.div>
          ) : null}
        </AnimatePresence>

        <details className="tm-details">
          <summary className="tm-details__summary">Detalhes tecnicos</summary>

          <Section title="2) Resumo" desc="Metricas do processamento e do consolidado (quando gerado).">
            <div className="tm-stats">
              <StatCard label="Itens extraidos" value={fmtInt(items.length)} />
              <StatCard label="Itens somados" value={fmtInt(aggItems.length)} sub={aggPhase === "ok" ? `Regra: ${keyLabel}` : "-"} />
              <div className="tm-stats__wide">
                <StatCard label="Tabelas / Itens" value={meta ? `${meta.tables_total} / ${meta.itens_tables}` : "-"} />
                <StatCard label="Ignoradas" value={meta ? fmtInt(meta.rows_ignored) : "-"} />
              </div>
            </div>
          </Section>

          <AnimatePresence>
            {phase === "ok" && items.length ? (
              <motion.div
                initial={{ opacity: 0, y: 8 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: 8 }}
                transition={{ duration: 0.2 }}
              >
                <Section title="4) Previa dos itens extraidos" desc="Mostra os 10 primeiros itens extraidos para conferencia rapida.">
                  <div className="tm-status__top" style={{ marginBottom: 10 }}>
                    <span className="tm-mini-title">Primeiros 10</span>
                    <span className="tm-status__file">{fmtInt(items.length)} linhas</span>
                  </div>

                  <div className="tm-preview">
                    {items.slice(0, 10).map((it, idx) => (
                      <div key={idx} className="tm-preview__item">
                        <div className="tm-preview__meta">
                          <div className="tm-preview__code">{it.codigo}</div>
                          <div className="tm-preview__desc">{it.descricao || "(sem descricao)"}</div>
                          <div className="tm-preview__origin">origem: {it.origem}</div>
                        </div>
                        <div className="tm-preview__qty">{it.quantidade_raw}</div>
                      </div>
                    ))}
                  </div>
                </Section>
              </motion.div>
            ) : null}
          </AnimatePresence>

          <Section title="5) Regras e privacidade" desc="Referencia rapida das regras de extracao e garantia de processamento local.">
            <div className="tm-info-grid">
              <div className="tm-info">
                <div className="tm-info__title">Regras de extracao</div>
                <ul className="tm-ul">
                  <li>Busca tabelas com cabecalho "Itens" na 1a linha.</li>
                  <li>Coluna 1: Codigo (aceita 17.4 / 13.12 etc). Ignora #N/D.</li>
                  <li>Coluna 2: Descricao.</li>
                  <li>Quantidade: prefere 3a coluna; fallback por numero na linha.</li>
                  <li>Exporta Excel (.xlsx) e Log (.txt).</li>
                </ul>
              </div>

              <div className="tm-info">
                <div className="tm-info__title">Privacidade</div>
                <p className="tm-panel__desc">O processamento acontece no seu navegador. Nenhum arquivo e enviado para servidor.</p>
              </div>
            </div>
          </Section>
        </details>

        <footer className="tm-footer">TM Sempre Tecnologia · Extrator DOCX · v1.3</footer>
      </main>
    </div>
  );
}

const styles = `
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Lora:wght@400;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

.tm-root{
  --TM-background:#f0f8ff;
  --TM-foreground:#374151;
  --TM-card:#ffffff;
  --TM-card-foreground:#374151;
  --TM-popover:#ffffff;
  --TM-popover-foreground:#374151;
  --TM-primary:#22c55e;
  --TM-primary-foreground:#ffffff;
  --TM-secondary:#e0f2fe;
  --TM-secondary-foreground:#4b5563;
  --TM-muted:#f3f4f6;
  --TM-muted-foreground:#6b7280;
  --TM-accent:#d1fae5;
  --TM-accent-foreground:#374151;
  --TM-destructive:#ef4444;
  --TM-destructive-foreground:#ffffff;
  --TM-border:#e5e7eb;
  --TM-input:#e5e7eb;
  --TM-ring:#22c55e;

  --TM-radius-sm:calc(0.5rem - 4px);
  --TM-radius-md:calc(0.5rem - 2px);
  --TM-radius-lg:0.5rem;
  --TM-radius-xl:calc(0.5rem + 4px);

  --TM-shadow-2xs:0px 4px 8px -1px rgba(0,0,0,0.05);
  --TM-shadow-xs:0px 4px 8px -1px rgba(0,0,0,0.05);
  --TM-shadow-sm:0px 4px 8px -1px rgba(0,0,0,0.10), 0px 1px 2px -2px rgba(0,0,0,0.10);
  --TM-shadow:0px 4px 8px -1px rgba(0,0,0,0.10), 0px 1px 2px -2px rgba(0,0,0,0.10);
  --TM-shadow-md:0px 4px 8px -1px rgba(0,0,0,0.10), 0px 2px 4px -2px rgba(0,0,0,0.10);
  --TM-shadow-lg:0px 4px 8px -1px rgba(0,0,0,0.10), 0px 4px 6px -2px rgba(0,0,0,0.10);
  --TM-shadow-xl:0px 4px 8px -1px rgba(0,0,0,0.10), 0px 8px 10px -2px rgba(0,0,0,0.10);
  --TM-shadow-2xl:0px 4px 8px -1px rgba(0,0,0,0.25);

  font-family:'DM Sans', system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
  color:var(--TM-foreground);
  background:radial-gradient(1200px 600px at 20% -10%, rgba(34,197,94,0.12), transparent 55%),
             radial-gradient(1000px 500px at 90% 0%, rgba(14,165,233,0.10), transparent 50%),
             var(--TM-background);
  min-height:100vh;
}

.tm-root.dark{
  --TM-background:#0f172a;
  --TM-foreground:#d1d5db;
  --TM-card:#1e293b;
  --TM-card-foreground:#d1d5db;
  --TM-popover:#1e293b;
  --TM-popover-foreground:#d1d5db;
  --TM-primary:#34d399;
  --TM-primary-foreground:#0f172a;
  --TM-secondary:#2d3748;
  --TM-secondary-foreground:#a1a1aa;
  --TM-muted:#19212e;
  --TM-muted-foreground:#6b7280;
  --TM-accent:#374151;
  --TM-accent-foreground:#a1a1aa;
  --TM-destructive:#ef4444;
  --TM-destructive-foreground:#0f172a;
  --TM-border:#4b5563;
  --TM-input:#4b5563;
  --TM-ring:#34d399;
}

.tm-header{
  position:sticky; top:0; z-index:50;
  display:flex; align-items:center; justify-content:space-between;
  padding:14px 18px;
  background:rgba(255,255,255,0.75);
  backdrop-filter: blur(10px);
  border-bottom:1px solid var(--TM-border);
  box-shadow:var(--TM-shadow-sm);
}
.tm-root.dark .tm-header{ background:rgba(30,41,59,0.72); }

.tm-brand{ display:flex; gap:12px; align-items:center; }
.tm-brand__mark{
  width:44px; height:44px; border-radius:14px;
  background:linear-gradient(135deg, var(--TM-primary), #0ea5e9);
  color:var(--TM-primary-foreground);
  font-weight:800; display:grid; place-items:center;
  box-shadow:var(--TM-shadow-md);
  letter-spacing:0.5px;
}
.tm-brand__name{ font-weight:800; font-size:14px; line-height:1.1; }
.tm-brand__sub{ color:var(--TM-muted-foreground); font-size:12px; margin-top:2px; }

.tm-header__right{ display:flex; gap:10px; align-items:center; }

.tm-pill{
  display:inline-flex; align-items:center; gap:8px;
  padding:6px 10px; border-radius:999px;
  background:var(--TM-secondary);
  color:var(--TM-secondary-foreground);
  border:1px solid var(--TM-border);
  font-size:12px; font-weight:600;
}
.tm-pill--online{ background:var(--TM-accent); color:var(--TM-accent-foreground); }
.tm-dot{ width:8px; height:8px; border-radius:999px; background:var(--TM-primary); box-shadow:0 0 0 4px rgba(34,197,94,0.18); }
.tm-root.dark .tm-dot{ box-shadow:0 0 0 4px rgba(52,211,153,0.18); }

.tm-main{ max-width:1100px; margin:0 auto; padding:18px; }

.tm-panel{
  background:var(--TM-card);
  color:var(--TM-card-foreground);
  border:1px solid var(--TM-border);
  border-radius:18px;
  box-shadow:var(--TM-shadow);
  overflow:hidden;
  margin:16px 0;
}
.tm-panel--inner{ padding:14px; }

.tm-panel__header{
  display:flex; justify-content:space-between; gap:14px;
  padding:14px 14px 12px 14px;
  border-bottom:1px solid var(--TM-border);
}
.tm-panel__left{ min-width:0; }
.tm-panel__title{ font-size:14px; font-weight:800; margin:0; }
.tm-panel__desc{ font-size:12px; color:var(--TM-muted-foreground); margin-top:6px; }
.tm-panel__body{ padding:14px; }

.tm-btn{
  display:inline-flex; align-items:center; gap:10px;
  border-radius:12px;
  padding:10px 14px;
  font-weight:700;
  font-size:13px;
  border:1px solid var(--TM-border);
  background:transparent;
  color:var(--TM-foreground);
  cursor:pointer;
  transition: transform .15s ease, box-shadow .15s ease, opacity .15s ease, background .15s ease;
}
.tm-btn:disabled{ opacity:0.55; cursor:not-allowed; }
.tm-btn--primary{
  background:var(--TM-primary);
  color:var(--TM-primary-foreground);
  border-color:transparent;
  box-shadow:var(--TM-shadow-sm);
}
.tm-btn--primary:hover{ transform: translateY(-1px); box-shadow:var(--TM-shadow-md); }
.tm-btn--outline:hover{ background:var(--TM-accent); }

.tm-actions{ display:flex; flex-wrap:wrap; gap:10px; }

.tm-drop{
  border:1.5px dashed var(--TM-border);
  border-radius:18px;
  padding:16px;
  background:linear-gradient(180deg, rgba(224,242,254,0.55), rgba(209,250,229,0.45));
  cursor:pointer;
  transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
}
.tm-root.dark .tm-drop{ background:linear-gradient(180deg, rgba(45,55,72,0.55), rgba(25,33,46,0.5)); }
.tm-drop--active{ border-color:var(--TM-ring); box-shadow: 0 0 0 4px rgba(34,197,94,0.12); }
.tm-root.dark .tm-drop--active{ box-shadow: 0 0 0 4px rgba(52,211,153,0.12); }

.tm-drop__row{ display:flex; align-items:center; justify-content:space-between; gap:16px; }
.tm-drop__left{ display:flex; gap:12px; align-items:center; }
.tm-drop__icon{
  width:52px; height:42px; border-radius:14px;
  background:rgba(255,255,255,0.75);
  border:1px solid var(--TM-border);
  display:grid; place-items:center;
  font-weight:900; font-size:12px;
  box-shadow:var(--TM-shadow-xs);
}
.tm-root.dark .tm-drop__icon{ background:rgba(30,41,59,0.65); }
.tm-drop__title{ font-weight:900; font-size:14px; }
.tm-drop__hint{ color:var(--TM-muted-foreground); font-size:12px; }

.tm-status{
  margin-top:12px;
  border:1px solid var(--TM-border);
  border-radius:14px;
  padding:12px;
  background:var(--TM-muted);
}
.tm-status__top{ display:flex; justify-content:space-between; gap:12px; font-size:12px; font-weight:700; }
.tm-status__file{ color:var(--TM-muted-foreground); font-weight:700; font-family:'IBM Plex Mono', ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; }
.tm-status__lines{ margin-top:8px; display:flex; flex-direction:column; gap:4px; font-size:12px; color:var(--TM-muted-foreground); }

.tm-badge{
  display:inline-flex; align-items:center; gap:8px;
  padding:6px 10px;
  border-radius:999px;
  font-size:12px;
  font-weight:800;
  border:1px solid var(--TM-border);
  background:var(--TM-muted);
  color:var(--TM-foreground);
}
.tm-badge--work{ background:rgba(14,165,233,0.12); }
.tm-badge--ok{ background:rgba(34,197,94,0.14); }
.tm-badge--err{ background:rgba(239,68,68,0.12); }
.tm-badge__icon{ display:grid; place-items:center; }

.tm-spin{ animation: tmspin 1s linear infinite; }
@keyframes tmspin{ from{ transform: rotate(0deg);} to{ transform: rotate(360deg);} }

.tm-subtle{ color:var(--TM-muted-foreground); font-size:12px; font-weight:700; margin-top:2px; }

.tm-rule-grid{ display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:10px; margin-top:10px; }
@media (max-width: 900px){ .tm-rule-grid{ grid-template-columns: 1fr; } }

.tm-rule{
  display:flex; gap:10px; align-items:center;
  padding:12px 12px;
  border-radius:14px;
  border:1px solid var(--TM-border);
  background:var(--TM-card);
  box-shadow:var(--TM-shadow-xs);
  cursor:pointer;
  transition: transform .12s ease, box-shadow .12s ease;
  font-weight:800;
  font-size:12px;
}
.tm-rule input{ accent-color: var(--TM-primary); }
.tm-rule:hover{ transform: translateY(-1px); box-shadow: var(--TM-shadow-md); }
.tm-rule--active{ outline: 3px solid rgba(34,197,94,0.18); }
.tm-root.dark .tm-rule--active{ outline: 3px solid rgba(52,211,153,0.18); }

.tm-preview{ display:flex; flex-direction:column; gap:8px; }
.tm-preview__item{
  display:flex; justify-content:space-between; gap:12px; align-items:flex-start;
  padding:12px;
  border:1px solid var(--TM-border);
  border-radius:14px;
  background:var(--TM-card);
}
.tm-preview__meta{ min-width:0; }
.tm-preview__code{ font-family:'IBM Plex Mono', ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-weight:900; font-size:12px; }
.tm-preview__desc{ font-size:12px; color:var(--TM-muted-foreground); margin-top:2px; word-break:break-word; }
.tm-preview__origin{ font-size:11px; color:var(--TM-muted-foreground); margin-top:4px; font-family:'IBM Plex Mono', ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; }
.tm-preview__qty{ font-weight:900; font-size:13px; color:var(--TM-foreground); white-space:nowrap; }

.tm-stats{ display:grid; grid-template-columns: 1fr 1fr; gap:12px; }
.tm-stats__wide{ grid-column: 1 / -1; display:grid; grid-template-columns: 1fr 1fr; gap:12px; }
@media (max-width: 700px){ .tm-stats{ grid-template-columns: 1fr; } .tm-stats__wide{ grid-template-columns: 1fr; } }

.tm-stat{
  border:1px solid var(--TM-border);
  border-radius:16px;
  padding:14px;
  background:linear-gradient(180deg, rgba(224,242,254,0.50), rgba(209,250,229,0.35));
}
.tm-root.dark .tm-stat{ background:linear-gradient(180deg, rgba(45,55,72,0.50), rgba(25,33,46,0.45)); }
.tm-stat__label{ font-size:12px; color:var(--TM-muted-foreground); font-weight:800; }
.tm-stat__value{ font-size:18px; font-weight:900; margin-top:4px; }
.tm-stat__sub{ font-size:12px; color:var(--TM-muted-foreground); margin-top:4px; }

.tm-info-grid{ display:grid; grid-template-columns: 1fr 1fr; gap:12px; }
@media (max-width: 900px){ .tm-info-grid{ grid-template-columns: 1fr; } }

.tm-info{
  border:1px solid var(--TM-border);
  border-radius:16px;
  padding:14px;
  background:var(--TM-card);
}
.tm-info__title{ font-weight:900; font-size:13px; margin-bottom:8px; }
.tm-ul{ padding-left:18px; color:var(--TM-muted-foreground); font-size:12px; }
.tm-ul li{ margin:6px 0; }

.tm-details{ margin:16px 0; }
.tm-details__summary{
  list-style:none;
  cursor:pointer;
  font-weight:900;
  font-size:12px;
  color:var(--TM-foreground);
  padding:12px 14px;
  border-radius:14px;
  border:1px solid var(--TM-border);
  background:rgba(255,255,255,0.55);
  box-shadow:var(--TM-shadow-xs);
}
.tm-root.dark .tm-details__summary{ background:rgba(30,41,59,0.55); }
.tm-details__summary::-webkit-details-marker{ display:none; }

.tm-mini-title{ font-weight:900; font-size:12px; }

.tm-footer{ color:var(--TM-muted-foreground); font-size:12px; padding:10px 6px 22px; text-align:center; }
`;
