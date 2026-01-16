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
} from "lucide-react";
import * as XLSX from "xlsx";
import JSZip from "jszip";

/**
 * Extrator de Itens DOCX (Online)
 * Layout com CSS proprio + hierarquia visual.
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

  return Array.from(map.values()).sort((a, b) => {
    const ak = `${a.codigo} ${a.descricao}`.trim().toLowerCase();
    const bk = `${b.codigo} ${b.descricao}`.trim().toLowerCase();
    return ak.localeCompare(bk, "pt-BR");
  });
}

function Badge({ kind, icon, children }) {
  const cls =
    kind === "idle"
      ? "badge badge--idle"
      : kind === "work"
      ? "badge badge--work"
      : kind === "ok"
      ? "badge badge--ok"
      : "badge badge--err";

  return (
    <span className={cls}>
      <span className="badge__icon">{icon}</span>
      <span>{children}</span>
    </span>
  );
}

function StatCard({ label, value, sub }) {
  return (
    <div className="stat-card">
      <div className="stat-card__label">{label}</div>
      <div className="stat-card__value">{value}</div>
      {sub ? <div className="stat-card__sub">{sub}</div> : null}
    </div>
  );
}

function Section({ title, desc, right, children }) {
  return (
    <section className="panel">
      <div className="panel__header">
        <div>
          <h2 className="panel__title">{title}</h2>
          {desc ? <p className="panel__desc">{desc}</p> : null}
        </div>
        {right ? <div className="panel__right">{right}</div> : null}
      </div>
      {children}
    </section>
  );
}

export default function AppExtratorDocx() {
  const inputRef = useRef(null);
  const [drag, setDrag] = useState(false);

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
      setAggText("Soma concluida!");
      setAggLines([
        `Regra: ${keyLabel}`,
        `Itens unicos: ${fmtInt(ag.length)}`,
        "Excel consolidado pronto",
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
    if (phase === "work") return { kind: "work", icon: <Loader2 size={16} className="spin" /> };
    if (phase === "ok") return { kind: "ok", icon: <CheckCircle2 size={16} /> };
    if (phase === "err") return { kind: "err", icon: <AlertTriangle size={16} /> };
    return { kind: "idle", icon: <Info size={16} /> };
  }, [phase]);

  const aggBadge = useMemo(() => {
    if (aggPhase === "work") return { kind: "work", icon: <Loader2 size={16} className="spin" /> };
    if (aggPhase === "ok") return { kind: "ok", icon: <Sigma size={16} /> };
    if (aggPhase === "err") return { kind: "err", icon: <AlertTriangle size={16} /> };
    return { kind: "idle", icon: <Info size={16} /> };
  }, [aggPhase]);

  const keyLabel =
    aggRule === "code_only" ? "Apenas Codigo" : aggRule === "desc_only" ? "Apenas Descricao" : "Codigo + Descricao";

  return (
    <div className="app">
      <header className="app__header">
        <div className="brand">
          <img className="brand__logo" src="/tm_logo.svg" alt="TM Sempre Tecnologia" />
          <div className="brand__name">TM Sempre Tecnologia</div>
          <div className="brand__sub">Extrator de Itens DOCX - Layout vertical</div>
        </div>
        <div className="header__meta">
          <span className="pill pill--online">
            <span className="dot" />
            Online
          </span>
          <span className="pill">v1.3</span>
        </div>
      </header>

      <main className="app__main">
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
            className={cn("dropzone", drag && "dropzone--active")}
            role="button"
            tabIndex={0}
            onClick={onPick}
          >
            <div className="dropzone__row">
              <div className="dropzone__row">
                <div className="dropzone__icon">DOCX</div>
                <div className="dropzone__copy">
                  <div className="dropzone__title">Arraste o arquivo aqui</div>
                  <div className="dropzone__hint">ou clique para selecionar</div>
                </div>
              </div>
              <div className="dropzone__hint">Max. recomendado: 20MB</div>
            </div>
          </div>

          <input ref={inputRef} type="file" accept=".docx" hidden onChange={onInputChange} />

          <div className="actions" style={{ marginTop: "14px" }}>
            <button type="button" onClick={processDoc} disabled={!canProcess} className="btn btn--primary">
              {phase === "work" ? <Loader2 size={16} className="spin" /> : <FileText size={16} />}
              Processar documento
            </button>
            <button type="button" onClick={onPick} disabled={phase === "work"} className="btn btn--outline">
              <CloudUpload size={16} />
              Selecionar outro
            </button>
          </div>

          <div className="status" style={{ marginTop: "14px" }}>
            <div className="status__top">
              <span>{statusText}</span>
              <span className="status__file">{file?.name ?? "(nenhum)"}</span>
            </div>
            <div className="status__lines">
              {lines.map((l, i) => (
                <span key={i}>{l}</span>
              ))}
            </div>

            <div className="actions" style={{ marginTop: "14px" }}>
              <button
                type="button"
                onClick={downloadBruto}
                disabled={phase !== "ok" || items.length === 0}
                className="btn btn--primary"
              >
                <Download size={16} />
                Baixar Excel bruto
              </button>
              <button type="button" onClick={downloadLog} disabled={!logText} className="btn btn--outline">
                <Download size={16} />
                Baixar Log
              </button>
            </div>
          </div>
        </Section>

        <Section title="2) Resumo" desc="Metricas do processamento e do consolidado (quando gerado).">
          <div className="stats">
            <StatCard label="Itens extraidos" value={fmtInt(items.length)} />
            <StatCard label="Itens somados" value={fmtInt(aggItems.length)} sub={aggPhase === "ok" ? `Regra: ${keyLabel}` : "-"} />
            <div className="stats__wide">
              <StatCard label="Tabelas / Itens" value={meta ? `${meta.tables_total} / ${meta.itens_tables}` : "-"} />
              <StatCard label="Ignoradas" value={meta ? fmtInt(meta.rows_ignored) : "-"} />
            </div>
          </div>
        </Section>

        <AnimatePresence>
          {phase === "ok" && items.length > 0 ? (
            <motion.div initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: 8 }} transition={{ duration: 0.2 }}>
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
                <div className="panel__desc">Regra de chave:</div>
                <div className="rule-grid">
                  {[
                    { v: "code_desc", label: "Codigo + Descricao" },
                    { v: "code_only", label: "Apenas Codigo" },
                    { v: "desc_only", label: "Apenas Descricao" },
                  ].map((opt) => (
                    <label key={opt.v} className="rule-card">
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

                <div className="actions" style={{ marginTop: "14px" }}>
                  <button type="button" onClick={doAggregate} disabled={!canAggregate} className="btn btn--primary">
                    {aggPhase === "work" ? <Loader2 size={16} className="spin" /> : <Sigma size={16} />}
                    Gerar planilha somada
                  </button>
                  <button
                    type="button"
                    onClick={downloadSomado}
                    disabled={aggPhase !== "ok" || aggItems.length === 0}
                    className="btn btn--outline"
                  >
                    <Download size={16} />
                    Baixar Excel consolidado
                  </button>
                </div>

                <div className="status" style={{ marginTop: "14px" }}>
                  <div className="status__top">
                    <span>{aggText}</span>
                  </div>
                  <div className="status__lines">
                    {aggLines.map((l, i) => (
                      <span key={i}>{l}</span>
                    ))}
                  </div>

                  {aggPhase === "ok" && aggItems.length ? (
                    <div className="panel" style={{ marginTop: "12px", padding: "14px" }}>
                      <div className="panel__title" style={{ fontSize: "12px" }}>
                        Previa do consolidado (8 primeiros)
                      </div>
                      <div className="preview" style={{ marginTop: "10px" }}>
                        {aggItems.slice(0, 8).map((x, idx) => (
                          <div key={idx} className="preview__item">
                            <div className="preview__meta">
                              <div className="preview__code">{x.codigo || "(sem codigo)"}</div>
                              <div className="preview__desc">{x.descricao || "(sem descricao)"}</div>
                            </div>
                            <div className="preview__qty">{fmtQty(x.quantidade)}</div>
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

        <AnimatePresence>
          {phase === "ok" && items.length ? (
            <motion.div initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: 8 }} transition={{ duration: 0.2 }}>
              <Section
                title="4) Previa dos itens extraidos"
                desc="Mostra os 10 primeiros itens extraidos para conferencia rapida."
              >
                <div className="status__top" style={{ marginBottom: "10px" }}>
                  <span className="panel__title" style={{ fontSize: "12px" }}>
                    Primeiros 10
                  </span>
                  <span className="status__file">{fmtInt(items.length)} linhas</span>
                </div>

                <div className="preview">
                  {items.slice(0, 10).map((it, idx) => (
                    <div key={idx} className="preview__item">
                      <div className="preview__meta">
                        <div className="preview__code">{it.codigo}</div>
                        <div className="preview__desc">{it.descricao || "(sem descricao)"}</div>
                        <div className="preview__origin">origem: {it.origem}</div>
                      </div>
                      <div className="preview__qty">{it.quantidade_raw}</div>
                    </div>
                  ))}
                </div>
              </Section>
            </motion.div>
          ) : null}
        </AnimatePresence>

        <Section title="5) Regras e privacidade" desc="Referencia rapida das regras de extracao e garantia de processamento local.">
          <div className="info-grid">
            <div className="info-card">
              <div className="info-card__title">Regras de extracao</div>
              <ul>
                <li>Busca tabelas com cabecalho "Itens" na 1a linha.</li>
                <li>Coluna 1: Codigo (aceita 17.4 / 13.12 etc). Ignora #N/D.</li>
                <li>Coluna 2: Descricao.</li>
                <li>Quantidade: prefere 3a coluna; fallback por numero na linha.</li>
                <li>Exporta Excel (.xlsx) e Log (.txt).</li>
              </ul>
            </div>

            <div className="info-card">
              <div className="info-card__title">Privacidade</div>
              <p className="panel__desc">
                O processamento acontece no seu navegador. Nenhum arquivo e enviado para servidor.
              </p>
            </div>
          </div>
        </Section>
      </main>

      <footer className="footer">TM Sempre Tecnologia - Extrator DOCX - v1.3</footer>
    </div>
  );
}
