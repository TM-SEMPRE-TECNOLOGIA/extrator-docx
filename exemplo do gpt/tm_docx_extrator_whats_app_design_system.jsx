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
  Search,
  MoreVertical,
  MessageCircle,
  ShieldCheck,
} from "lucide-react";
import * as XLSX from "xlsx";
import JSZip from "jszip";

/**
 * TM — Extrator de Itens DOCX (Online)
 * Tema: WhatsApp Design System (tokens + componentes)
 * - Visual inspirado no WhatsApp (verde, superfícies claras, chips, listas, cards)
 * - Sem alterar regras/funcionalidade do extrator (somente UI)
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

  const ignored = (meta?.ignored_details ?? []).slice(0, 250);
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

function Chip({ active, icon, children, onClick }) {
  return (
    <button type="button" className={cn("wa-chip", active && "wa-chip--active")} onClick={onClick}>
      {icon ? <span className="wa-chip__icon">{icon}</span> : null}
      <span>{children}</span>
    </button>
  );
}

function Badge({ kind, icon, children }) {
  const cls =
    kind === "idle"
      ? "wa-badge"
      : kind === "work"
      ? "wa-badge wa-badge--work"
      : kind === "ok"
      ? "wa-badge wa-badge--ok"
      : "wa-badge wa-badge--err";

  return (
    <span className={cls}>
      <span className="wa-badge__icon">{icon}</span>
      <span>{children}</span>
    </span>
  );
}

function Section({ title, desc, right, children }) {
  return (
    <section className="wa-card">
      <div className="wa-card__head">
        <div className="wa-card__head-left">
          <div className="wa-card__title">{title}</div>
          {desc ? <div className="wa-card__desc">{desc}</div> : null}
        </div>
        {right ? <div className="wa-card__head-right">{right}</div> : null}
      </div>
      <div className="wa-card__body">{children}</div>
    </section>
  );
}

export default function AppWhatsDocx() {
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

  const [query, setQuery] = useState("");

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
        aggRule === "code_only" ? "Apenas Codigo" : aggRule === "desc_only" ? "Apenas Descricao" : "Codigo + Descricao";
      setAggText("Soma concluida!");
      setAggLines([`Regra: ${keyLabel2}`, `Itens unicos: ${fmtInt(ag.length)}`, "Excel consolidado pronto"]);

      if (file) {
        const extra = `\n\n--- Consolidado (amostra) ---\n${ag
          .slice(0, 10)
          .map((x, i) => `${String(i + 1).padStart(2, "0")}. ${x.codigo} | ${x.descricao} | qtd=${fmtQty(x.quantidade)}`)
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
    if (phase === "work") return { kind: "work", icon: <Loader2 size={16} className="wa-spin" /> };
    if (phase === "ok") return { kind: "ok", icon: <CheckCircle2 size={16} /> };
    if (phase === "err") return { kind: "err", icon: <AlertTriangle size={16} /> };
    return { kind: "idle", icon: <Info size={16} /> };
  }, [phase]);

  const aggBadge = useMemo(() => {
    if (aggPhase === "work") return { kind: "work", icon: <Loader2 size={16} className="wa-spin" /> };
    if (aggPhase === "ok") return { kind: "ok", icon: <Sigma size={16} /> };
    if (aggPhase === "err") return { kind: "err", icon: <AlertTriangle size={16} /> };
    return { kind: "idle", icon: <Info size={16} /> };
  }, [aggPhase]);

  const filteredItems = useMemo(() => {
    const q = norm(query).toLowerCase();
    if (!q) return items;
    return items.filter((it) => {
      const a = `${it.codigo} ${it.descricao} ${it.quantidade_raw}`.toLowerCase();
      return a.includes(q);
    });
  }, [items, query]);

  return (
    <div className="wa-root">
      <style>{styles}</style>

      {/* Top Bar (WhatsApp-like) */}
      <header className="wa-top">
        <div className="wa-top__left">
          <div className="wa-avatar">TM</div>
          <div className="wa-top__titles">
            <div className="wa-top__title">Extrator DOCX</div>
            <div className="wa-top__subtitle">TM Sempre Tecnologia</div>
          </div>
        </div>
        <div className="wa-top__right">
          <button className="wa-icon" type="button" aria-label="Pesquisa" title="Pesquisa">
            <Search size={18} />
          </button>
          <button className="wa-icon" type="button" aria-label="Menu" title="Menu">
            <MoreVertical size={18} />
          </button>
        </div>
      </header>

      {/* Tabs / Chips */}
      <div className="wa-tabs">
        <Chip active icon={<MessageCircle size={16} />} onClick={() => {}}>
          Extrair
        </Chip>
        <Chip icon={<Sigma size={16} />} onClick={() => {}}>
          Somar
        </Chip>
        <Chip icon={<ShieldCheck size={16} />} onClick={() => {}}>
          Privacidade
        </Chip>
      </div>

      <main className="wa-main">
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
            className={cn("wa-drop", drag && "wa-drop--active")}
            role="button"
            tabIndex={0}
            onClick={onPick}
          >
            <div className="wa-drop__row">
              <div className="wa-drop__left">
                <div className="wa-doc">DOCX</div>
                <div>
                  <div className="wa-drop__title">Arraste o arquivo aqui</div>
                  <div className="wa-drop__hint">ou clique para selecionar</div>
                </div>
              </div>
              <div className="wa-drop__hint">Recomendado: ate 20MB</div>
            </div>
          </div>

          <input ref={inputRef} type="file" accept=".docx" hidden onChange={onInputChange} />

          <div className="wa-actions" style={{ marginTop: 12 }}>
            <button type="button" onClick={processDoc} disabled={!canProcess} className="wa-btn wa-btn--primary">
              {phase === "work" ? <Loader2 size={16} className="wa-spin" /> : <FileText size={16} />}
              Processar documento
            </button>
            <button type="button" onClick={onPick} disabled={phase === "work"} className="wa-btn wa-btn--ghost">
              <CloudUpload size={16} />
              Selecionar outro
            </button>
          </div>

          <div className="wa-status" style={{ marginTop: 12 }}>
            <div className="wa-status__top">
              <span>{statusText}</span>
              <span className="wa-mono">{file?.name ?? "(nenhum)"}</span>
            </div>
            <div className="wa-status__lines">
              {lines.map((l, i) => (
                <span key={i}>{l}</span>
              ))}
            </div>
          </div>
        </Section>

        {phase === "ok" ? (
          <Section title="Resultado" desc="Pronto para download.">
            <div className="wa-row">
              <div className="wa-pill">
                <span className="wa-dot" />
                {fmtInt(items.length)} itens extraidos
              </div>
              <div className="wa-pill wa-pill--muted">Log: auditoria</div>
            </div>

            <div className="wa-actions" style={{ marginTop: 12 }}>
              <button type="button" onClick={downloadBruto} disabled={!items.length} className="wa-btn wa-btn--primary">
                <Download size={16} />
                Baixar Excel bruto
              </button>
              <button type="button" onClick={downloadLog} disabled={!logText} className="wa-btn wa-btn--ghost">
                <Download size={16} />
                Baixar Log
              </button>
            </div>

            <div className="wa-search" style={{ marginTop: 12 }}>
              <Search size={16} />
              <input
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="Filtrar itens extraidos (codigo, descricao...)"
              />
            </div>
          </Section>
        ) : null}

        <AnimatePresence>
          {phase === "ok" && items.length > 0 ? (
            <motion.div initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: 8 }} transition={{ duration: 0.2 }}>
              <Section
                title="2) Somar itens iguais"
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
                <div className="wa-sub">Regra de chave</div>
                <div className="wa-rule">
                  {[
                    { v: "code_desc", label: "Codigo + Descricao" },
                    { v: "code_only", label: "Apenas Codigo" },
                    { v: "desc_only", label: "Apenas Descricao" },
                  ].map((opt) => (
                    <label key={opt.v} className={cn("wa-radio", aggRule === opt.v && "wa-radio--on")}>
                      <input type="radio" name="rule" value={opt.v} checked={aggRule === opt.v} onChange={() => setAggRule(opt.v)} />
                      {opt.label}
                    </label>
                  ))}
                </div>

                <div className="wa-actions" style={{ marginTop: 12 }}>
                  <button type="button" onClick={doAggregate} disabled={!canAggregate} className="wa-btn wa-btn--primary">
                    {aggPhase === "work" ? <Loader2 size={16} className="wa-spin" /> : <Sigma size={16} />}
                    Gerar planilha somada
                  </button>
                  <button
                    type="button"
                    onClick={downloadSomado}
                    disabled={aggPhase !== "ok" || aggItems.length === 0}
                    className="wa-btn wa-btn--ghost"
                  >
                    <Download size={16} />
                    Baixar consolidado
                  </button>
                </div>

                <div className="wa-status" style={{ marginTop: 12 }}>
                  <div className="wa-status__top">
                    <span>{aggText}</span>
                  </div>
                  <div className="wa-status__lines">
                    {aggLines.map((l, i) => (
                      <span key={i}>{l}</span>
                    ))}
                  </div>
                </div>

                {aggPhase === "ok" && aggItems.length ? (
                  <div className="wa-list" style={{ marginTop: 12 }}>
                    <div className="wa-list__head">Previa do consolidado (8 primeiros)</div>
                    {aggItems.slice(0, 8).map((x, idx) => (
                      <div key={idx} className="wa-item">
                        <div className="wa-item__left">
                          <div className="wa-item__code">{x.codigo || "(sem codigo)"}</div>
                          <div className="wa-item__desc">{x.descricao || "(sem descricao)"}</div>
                        </div>
                        <div className="wa-item__qty">{fmtQty(x.quantidade)}</div>
                      </div>
                    ))}
                  </div>
                ) : null}
              </Section>
            </motion.div>
          ) : null}
        </AnimatePresence>

        {phase === "ok" && items.length ? (
          <Section title="3) Previa dos itens extraidos" desc="Mostra os primeiros itens para conferencia rapida.">
            <div className="wa-list">
              <div className="wa-list__head">Primeiros 10 · {fmtInt(filteredItems.length)} exibidos</div>
              {filteredItems.slice(0, 10).map((it, idx) => (
                <div key={idx} className="wa-item">
                  <div className="wa-item__left">
                    <div className="wa-item__code">{it.codigo}</div>
                    <div className="wa-item__desc">{it.descricao || "(sem descricao)"}</div>
                    <div className="wa-item__meta">origem: {it.origem} · qtd: {it.quantidade_raw}</div>
                  </div>
                  <div className="wa-item__qty">{it.quantidade_raw}</div>
                </div>
              ))}
            </div>
          </Section>
        ) : null}

        <Section
          title="4) Regras e privacidade"
          desc={
            <>
              O processamento acontece no seu navegador. Nenhum arquivo e enviado para servidor.
            </>
          }
        >
          <div className="wa-info">
            <div className="wa-info__title">Regras de extracao</div>
            <ul>
              <li>Busca tabelas com cabecalho "Itens" na 1a linha.</li>
              <li>Coluna 1: Codigo (aceita 17.4 / 13.12 etc). Ignora #N/D.</li>
              <li>Coluna 2: Descricao.</li>
              <li>Quantidade: prefere 3a coluna; fallback por numero na linha.</li>
              <li>Exporta Excel (.xlsx) e Log (.txt).</li>
            </ul>
          </div>
        </Section>
      </main>

      {/* Floating Action Button */}
      <button className="wa-fab" type="button" onClick={onPick} aria-label="Selecionar DOCX" title="Selecionar DOCX">
        <CloudUpload size={20} />
      </button>
    </div>
  );
}

const styles = `
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500&display=swap');

:root{
  --wa-green:#25D366;
  --wa-green-d:#1DAA57;
  --wa-teal:#128C7E;
  --wa-bg:#ECE5DD;
  --wa-surface:#FFFFFF;
  --wa-text:#111827;
  --wa-sub:#6B7280;
  --wa-border:rgba(17,24,39,0.10);
  --wa-shadow:0 10px 30px rgba(0,0,0,0.08);
  --wa-shadow2:0 14px 40px rgba(0,0,0,0.12);
  --wa-radius:18px;
  --wa-radius-sm:14px;
}

*{ box-sizing:border-box; }

.wa-root{
  font-family:Inter, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
  color:var(--wa-text);
  min-height:100vh;
  background:
    radial-gradient(1200px 700px at 30% -10%, rgba(37,211,102,0.18), transparent 55%),
    radial-gradient(1000px 600px at 90% 0%, rgba(18,140,126,0.14), transparent 55%),
    var(--wa-bg);
}

.wa-top{
  position:sticky; top:0; z-index:40;
  display:flex; align-items:center; justify-content:space-between;
  padding:12px 14px;
  background:linear-gradient(180deg, #075E54, #0B7A6E);
  color:#fff;
  box-shadow:var(--wa-shadow);
}

.wa-top__left{ display:flex; align-items:center; gap:10px; }
.wa-avatar{
  width:40px; height:40px; border-radius:999px;
  background:rgba(255,255,255,0.18);
  display:grid; place-items:center;
  font-weight:900;
}
.wa-top__titles{ line-height:1.1; }
.wa-top__title{ font-weight:900; font-size:14px; letter-spacing:0.2px; }
.wa-top__subtitle{ opacity:0.85; font-size:12px; margin-top:2px; }

.wa-top__right{ display:flex; align-items:center; gap:6px; }
.wa-icon{
  width:38px; height:38px; border-radius:999px;
  display:grid; place-items:center;
  background:transparent;
  color:#fff;
  border:none;
  cursor:pointer;
  transition: background .15s ease;
}
.wa-icon:hover{ background:rgba(255,255,255,0.14); }

.wa-tabs{
  position:sticky; top:64px; z-index:30;
  display:flex; gap:8px;
  padding:10px 12px;
  background:rgba(236,229,221,0.72);
  backdrop-filter: blur(10px);
  border-bottom:1px solid var(--wa-border);
}

.wa-chip{
  display:inline-flex; align-items:center; gap:8px;
  padding:8px 12px;
  border-radius:999px;
  border:1px solid var(--wa-border);
  background:rgba(255,255,255,0.75);
  box-shadow:0 8px 18px rgba(0,0,0,0.06);
  font-weight:800;
  font-size:12px;
  cursor:pointer;
  transition: transform .12s ease, box-shadow .12s ease;
}
.wa-chip:hover{ transform: translateY(-1px); box-shadow:0 12px 26px rgba(0,0,0,0.10); }
.wa-chip--active{
  background:rgba(37,211,102,0.18);
  border-color:rgba(37,211,102,0.35);
}
.wa-chip__icon{ display:grid; place-items:center; }

.wa-main{ max-width:1000px; margin:0 auto; padding:14px; }

.wa-card{
  background:rgba(255,255,255,0.82);
  border:1px solid var(--wa-border);
  border-radius:var(--wa-radius);
  box-shadow:var(--wa-shadow);
  overflow:hidden;
  margin:12px 0;
}
.wa-card__head{
  display:flex; justify-content:space-between; gap:12px;
  padding:12px 12px 10px;
  border-bottom:1px solid var(--wa-border);
}
.wa-card__title{ font-weight:900; font-size:13px; }
.wa-card__desc{ margin-top:6px; color:var(--wa-sub); font-size:12px; }
.wa-card__body{ padding:12px; }

.wa-badge{
  display:inline-flex; align-items:center; gap:8px;
  padding:6px 10px;
  border-radius:999px;
  background:rgba(255,255,255,0.6);
  border:1px solid var(--wa-border);
  font-weight:900;
  font-size:12px;
}
.wa-badge--work{ background:rgba(18,140,126,0.14); }
.wa-badge--ok{ background:rgba(37,211,102,0.18); }
.wa-badge--err{ background:rgba(239,68,68,0.12); }
.wa-badge__icon{ display:grid; place-items:center; }

.wa-drop{
  border:1.5px dashed rgba(18,140,126,0.28);
  border-radius:var(--wa-radius);
  padding:14px;
  background:linear-gradient(180deg, rgba(255,255,255,0.86), rgba(255,255,255,0.70));
  cursor:pointer;
  transition: border-color .15s ease, box-shadow .15s ease, transform .15s ease;
}
.wa-drop--active{ border-color:rgba(37,211,102,0.55); box-shadow:0 0 0 4px rgba(37,211,102,0.12); }
.wa-drop:hover{ transform: translateY(-1px); box-shadow:var(--wa-shadow2); }

.wa-drop__row{ display:flex; align-items:center; justify-content:space-between; gap:14px; }
.wa-drop__left{ display:flex; align-items:center; gap:12px; }
.wa-doc{
  width:56px; height:42px;
  border-radius:16px;
  display:grid; place-items:center;
  background:rgba(37,211,102,0.14);
  border:1px solid rgba(37,211,102,0.25);
  font-weight:900;
  font-size:12px;
}
.wa-drop__title{ font-weight:900; font-size:13px; }
.wa-drop__hint{ color:var(--wa-sub); font-size:12px; }

.wa-actions{ display:flex; gap:10px; flex-wrap:wrap; }

.wa-btn{
  display:inline-flex; align-items:center; gap:10px;
  padding:10px 14px;
  border-radius:999px;
  border:1px solid var(--wa-border);
  background:rgba(255,255,255,0.75);
  cursor:pointer;
  font-weight:900;
  font-size:13px;
  transition: transform .15s ease, box-shadow .15s ease, opacity .15s ease;
}
.wa-btn:disabled{ opacity:0.55; cursor:not-allowed; }
.wa-btn:hover{ transform: translateY(-1px); box-shadow:0 12px 26px rgba(0,0,0,0.12); }

.wa-btn--primary{
  background:linear-gradient(180deg, var(--wa-green), var(--wa-green-d));
  color:#083b2a;
  border-color:rgba(0,0,0,0.08);
}

.wa-btn--ghost{ background:rgba(255,255,255,0.62); }

.wa-status{
  background:rgba(255,255,255,0.64);
  border:1px solid var(--wa-border);
  border-radius:var(--wa-radius-sm);
  padding:10px 12px;
}
.wa-status__top{ display:flex; justify-content:space-between; gap:12px; font-weight:900; font-size:12px; }
.wa-status__lines{ margin-top:8px; display:flex; flex-direction:column; gap:4px; font-size:12px; color:var(--wa-sub); }

.wa-mono{ font-family:'IBM Plex Mono', ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; color:var(--wa-sub); }

.wa-row{ display:flex; gap:10px; flex-wrap:wrap; align-items:center; }
.wa-pill{
  display:inline-flex; align-items:center; gap:8px;
  padding:7px 12px;
  border-radius:999px;
  background:rgba(37,211,102,0.14);
  border:1px solid rgba(37,211,102,0.22);
  font-weight:900;
  font-size:12px;
}
.wa-pill--muted{ background:rgba(255,255,255,0.6); border-color:var(--wa-border); }
.wa-dot{ width:8px; height:8px; border-radius:999px; background:var(--wa-green); box-shadow:0 0 0 4px rgba(37,211,102,0.18); }

.wa-search{
  display:flex; align-items:center; gap:10px;
  padding:10px 12px;
  border-radius:999px;
  border:1px solid var(--wa-border);
  background:rgba(255,255,255,0.75);
}
.wa-search input{
  border:none; outline:none; width:100%;
  background:transparent;
  font-weight:700;
}

.wa-sub{ color:var(--wa-sub); font-size:12px; font-weight:900; }

.wa-rule{ display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:10px; margin-top:10px; }
@media (max-width: 900px){ .wa-rule{ grid-template-columns: 1fr; } }

.wa-radio{
  display:flex; align-items:center; gap:10px;
  padding:10px 12px;
  border-radius:16px;
  border:1px solid var(--wa-border);
  background:rgba(255,255,255,0.7);
  font-weight:900;
  font-size:12px;
  cursor:pointer;
}
.wa-radio input{ accent-color: var(--wa-green); }
.wa-radio--on{ outline: 3px solid rgba(37,211,102,0.14); border-color: rgba(37,211,102,0.22); }

.wa-list{
  border:1px solid var(--wa-border);
  border-radius:var(--wa-radius);
  overflow:hidden;
  background:rgba(255,255,255,0.62);
}
.wa-list__head{
  padding:10px 12px;
  font-weight:900;
  font-size:12px;
  color:var(--wa-sub);
  background:rgba(255,255,255,0.72);
  border-bottom:1px solid var(--wa-border);
}

.wa-item{
  display:flex; justify-content:space-between; gap:14px;
  padding:12px;
  border-top:1px solid var(--wa-border);
}
.wa-item:first-of-type{ border-top:none; }
.wa-item__left{ min-width:0; }
.wa-item__code{ font-family:'IBM Plex Mono', ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-weight:900; font-size:12px; }
.wa-item__desc{ margin-top:2px; color:var(--wa-sub); font-size:12px; word-break:break-word; }
.wa-item__meta{ margin-top:6px; color:var(--wa-sub); font-size:11px; }
.wa-item__qty{ font-weight:900; font-size:13px; white-space:nowrap; }

.wa-info{ padding:2px; }
.wa-info__title{ font-weight:900; font-size:12px; margin-bottom:8px; color:var(--wa-sub); }
.wa-info ul{ margin:0; padding-left:18px; color:var(--wa-sub); font-size:12px; }
.wa-info li{ margin:6px 0; }

.wa-fab{
  position:fixed;
  right:18px;
  bottom:18px;
  width:54px;
  height:54px;
  border-radius:999px;
  border:none;
  cursor:pointer;
  color:#083b2a;
  background:linear-gradient(180deg, var(--wa-green), var(--wa-green-d));
  box-shadow:0 16px 40px rgba(0,0,0,0.22);
  display:grid;
  place-items:center;
  transition: transform .15s ease;
}
.wa-fab:hover{ transform: translateY(-2px); }

.wa-spin{ animation: spin 1s linear infinite; }
@keyframes spin{ from{ transform: rotate(0deg);} to{ transform: rotate(360deg);} }
`;
