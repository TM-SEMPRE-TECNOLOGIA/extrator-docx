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
  ClipboardList,
  Table2,
} from "lucide-react";
import * as XLSX from "xlsx";
import JSZip from "jszip";

/**
 * Extrator de Itens DOCX - Ocean Breeze Design
 * Redesigned com foco em simplicidade e hierarquia visual
 */

const CODE_RE = /^\s*\d+(?:\.\d+)?\s*$/;
const ITEM_RE = /\(\s*I*TEM\s+([\d.]+)\s*\)/i;
const TOTAL_UNIT_RE = /TOTAL\s*\(([^)]+)\)/i;
const MEMORIAL_SECTION_RE = /MEMORIAL DE CÁLCULO E ITENS/i;

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

function naturalKey(k) {
  return k.split(/(\d+)/).map((s) => (/^\d+$/.test(s) ? Number(s) : s));
}

function naturalSort(a, b) {
  const ka = naturalKey(a);
  const kb = naturalKey(b);
  for (let i = 0; i < Math.max(ka.length, kb.length); i++) {
    const va = ka[i] ?? "";
    const vb = kb[i] ?? "";
    if (typeof va === "number" && typeof vb === "number") {
      if (va !== vb) return va - vb;
    } else {
      const cmp = String(va).localeCompare(String(vb));
      if (cmp !== 0) return cmp;
    }
  }
  return 0;
}

function buildMemorialXlsx(items, meta, filenameBase) {
  const wb = XLSX.utils.book_new();
  const itemInfo = meta?.itemInfo ?? {};
  const memorialsDoc = meta?.memorials ?? {};

  const fmtPt = (val) =>
    val.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // --- Aba 1: Itens ---
  const wsItens = XLSX.utils.aoa_to_sheet([
    ["Codigo", "Quantidade"],
    ...items.map((it) => [it.codigo, it.quantidade_raw]),
  ]);
  wsItens["!cols"] = [{ wch: 12 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, wsItens, "Itens");

  // --- Aba 2: Consolidado ---
  const consMap = new Map();
  items.forEach((it) => {
    const prev = consMap.get(it.codigo) ?? 0;
    consMap.set(it.codigo, prev + (Number.isFinite(it.quantidade) ? it.quantidade : 0));
  });
  const consRows = Array.from(consMap.entries()).sort((a, b) => naturalSort(a[0], b[0]));
  const wsCons = XLSX.utils.aoa_to_sheet([
    ["Codigo", "Quantidade Total"],
    ...consRows.map(([c, q]) => [c, fmtPt(q)]),
  ]);
  wsCons["!cols"] = [{ wch: 12 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, wsCons, "Consolidado");

  // --- Aba 3: Memorial Final ---
  const hasDocMemorial = Object.keys(memorialsDoc).length > 0;

  // Calcular Total real por código (usando Total das tabelas)
  const realSums = {};
  for (const [code, info] of Object.entries(itemInfo)) {
    let total = 0;
    for (const meas of info.measurements) {
      if (meas.totalRow) {
        total += parsePtNumber(meas.totalRow[meas.totalRow.length - 1]);
      } else if (meas.dataRows.length) {
        for (const dr of meas.dataRows) total += parsePtNumber(dr[dr.length - 1]);
      }
    }
    realSums[code] = total;
  }

  // Fallback: somas brutas
  const rawSums = {};
  items.forEach((it) => {
    rawSums[it.codigo] = (rawSums[it.codigo] ?? 0) + (Number.isFinite(it.quantidade) ? it.quantidade : 0);
  });

  const allCodes = [...new Set([...Object.keys(rawSums), ...Object.keys(itemInfo), ...Object.keys(memorialsDoc)])];
  allCodes.sort(naturalSort);

  const finalRows = [];
  if (hasDocMemorial) {
    finalRows.push(["Código", "Descrição", "Qtd Script", "Qtd Documento", "Unidade", "Status"]);
  } else {
    finalRows.push(["Código", "Descrição", "Quantidade Total", "Unidade"]);
  }

  for (const code of allCodes) {
    const totalScript = realSums[code] ?? rawSums[code] ?? 0;
    const info = itemInfo[code] ?? {};
    const docMem = memorialsDoc[code];
    const desc = info.desc ?? (docMem?.desc ?? "");
    const unit = info.unit ?? (docMem?.unit ?? "");

    if (hasDocMemorial) {
      const qtyDocStr = docMem?.qtyDoc ?? "N/A";
      const qtyDocVal = docMem ? parsePtNumber(qtyDocStr) : 0;
      const diff = Math.abs(totalScript - qtyDocVal);
      let status = diff < 0.01 ? "CONFERE" : "DIVERGENTE";
      if (!docMem) status = "SEM MEMORIAL NO DOC";
      finalRows.push([code, desc, fmtPt(totalScript), qtyDocStr, unit, status]);
    } else {
      finalRows.push([code, desc, fmtPt(totalScript), unit]);
    }
  }

  const wsFinal = XLSX.utils.aoa_to_sheet(finalRows);
  wsFinal["!cols"] = hasDocMemorial
    ? [{ wch: 10 }, { wch: 60 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 14 }]
    : [{ wch: 10 }, { wch: 60 }, { wch: 18 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, wsFinal, "Memorial Final");

  // --- Abas 4+: MemCalc por Item ---
  const sortedItems = Object.keys(itemInfo).sort(naturalSort);

  for (const code of sortedItems) {
    const info = itemInfo[code];
    const { measurements } = info;
    const colHeaders = info.colHeaders ?? [];
    const nCols = colHeaders.length || 2;
    const desc = info.desc ?? "";
    const unit = info.unit ?? "?";

    // Detectar multiplicadores
    let hasSpecial = false;
    for (const meas of measurements) {
      if (meas.totalRow) {
        const label = meas.totalRow[0].toLowerCase();
        if (label !== "total" && label.includes("total")) hasSpecial = true;
      }
    }

    const sheetName = `MemCalc_${code}`.slice(0, 31);
    const aoa = [];

    // Header com descrição + código
    const headerRow = new Array(nCols).fill("");
    headerRow[0] = `${desc} (ITEM ${code})`;
    aoa.push(headerRow);

    // Cabeçalhos de coluna
    aoa.push(colHeaders.length ? colHeaders : ["REFERÊNCIA", `TOTAL (${unit})`]);

    if (hasSpecial) {
      let grandTotal = 0;
      for (const meas of measurements) {
        for (const dr of meas.dataRows) aoa.push(dr);
        if (meas.subtotalRow) aoa.push(meas.subtotalRow);
        if (meas.totalRow) {
          aoa.push(meas.totalRow);
          grandTotal += parsePtNumber(meas.totalRow[meas.totalRow.length - 1]);
        } else if (meas.dataRows.length) {
          for (const dr of meas.dataRows) grandTotal += parsePtNumber(dr[dr.length - 1]);
        }
      }
      const tgRow = new Array(nCols).fill("");
      tgRow[0] = "TOTAL GERAL";
      tgRow[nCols - 1] = fmtPt(grandTotal);
      aoa.push(tgRow);
    } else {
      let grandTotal = 0;
      for (const meas of measurements) {
        for (const dr of meas.dataRows) {
          aoa.push(dr);
          grandTotal += parsePtNumber(dr[dr.length - 1]);
        }
      }
      // Subtotal
      const subRow = new Array(nCols).fill("");
      subRow[0] = "Subtotal";
      if (nCols >= 4) {
        let discountSum = 0;
        for (const meas of measurements) {
          for (const dr of meas.dataRows) {
            if (dr.length >= 4) discountSum += parsePtNumber(dr[dr.length - 2]);
          }
        }
        subRow[nCols - 2] = fmtPt(discountSum);
      }
      subRow[nCols - 1] = fmtPt(grandTotal);
      aoa.push(subRow);

      // Total (real)
      let realTotal = 0;
      for (const meas of measurements) {
        if (meas.totalRow) {
          realTotal += parsePtNumber(meas.totalRow[meas.totalRow.length - 1]);
        } else if (meas.dataRows.length) {
          for (const dr of meas.dataRows) realTotal += parsePtNumber(dr[dr.length - 1]);
        }
      }
      const totRow = new Array(nCols).fill("");
      totRow[0] = "Total";
      totRow[nCols - 1] = fmtPt(realTotal);
      aoa.push(totRow);
    }

    const wsItem = XLSX.utils.aoa_to_sheet(aoa);
    wsItem["!cols"] = colHeaders.map(() => ({ wch: 20 }));
    if (wsItem["!cols"].length) wsItem["!cols"][0] = { wch: 40 };
    XLSX.utils.book_append_sheet(wb, wsItem, sheetName);
  }

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

  /** @type {Item[]} */
  const results = [];
  /** @type {string[]} */
  const ignored = [];
  let itensTables = 0;
  const memorials = {};
  const itemInfo = {};
  let inCalcSection = false;

  // Iterar na ordem do documento (parágrafos + tabelas)
  const body = xml.getElementsByTagName("w:body")[0];
  if (!body) throw new Error("Não encontrou w:body no XML.");

  const children = Array.from(body.childNodes);

  for (const child of children) {
    // Detectar parágrafo com header de seção
    if (child.nodeName === "w:p") {
      const pText = xmlTextOf(child).toUpperCase();
      if (MEMORIAL_SECTION_RE.test(pText)) {
        inCalcSection = true;
      }
      continue;
    }

    // Processar tabelas
    if (child.nodeName !== "w:tbl") continue;

    const tbl = child;
    const rows = Array.from(tbl.getElementsByTagName("w:tr"));
    if (!rows.length) continue;

    const headerCells = Array.from(rows[0].getElementsByTagName("w:tc"));
    const headerTexts = headerCells.map((tc) => xmlTextOf(tc));
    const headerJoined = headerTexts.join(" ");

    // Detectar formato
    const mNew = ITEM_RE.exec(headerJoined);
    const isItens = headerTexts.some((t) => norm(t).toLowerCase() === "itens");

    let fmt = "none";
    let codeDetected = null;

    if (mNew) {
      fmt = "new";
      codeDetected = mNew[1].trim();
    } else if (isItens && rows.length > 1) {
      const r1Cells = Array.from(rows[1].getElementsByTagName("w:tc")).map((tc) => xmlTextOf(tc));
      if (r1Cells.length === 4 && CODE_RE.test(r1Cells[0]) && /^[\d.,]+$/.test(r1Cells[2])) {
        fmt = "memorial";
        codeDetected = r1Cells[0];
      } else {
        fmt = "old";
      }
    } else if (isItens) {
      fmt = "old";
    }

    if (fmt === "none") continue;
    if (inCalcSection && fmt === "new") continue;

    itensTables += 1;

    if (fmt === "old") {
      rows.slice(1).forEach((tr, rOffset) => {
        const rNumber = rOffset + 2;
        const tcs = Array.from(tr.getElementsByTagName("w:tc"));
        if (!tcs.length) { ignored.push(`T${itensTables} L${rNumber}: skip_empty_row`); return; }
        const cellsText = tcs.map((tc) => xmlTextOf(tc));
        const code = norm(cellsText[0] ?? "");
        const desc = norm(cellsText[1] ?? "");
        if (!code || code.toUpperCase() === "#N/D") { ignored.push(`T${itensTables} L${rNumber}: skip_code_empty_or_ND`); return; }
        if (!CODE_RE.test(code)) { ignored.push(`T${itensTables} L${rNumber}: skip_code_invalid ${code}`); return; }
        const qtyRaw = pickQuantityFromRow(cellsText);
        if (!qtyRaw || qtyRaw.toUpperCase() === "#N/D") { ignored.push(`T${itensTables} L${rNumber}: skip_qty_empty_or_ND ${code}`); return; }
        results.push({ codigo: code, descricao: desc, quantidade_raw: qtyRaw, quantidade: parsePtNumber(qtyRaw), origem: `T${itensTables}/L${rNumber}` });
      });
    } else if (fmt === "new") {
      // Capturar info completa
      const descFull = norm(headerTexts[0]).replace(/\s*\(I*TEM\s+[\d.]+\)\s*$/i, "").trim();
      const colHeaders = rows.length > 1
        ? Array.from(rows[1].getElementsByTagName("w:tc")).map((tc) => xmlTextOf(tc))
        : [];
      let unit = "?";
      for (const ch of colHeaders) {
        const um = TOTAL_UNIT_RE.exec(ch);
        if (um) { unit = um[1].trim(); break; }
      }

      if (!itemInfo[codeDetected]) {
        itemInfo[codeDetected] = { desc: descFull, unit, colHeaders, measurements: [] };
      }

      const dataRows = [];
      let subtotalRow = null;
      let totalRow = null;

      rows.slice(2).forEach((tr, rOffset) => {
        const rNum = rOffset + 3;
        const tcs = Array.from(tr.getElementsByTagName("w:tc"));
        if (!tcs.length) { ignored.push(`T${itensTables} L${rNum}: skip_empty_row_NEW`); return; }
        const cellsText = tcs.map((tc) => xmlTextOf(tc));
        const label = cellsText[0].toLowerCase();

        if (label.includes("subtotal")) { subtotalRow = cellsText; return; }
        if (label.includes("total")) { totalRow = cellsText; return; }

        const qty = norm(cellsText[cellsText.length - 1]);
        if (!qty || qty.toUpperCase() === "#N/D") { ignored.push(`T${itensTables} L${rNum}: skip_qty_empty_or_ND_NEW ${codeDetected}`); return; }
        if (/^[\d.,]+$/.test(qty)) {
          results.push({ codigo: codeDetected, descricao: descFull, quantidade_raw: qty, quantidade: parsePtNumber(qty), origem: `T${itensTables}/L${rNum}` });
          dataRows.push(cellsText);
        } else {
          ignored.push(`T${itensTables} L${rNum}: skip_qty_invalid_NEW ${qty}`);
        }
      });

      itemInfo[codeDetected].measurements.push({ tableIdx: itensTables, dataRows, subtotalRow, totalRow });
    } else if (fmt === "memorial") {
      const r1Cells = Array.from(rows[1].getElementsByTagName("w:tc")).map((tc) => xmlTextOf(tc));
      memorials[r1Cells[0]] = { desc: r1Cells[1], qtyDoc: r1Cells[2], unit: r1Cells[3] };
    }
  }

  const meta = {
    tables_total: itensTables,
    itens_tables: itensTables,
    rows_extracted: results.length,
    rows_ignored: ignored.length,
    ignored_details: ignored,
    memorials,
    itemInfo,
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

  const downloadMemorial = useCallback(() => {
    if (!items.length || !meta) return;
    buildMemorialXlsx(items, meta, `memorial_${safeBaseName(file?.name)}`);
  }, [items, meta, file]);

  const memorialItemCount = useMemo(() => {
    return Object.keys(meta?.itemInfo ?? {}).length;
  }, [meta]);

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
                  {memorialItemCount > 0 && (
                    <StatCard label="Itens no Memorial" value={fmtInt(memorialItemCount)} sub="Memorial de Cálculo" />
                  )}
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

              {/* MEMORIAL DE CÁLCULO */}
              {memorialItemCount > 0 && (
                <Section
                  title="Memorial de Cálculo"
                  desc={`${fmtInt(memorialItemCount)} itens com medições detalhadas`}
                  right={
                    <Badge kind="ok" icon={<ClipboardList size={16} />}>
                      {Object.keys(meta?.memorials ?? {}).length > 0 ? "Com comparação" : "Gerado do zero"}
                    </Badge>
                  }
                >
                  <div className="tm-preview">
                    {Object.entries(meta?.itemInfo ?? {})
                      .sort(([a], [b]) => naturalSort(a, b))
                      .slice(0, 6)
                      .map(([code, info]) => {
                        const totalMeas = info.measurements?.reduce((s, m) => s + m.dataRows.length, 0) ?? 0;
                        return (
                          <div key={code} className="tm-preview__item">
                            <div className="tm-preview__meta">
                              <div className="tm-preview__code">{code}</div>
                              <div className="tm-preview__desc">
                                {(info.desc || "(sem descrição)").slice(0, 80)}
                                {(info.desc || "").length > 80 ? "..." : ""}
                              </div>
                            </div>
                            <div className="tm-preview__qty" style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: "2px" }}>
                              <span>{totalMeas} medições</span>
                              <span style={{ fontSize: "11px", opacity: 0.7 }}>{info.unit ?? "?"}</span>
                            </div>
                          </div>
                        );
                      })}
                  </div>

                  <div className="tm-actions" style={{ marginTop: "16px" }}>
                    <button type="button" onClick={downloadMemorial} className="tm-btn tm-btn--primary">
                      <Table2 size={16} />
                      Baixar Memorial Completo
                    </button>
                  </div>
                </Section>
              )}

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
