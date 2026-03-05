import fs from 'fs';
import JSZip from 'jszip';
import { DOMParser } from 'xmldom';

const CODE_RE = /^\s*\d+(?:\.\d+)?\s*$/;

function norm(s) {
    return (s ?? "").replace(/\u00A0/g, " ").trim();
}

function parsePtNumber(s) {
    const t = norm(s);
    if (!t) return NaN;
    const cleaned = t.replace(/\./g, "").replace(/,/g, ".");
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : NaN;
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

function xmlTextOf(node) {
    // xmldom specific lookup
    const ts = node.getElementsByTagName("w:t");
    let out = "";
    for (let i = 0; i < ts.length; i++) {
        // xmldom textContent usage
        out += ts[i].textContent ?? "";
    }
    return norm(out);
}

async function extractItemsFromDocx(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const docFile = zip.file("word/document.xml");
    if (!docFile) throw new Error("No word/document.xml");

    const docXml = await docFile.async("string");
    const parser = new DOMParser();
    const xml = parser.parseFromString(docXml, "application/xml");

    const tables = Array.from(xml.getElementsByTagName("w:tbl"));
    const results = [];

    tables.forEach((tbl, tIndex) => {
        const rows = Array.from(tbl.getElementsByTagName("w:tr"));
        if (!rows.length) return;

        // Check header
        const headerCells = Array.from(rows[0].getElementsByTagName("w:tc"));
        const headerTexts = headerCells.map(tc => xmlTextOf(tc));
        const isItens = headerTexts.some(t => norm(t).toLowerCase() === "itens");

        if (!isItens) return;

        rows.slice(1).forEach((tr, rOffset) => {
            const rNumber = rOffset + 2;
            const tNumber = tIndex + 1;

            const tcs = Array.from(tr.getElementsByTagName("w:tc"));
            if (!tcs.length) return;

            const cellsText = tcs.map(tc => xmlTextOf(tc));
            const code = norm(cellsText[0] ?? "");
            const desc = norm(cellsText[1] ?? "");

            if (!code || code.toUpperCase() === "#N/D") return;
            if (!CODE_RE.test(code)) return;

            const qtyRaw = pickQuantityFromRow(cellsText);
            const qty = parsePtNumber(qtyRaw);

            results.push({
                codigo: code,
                descricao: desc,
                quantidade_raw: qtyRaw,
                quantidade: qty,
                origem: `T${tNumber}/L${rNumber}`
            });
        });
    });

    return results;
}

function aggregateItems(items, rule) {
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
                key,
                codigo: rule === "desc_only" ? "" : it.codigo,
                descricao: rule === "code_only" ? "" : it.descricao,
                quantidade: safeQ,
                origens: [it.origem]
            });
        } else {
            prev.quantidade += safeQ;
            prev.origens.push(it.origem);
        }
    });

    return Array.from(map.values());
}

async function run() {
    try {
        const data = fs.readFileSync("RELATÓRIO FOTOGRÁFICO - RESPLENDOR - LEVANTAMENTO PREVENTIVO.docx");
        console.log("File loaded. Extracting...");

        const items = await extractItemsFromDocx(data);
        console.log(`Extracted ${items.length} items.`);

        // Filter for 19.53
        const targetItems = items.filter(it => it.codigo.includes("19.53"));
        console.log("\n--- Raw Items containing 19.53 ---");
        targetItems.forEach(it => {
            console.log(`[${it.origem}] Code: '${it.codigo}' Desc: '${it.descricao}' Qty: ${it.quantidade}`);
            console.log(`Code chars: ${it.codigo.split('').map(c => c.charCodeAt(0)).join(',')}`);
            console.log(`Desc chars: ${it.descricao.split('').map(c => c.charCodeAt(0)).join(',')}`);
        });

        const aggregated = aggregateItems(items, 'code_desc');
        const aggTarget = aggregated.filter(it => it.codigo.includes("19.53"));

        console.log("\n--- Aggregated Items containing 19.53 ---");
        aggTarget.forEach(it => {
            console.log(`Key: '${it.key}'`);
            console.log(`Key chars: ${it.key.split('').map(c => c.charCodeAt(0)).join(',')}`);
            console.log(`Code: '${it.codigo}' Desc: '${it.descricao}' Qty: ${it.quantidade}`);
            console.log(`Origens: ${it.origens.join(', ')}`);
        });

        if (aggTarget.length > 1) {
            console.log("\nISSUE REPRODUCED: Multiple entries for 19.53 found after aggregation.");
        } else {
            console.log("\nISSUE NOT REPRODUCED: Only one entry for 19.53 found.");
        }

    } catch (e) {
        console.error(e);
    }
}

run();
