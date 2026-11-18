// Excel (exceljs) e Word (docx) generati nel browser: nessun file sul server.
// Le due librerie si caricano con import() dinamico solo al primo download,
// così il bundle iniziale resta quello di v5 (docs/mockup-ai-agent.md §2.2).
// Le funzioni build* ricevono il modulo, così i test le provano in Node.
import type * as ExcelNS from "exceljs";
import type * as DocxNS from "docx";
import type { MockupAgentManifest } from "./contract.ts";
import { columnIndex, isNormattivaUrl, monthLabel, numbersOf, parseNumber, type AgentTable } from "./compute.ts";
import type { DocModel } from "./document.ts";
import type { Workspace } from "./runtime.ts";

type ExcelModule = typeof ExcelNS;
type DocxModule = typeof DocxNS;

const euro = '#,##0.00 "€"';
const pct = "0.0%";
const moneyName = /import|accert|riscoss|incass|total|€|euro|valore|spesa|2025|2026/i;

/** Lettera di colonna Excel (0 → A, 26 → AA). */
export function columnLetter(index: number): string {
  let n = index + 1, out = "";
  while (n > 0) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26); }
  return out;
}
const isNumericColumn = (table: AgentTable, i: number) =>
  table.rows.length > 0 && numbersOf(table, table.columns[i]).length === table.rows.length;
const sheetRef = (name: string) => `'${name.replace(/'/g, "''")}'`;

/** Fogli fissi del file: il foglio dati non può chiamarsi come uno di loro. */
const reservedSheets = ["Confronto", "Mensile", "Statistiche", "KPI"];
/**
 * Nome valido per un foglio Excel: senza / \ : [ ] * ?, senza apostrofi ai bordi,
 * al massimo 31 caratteri, diverso dai fogli fissi (senza distinzione di maiuscole).
 */
export function sheetName(raw: string | undefined, fallback = "Dati"): string {
  let name = (raw ?? "").replace(/[/\\:[\]*?]/g, " ").replace(/\s+/g, " ").trim().replace(/^'+|'+$/g, "").trim().slice(0, 31).trim();
  if (!name || /^history$/i.test(name)) name = fallback;
  if (reservedSheets.some((r) => r.toLowerCase() === name.toLowerCase())) name = `${name.slice(0, 25)} (dati)`;
  return name;
}

export type WorkbookInput = { manifest: MockupAgentManifest; workspace: Workspace; dataSheet?: string };

/** Fogli «Dati» (o il nome dato), «Confronto» e «Mensile» se presenti, «Statistiche», «KPI»: formule vere. */
export async function buildWorkbook(ExcelJS: ExcelModule, input: WorkbookInput): Promise<ArrayBuffer> {
  const table = input.workspace.table;
  if (!table) throw Error("mockup_agent_no_table");
  const wb = new ExcelJS.Workbook();
  wb.creator = "Prototipo Floware (dati dimostrativi)";
  wb.created = new Date(Date.UTC(2026, 0, 1));
  const dataName = sheetName(input.dataSheet);
  const data = wb.addWorksheet(dataName, { views: [{ state: "frozen", ySplit: 1 }] });
  const numeric = table.columns.map((_, i) => isNumericColumn(table, i));
  data.addRow(table.columns);
  for (const row of table.rows) data.addRow(row.map((cell, i) => (numeric[i] ? parseNumber(cell) : cell)));
  data.getRow(1).font = { bold: true };
  table.columns.forEach((name, i) => {
    const col = data.getColumn(i + 1);
    col.width = Math.min(Math.max(name.length + 4, 14), 40);
    if (numeric[i] && moneyName.test(name)) col.numFmt = euro;
  });
  const last = table.rows.length + 1;
  const range = (i: number) => `${sheetRef(dataName)}!$${columnLetter(i)}$2:$${columnLetter(i)}$${last}`;

  const cmp = input.workspace.comparison;
  if (cmp) {
    const sheet = wb.addWorksheet("Confronto", { views: [{ state: "frozen", ySplit: 1 }] });
    const headers = ["Voce", cmp.column, ...(cmp.compareColumn ? [cmp.compareColumn, "Δ %"] : []), ...(cmp.baseColumn ? [cmp.baseColumn, "Capacità di riscossione"] : [])];
    sheet.addRow(headers).font = { bold: true };
    cmp.rows.forEach((r, k) => {
      const n = k + 2, cells: (string | number | { formula: string; result?: number })[] = [r.label, r.value];
      if (cmp.compareColumn) cells.push(r.compare, { formula: `IF(C${n}=0,"",B${n}/C${n}-1)`, ...(r.delta !== null ? { result: r.delta } : {}) });
      if (cmp.baseColumn) {
        const base = r.base ?? 0;
        const b = cmp.compareColumn ? "E" : "C";
        cells.push(base, { formula: `IF(${b}${n}=0,"",B${n}/${b}${n})`, ...(r.capacity !== null ? { result: r.capacity } : {}) });
      }
      sheet.addRow(cells);
    });
    headers.forEach((h, i) => {
      const col = sheet.getColumn(i + 1);
      col.width = i === 0 ? 34 : 18;
      if (h === "Δ %" || h.startsWith("Capacità")) col.numFmt = pct;
      else if (i > 0) col.numFmt = euro;
    });
  }

  const chart = input.workspace.chart;
  if (chart) {
    const sheet = wb.addWorksheet("Mensile", { views: [{ state: "frozen", ySplit: 1 }] });
    sheet.addRow(["Mese", chart.column, ...(chart.compareColumn ? [chart.compareColumn, "Δ %"] : [])]).font = { bold: true };
    chart.points.forEach((p, k) => {
      const n = k + 2;
      sheet.addRow([monthLabel(p.month), p.value, ...(chart.compareColumn ? [p.compare ?? 0, { formula: `IF(C${n}=0,"",B${n}/C${n}-1)`, ...(p.compare ? { result: p.value / p.compare - 1 } : {}) }] : [])]);
    });
    sheet.getColumn(1).width = 14;
    [2, 3].forEach((i) => { sheet.getColumn(i).numFmt = euro; sheet.getColumn(i).width = 18; });
    sheet.getColumn(4).numFmt = pct;
  }

  const stats = wb.addWorksheet("Statistiche", { views: [{ state: "frozen", ySplit: 1 }] });
  stats.addRow(["Colonna", "Somma", "Media", "Minimo", "Massimo", "Conteggio"]).font = { bold: true };
  table.columns.forEach((name, i) => {
    if (!numeric[i]) return;
    const r = range(i);
    stats.addRow([name, { formula: `SUM(${r})` }, { formula: `AVERAGE(${r})` }, { formula: `MIN(${r})` }, { formula: `MAX(${r})` }, { formula: `COUNT(${r})` }]);
  });
  const groupCol = table.columns.findIndex((_, i) => !numeric[i] && !/mese|data|period/i.test(table.columns[i]));
  if (groupCol >= 0) {
    stats.addRow([]);
    stats.addRow([`Righe per ${table.columns[groupCol]}`, "Conteggio"]).font = { bold: true };
    for (const value of [...new Set(table.rows.map((r) => r[groupCol]))].slice(0, 40))
      stats.addRow([value, { formula: `COUNTIF(${range(groupCol)},"${value.replace(/"/g, '""')}")` }]);
  }
  stats.getColumn(1).width = 30;
  [2, 3, 4, 5].forEach((i) => { stats.getColumn(i).width = 16; stats.getColumn(i).numFmt = "#,##0.00"; });

  const kpi = wb.addWorksheet("KPI");
  kpi.addRow(["Indicatore", "Valore", "Formula"]).font = { bold: true };
  for (const k of input.manifest.kpis) {
    // Formule del contratto (AgentFormula): `by` è il denominatore, il riferimento o la colonna data.
    const c = columnIndex(table.columns, k.formula.column), d = columnIndex(table.columns, k.formula.by ?? undefined);
    if (c < 0) continue;
    const r = range(c), other = d >= 0 ? range(d) : "";
    const cmp = { lt: "<", lte: "<=", gt: ">", gte: ">=" } as const;
    const f = k.formula;
    const formula =
      f.op === "sum" ? `SUM(${r})`
      : f.op === "avg" ? `AVERAGE(${r})`
      : f.op === "count" ? `COUNTA(${r})`
      : f.op === "ratio" && other ? `SUM(${r})/SUM(${other})`
      : f.op === "delta_pct" && other ? `SUM(${r})/SUM(${other})-1`
      : f.op === "threshold" && f.limit !== null && f.comparison ? `COUNTIF(${r},"${cmp[f.comparison]}${f.limit}")/COUNT(${r})`
      : null;
    const fixed = input.workspace.kpis?.find((x) => x.id === k.id)?.raw ?? null;
    const row = kpi.addRow([k.label, formula ? { formula, ...(fixed !== null ? { result: fixed } : {}) } : fixed, formula ? `=${formula}` : "valore calcolato nel prototipo (ultimo mese)"]);
    row.getCell(2).numFmt = k.unit === "€" ? euro : k.unit === "%" ? pct : "#,##0.0";
  }
  // Indicatori calcolati dal pannello oltre a quelli del manifesto (rotazione, andamento peggiore).
  for (const extra of input.workspace.kpis?.filter((x) => !input.manifest.kpis.some((k) => k.id === x.id)) ?? []) {
    const row = kpi.addRow([extra.label, extra.unit === "testo" ? extra.value : extra.raw, "valore calcolato nel prototipo"]);
    if (extra.unit === "%") row.getCell(2).numFmt = pct;
  }
  kpi.getColumn(1).width = 36; kpi.getColumn(2).width = 18; kpi.getColumn(3).width = 60;
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}

/** Word della determinazione dal modello comune (document.ts). */
export async function buildDocx(docx: DocxModule, model: DocModel): Promise<Blob | Uint8Array> {
  const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, ExternalHyperlink } = docx;
  const children: InstanceType<typeof Paragraph>[] = [
    ...model.header.map((h, i) => new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: i === 0, size: i === 0 ? 28 : 20 })] })),
    new Paragraph({ text: "" }),
    new Paragraph({ alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: model.title, bold: true })] }),
    new Paragraph({ spacing: { before: 240, after: 240 }, children: [new TextRun({ text: model.subject, bold: true })] }),
  ];
  for (const block of model.blocks) {
    if (block.heading) children.push(new Paragraph({ alignment: block.heading === block.heading.toUpperCase() ? AlignmentType.CENTER : AlignmentType.LEFT, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: block.heading, bold: true })] }));
    for (const p of block.paragraphs) children.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 120 }, ...(block.list ? { bullet: { level: 0 } } : {}), children: [new TextRun(p)] }));
  }
  if (model.notes.length) {
    children.push(new Paragraph({ spacing: { before: 360, after: 120 }, children: [new TextRun({ text: "Note e riferimenti normativi", bold: true })] }));
    model.notes.forEach((n, i) => children.push(new Paragraph({
      spacing: { after: 80 },
      children: [
        new TextRun(`[${i + 1}] ${n.citation} — `),
        ...(isNormattivaUrl(n.url) ? [new ExternalHyperlink({ link: n.url, children: [new TextRun({ text: "testo su Normattiva", style: "Hyperlink" })] })] : [new TextRun("link non disponibile")]),
        ...(n.verified ? [new TextRun(` (testo vigente verificato il ${n.verified})`)] : []),
      ],
    })));
  }
  children.push(new Paragraph({ spacing: { before: 480 }, children: [new TextRun({ text: model.footer, italics: true, size: 18 })] }));
  const doc = new Document({ creator: "Prototipo Floware", title: model.subject, sections: [{ children }] });
  return typeof Blob !== "undefined" ? Packer.toBlob(doc) : Packer.toBuffer(doc);
}

// ---- Solo browser ----------------------------------------------------------

const xlsxType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const docxType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export function downloadBlob(data: Blob | ArrayBuffer | Uint8Array, name: string, type: string) {
  const blob = data instanceof Blob ? data : new Blob([data as BlobPart], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.rel = "noopener";
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
const moduleOf = <T,>(mod: T): T => (mod as T & { default?: T }).default ?? mod;

export async function downloadWorkbook(input: WorkbookInput, name: string) {
  const ExcelJS = moduleOf(await import("exceljs"));
  downloadBlob(await buildWorkbook(ExcelJS, input), name, xlsxType);
}
export async function downloadDocx(model: DocModel, name: string) {
  const docx = await import("docx");
  downloadBlob(await buildDocx(docx, model), name, docxType);
}
export const fileName = (base: string, ext: string) =>
  base.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60) + "." + ext;
