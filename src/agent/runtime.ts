// Orchestratore dei percorsi guidati (docs/mockup-ai-agent.md §2.2).
// Esegue i passi del manifesto nell'ordine: i tool del gateway passano dal
// client, gli altri girano qui con le funzioni pure di compute.ts. Ogni passo
// del gateway che non riceve un output valido (12 s, 429, 503/504, rete,
// risposta non valida) usa il copione del manifesto e lo dichiara
// («esempio preparato»). Nessun DOM: lo stato va all'interfaccia con
// `onChange`, così il runtime si prova con un gateway finto.
//
// Chiavi di input dei passi locali (minuscole con trattini bassi, come vuole il
// parser del manifesto):
// Le chiavi canoniche sono in `mockupAgentStepInputKeys` (contratto); gli alias
// italiani del modello (`importo`, `tipologia`, `fonte`…) diventano canonici con
// `canonicalStepInput`, sia nel parser del manifesto sia qui.
// - rule.check_threshold: amount, category, item_key, fact_key. Non usa mai il
//   copione: importo illeggibile, categoria senza soglia o soglia assente danno
//   un errore esplicito. La soglia viene dai `facts` di una citazione ricevuta
//   dal vivo; se la ricerca è passata al copione il passo è «esempio preparato».
// - data.compute: op = rotation | compare | month_over_month | normalize | sum | avg
//   | count | ratio | delta_pct | group_by_month | threshold | distinct_pct;
//   table, column, compare_column, base_column, date_column, label_columns,
//   operator, operator_column, months, threshold, delta_limit, capacity_limit, label, unit
// - kpi.show: table, group_column, column, compare_column, worst_label, distinct_column, distinct_label
// - chart.show: table, date_column, column, compare_column, title
// - file.xlsx: table, sheet · file.docx: nessuna
// `table` è l'id di un documento di esempio con tabella; senza, la tabella in uso.
import {
  canonicalStepInput,
  isGatewayTool,
  isModelTool,
  itemKeyOf,
  mockupAgentMessages,
  mockupAgentPolicy,
  type MockupAgentCitation,
  type MockupAgentManifest,
  type MockupAgentStep,
} from "./contract.ts";
import {
  answerDataQuestion,
  checkRotation,
  checkThreshold,
  compareRows,
  evaluateFormula,
  findAnomalies,
  formatPct,
  formatSignedPct,
  formatValue,
  groupByMonth,
  isNormattivaUrl,
  kpiValue,
  monthOverMonth,
  normalizeTable,
  parseNumber,
  resolveInput,
  worstGroup,
  type AgentTable,
  type Anomaly,
  type ComparisonRow,
  type MonthPoint,
  type RotationResult,
  type ThresholdFact,
  type ThresholdResult,
} from "./compute.ts";
import { requestKey, type GatewayClient, type GatewayFailure } from "./gateway.ts";

/** Citazione del catalogo, con i campi in più degli output di catalog.search e catalog.cite. */
export type Citation = MockupAgentCitation & { summary?: string; text?: string; facts?: ThresholdFact[] };
export type Draft = { title?: string; sections: { heading: string; paragraphs: string[] }[] };
export type StepStatus = "pending" | "running" | "done" | "error";
export type StepView = {
  id: string;
  label: string;
  tool: string;
  status: StepStatus;
  detail?: string;
  ms?: number;
  /** Vero se il passo ha usato il copione del manifesto (Decisione 1). */
  prepared?: boolean;
  /** Inizio del passo in corso: il pannello mostra i secondi trascorsi. */
  startedAt?: number;
  /** Passo con il modello (può durare decine di secondi). */
  model?: boolean;
  citations?: Citation[];
};
export type KpiView = { id: string; label: string; value: string; raw: number | null; unit: "€" | "%" | "n" | "testo" };
export type Workspace = {
  variables: Record<string, string>;
  documentId?: string;
  documentTitle?: string;
  sampleDocumentId?: string;
  citations: Citation[];
  table: AgentTable | null;
  tableSource?: string;
  comparison?: { rows: ComparisonRow[]; anomalies: Anomaly[]; column: string; compareColumn?: string; baseColumn?: string };
  /** `prepared`: soglia presa da un esempio preparato, da verificare sul testo vigente. */
  threshold?: ThresholdResult & { prepared?: boolean };
  rotation?: RotationResult;
  draft?: Draft;
  kpis?: KpiView[];
  chart?: { title: string; points: MonthPoint[]; column: string; compareColumn?: string };
  files: ("xlsx" | "docx" | "pdf")[];
  /** Nome del foglio dati dell'Excel (input `sheet` di file.xlsx). */
  sheet?: string;
  notice?: string;
  /** Risposte dal vivo ancora disponibili (contatore del pannello); null = sconosciuto. */
  remaining?: number | null;
};
export type RunState = { pathId: string | null; steps: StepView[]; running: boolean };

export type AskOutcome =
  | { kind: "path"; pathId: string; prepared?: boolean }
  /** `suggest`: percorso pertinente da proporre (si avvia solo se l'utente lo sceglie). */
  | { kind: "answer"; text: string; citations: Citation[]; local?: boolean; suggest?: { pathId: string; prepared?: boolean } }
  | { kind: "hint"; text: string; error?: boolean; suggest?: { pathId: string; prepared?: boolean } };

/** Una domanda (e non un comando): «posso…?», «quale articolo…», «come si…». */
export function isQuestion(text: string): boolean {
  const s = text.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  return /\?\s*$/.test(s) || /^(posso|si puo|e possibile|quale|quali|qual|quanto|quanti|quando|come|perche|cosa|che cosa|chi|dove|serve|devo|bisogna|e (ammesso|consentito|obbligatorio))\b/.test(s);
}
/** Risposta del copione a una domanda quando il modello non risponde: propone il percorso, non lo avvia. */
export const preparedQuestionHint = (label: string) => `Non riesco a rispondere dal vivo in questo momento. Il percorso «${label}» risponde a questa domanda con i calcoli sui dati e la norma citata.`;
/** Il modello ha indicato solo un percorso per una domanda: lo si propone. */
export const questionPathHint = (label: string) => `Il percorso «${label}» risponde a questa domanda con i calcoli sui dati e la norma citata.`;

/** Il copione può usare i `{{chiave}}` delle variabili: si sostituiscono come stringhe JSON. */
export function fillFallback(fallback: string, values: Record<string, string>): string {
  return fallback.replace(/\{\{\s*([a-z0-9_]+)\s*\}\}/gi, (_, key: string) => (values[key] === undefined ? `[${key}]` : JSON.stringify(values[key]).slice(1, -1)));
}
const parseJson = (value: unknown): unknown => {
  if (typeof value !== "string") return value;
  try { return JSON.parse(value); } catch { return { detail: value }; }
};
const isStrings = (v: unknown): v is string[] => Array.isArray(v) && v.every((x) => typeof x === "string");
const isCitation = (v: unknown): v is Citation => {
  const c = v as Citation;
  return !!c && typeof c === "object" && typeof c.itemKey === "string" && typeof c.citation === "string" && isNormattivaUrl(c.url);
};
const isFact = (f: unknown): f is ThresholdFact => !!f && typeof f === "object" && ["key", "label", "value"].every((k) => typeof (f as Record<string, unknown>)[k] === "string");
const cleanCitation = (c: Citation): Citation => ({
  sourceKey: typeof c.sourceKey === "string" ? c.sourceKey : "normattiva",
  itemKey: c.itemKey,
  citation: c.citation,
  title: typeof c.title === "string" ? c.title : c.citation,
  url: c.url,
  verifiedAt: typeof c.verifiedAt === "number" ? c.verifiedAt : null,
  ...(typeof c.summary === "string" ? { summary: c.summary } : {}),
  ...(typeof c.text === "string" ? { text: c.text } : {}),
  ...(Array.isArray(c.facts) ? { facts: c.facts.filter(isFact) } : {}),
});

type Normalized =
  | { tool: "catalog.search"; hits: Citation[] }
  | { tool: "catalog.cite"; item: Citation }
  | { tool: "document.read"; summary: string; sections: string[]; fields: { key: string; label: string; example: string }[] }
  | { tool: "data.extract_table"; columns: string[]; rows: string[][] }
  | { tool: "draft.compose"; draft: Draft };

/**
 * Valida l'output di un tool del gateway, o il suo copione, nella forma di
 * `MockupAgentRunOutput` (il campo `tool` può mancare nel copione). null = non valido.
 */
export function validToolOutput(tool: string, raw: unknown): Normalized | null {
  const v = parseJson(raw) as Record<string, unknown> | null;
  if (!v || typeof v !== "object" || Array.isArray(v) || (v.tool !== undefined && v.tool !== tool)) return null;
  switch (tool) {
    case "catalog.search":
      return Array.isArray(v.hits) && v.hits.every(isCitation) ? { tool, hits: v.hits.slice(0, 5).map(cleanCitation) } : null;
    case "catalog.cite":
      return isCitation(v.item) ? { tool, item: cleanCitation(v.item) } : null;
    case "document.read": {
      const fields = Array.isArray(v.fields) ? v.fields.filter((f) => f && typeof f === "object" && typeof f.key === "string") : null;
      if (typeof v.summary !== "string" || !isStrings(v.sections) || !fields) return null;
      return { tool, summary: v.summary, sections: v.sections, fields: fields.map((f) => ({ key: f.key, label: String(f.label ?? f.key), example: String(f.example ?? "") })) };
    }
    case "data.extract_table": {
      const { columns, rows } = v as { columns: unknown; rows: unknown };
      if (!isStrings(columns) || !columns.length || columns.length > mockupAgentPolicy.extractTableMaxColumns) return null;
      if (!Array.isArray(rows) || rows.length > mockupAgentPolicy.extractTableMaxRows || !rows.every((r) => isStrings(r) && r.length === columns.length)) return null;
      return { tool, columns, rows };
    }
    case "draft.compose": {
      const sections = v.sections;
      if (!Array.isArray(sections) || !sections.length || !sections.every((s) => s && typeof s.heading === "string" && typeof s.text === "string")) return null;
      const paragraphs = (text: string) => text.split(/\n+/).map((p) => p.trim()).filter(Boolean);
      return { tool, draft: { title: typeof v.title === "string" && v.title.trim() ? v.title : undefined, sections: sections.map((s) => ({ heading: s.heading, paragraphs: paragraphs(s.text) })) } };
    }
    default:
      return null;
  }
}

/** Avviso da mostrare accanto ai passi quando il gateway non risponde dal vivo. */
export const networkFallback = "Il servizio dell'assistente non è raggiungibile: proseguo con l'esempio preparato.";
export function failureNotice(f: GatewayFailure): string | undefined {
  if (f.status === 429) return f.message && /alle \d/.test(f.message) ? `${mockupAgentMessages.limitFallback} ${f.message}` : mockupAgentMessages.limitFallback;
  if (f.status === 401) return mockupAgentMessages.expired;
  if (f.timedOut || (f.status !== undefined && f.status >= 500)) return mockupAgentMessages.modelFallback;
  if (f.status === undefined) return networkFallback;
  return undefined;
}
/** Testo del passo della soglia quando la soglia viene da un esempio preparato. */
export const preparedThresholdNote = "soglia dell'esempio preparato: verificala sul testo vigente";

/** Esito di un tool locale: `error` = passo in errore con il messaggio, senza copione. */
type LocalOutcome = { detail: string; error?: boolean; prepared?: boolean };

/** Avviso del pannello minimo quando il manifesto non arriva dal gateway. */
export const offlineNotice = "Assistente dal vivo non disponibile, esempio preparato";
/** Attesa prima del tentativo `n` (1, 2, 3) di caricare il manifesto: backoff esponenziale con jitter ±50%. */
export function manifestRetryDelay(n: number, random: () => number = Math.random): number {
  return Math.round(1000 * 2 ** (n - 1) * (0.5 + random()));
}
export type ManifestOutcome =
  | { kind: "ready"; manifest: MockupAgentManifest; remaining: number | null }
  | { kind: "expired" }
  | { kind: "unavailable" }
  | { kind: "stopped" }
  /** Dopo 1 + 3 tentativi falliti con `fallback: true`: pannello minimo locale. */
  | { kind: "offline"; attempts: number };
/**
 * `POST /manifest`: con «pending» riprova ogni 5 s per 3 minuti; con un errore
 * `fallback: true` (rete, timeout, 429, 5xx) altri 3 tentativi con backoff e jitter.
 */
export async function loadManifest(client: GatewayClient, options: {
  sleep: (ms: number) => Promise<void>;
  random?: () => number;
  stopped?: () => boolean;
  onPending?: () => void;
}): Promise<ManifestOutcome> {
  let pending = 0, retries = 0;
  for (;;) {
    let result: Awaited<ReturnType<GatewayClient["manifest"]>>;
    try { result = await client.manifest(); } catch { result = { ok: false, fallback: true }; }
    if (options.stopped?.()) return { kind: "stopped" };
    if (result.ok && result.data.status === "ready" && result.data.manifest?.version === "mockup-agent-v1") {
      const left = result.data.limits?.remaining?.mockup;
      return { kind: "ready", manifest: result.data.manifest, remaining: typeof left === "number" ? left : null };
    }
    if (result.ok && result.data.status === "pending" && ++pending < 36) {
      options.onPending?.();
      await options.sleep(5000);
    } else if (!result.ok && result.status === 401) return { kind: "expired" };
    else if (!result.ok && result.fallback && retries < 3) await options.sleep(manifestRetryDelay(++retries, options.random));
    else if (!result.ok && result.fallback) return { kind: "offline", attempts: retries + 1 };
    else return { kind: "unavailable" };
    if (options.stopped?.()) return { kind: "stopped" };
  }
}

export type DemoCase = { id: string; keywords: string[]; manifest: MockupAgentManifest };
/**
 * Caso dimostrativo locale (demo.json) più vicino al prototipo, contando le
 * parole chiave nel titolo, nella descrizione e nei titoli delle pagine. null = nessuno.
 */
export function pickDemoManifest(spec: { title?: string; description?: string; pages?: { title?: string }[] }, cases: DemoCase[]): MockupAgentManifest | null {
  const text = [spec.title, spec.description, ...(spec.pages ?? []).map((p) => p.title)].filter(Boolean).join(" ").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  let best: { manifest: MockupAgentManifest; score: number } | null = null;
  for (const c of cases) {
    const score = c.keywords.filter((k) => text.includes(k)).length;
    if (score > (best?.score ?? 0)) best = { manifest: c.manifest, score };
  }
  return best?.manifest ?? null;
}

export function createAgentRuntime(options: {
  manifest: MockupAgentManifest;
  client: GatewayClient | null;
  now?: () => number;
  /** Durata minima visibile di un passo, per leggere «in corso…». */
  minStepMs?: number;
  /** Risposte dal vivo disponibili all'apertura (dal manifesto). */
  remaining?: number | null;
  onChange?: (run: RunState, workspace: Workspace) => void;
}) {
  const { manifest, client } = options;
  const now = options.now ?? (() => Date.now());
  const minStepMs = options.minStepMs ?? 0;
  const firstTable = manifest.sampleDocuments.find((d) => d.table)?.table ?? null;
  let workspace: Workspace = {
    variables: Object.fromEntries(manifest.variables.map((v) => [v.key, v.example])),
    citations: [],
    table: null,
    files: [],
    remaining: options.remaining ?? null,
  };
  let run: RunState = { pathId: null, steps: [], running: false };
  // Fatti (soglie) delle sole citazioni ricevute dal vivo dal gateway: mai dal copione del manifesto.
  const liveFacts = new Map<string, ThresholdFact[]>();
  const rememberLive = (list: Citation[]) => { for (const c of list) if (c.facts?.length) liveFacts.set(c.itemKey, c.facts); };
  const emit = () => options.onChange?.(run, workspace);
  const patch = (next: Partial<Workspace>) => { workspace = { ...workspace, ...next }; emit(); };
  const setStep = (id: string, next: Partial<StepView>) => {
    run = { ...run, steps: run.steps.map((s) => (s.id === id ? { ...s, ...next } : s)) };
    emit();
  };
  const addCitations = (list: Citation[]) => {
    const byKey = new Map(workspace.citations.map((c) => [c.itemKey, c]));
    // Una voce già citata si arricchisce (testo, fatti), non si duplica.
    for (const c of list) byKey.set(c.itemKey, { ...byKey.get(c.itemKey), ...c });
    patch({ citations: [...byKey.values()] });
  };
  const offer = (...files: Workspace["files"]) => patch({ files: [...new Set([...workspace.files, ...files])] });
  const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const tableFor = (input: Record<string, string>): AgentTable | null => {
    const sample = input.table ? manifest.sampleDocuments.find((d) => d.id === input.table)?.table : undefined;
    return sample ?? workspace.table ?? firstTable;
  };
  const useTable = (table: AgentTable, source: string) => { if (!workspace.table) patch({ table, tableSource: source }); };
  const fail = (f: GatewayFailure) => { const notice = failureNotice(f); if (notice) patch({ notice }); };
  /** Contatore delle risposte dal vivo: il valore del gateway se c'è, altrimenti una in meno. */
  const countLive = (remaining: unknown) => {
    if (typeof remaining === "number" && Number.isFinite(remaining)) patch({ remaining: Math.max(0, Math.floor(remaining)) });
    else if (typeof workspace.remaining === "number") patch({ remaining: Math.max(0, workspace.remaining - 1) });
  };
  /** Valore di una variabile del manifesto per tipo o per nome (ripiego se il passo non la nomina). */
  const variableLike = (type: string, name: RegExp) => {
    const v = manifest.variables.find((x) => x.type === type && name.test(x.key)) ?? manifest.variables.find((x) => name.test(x.key) || name.test(x.label.toLowerCase())) ?? (type === "importo" ? manifest.variables.find((x) => x.type === type) : undefined);
    return v ? workspace.variables[v.key] ?? "" : "";
  };

  /** Applica l'output (o il copione) di un tool del gateway; restituisce dettaglio e citazioni del passo. */
  function applyGateway(out: Normalized, citations: Citation[], live: boolean): { detail: string; cites: Citation[] } {
    if (citations.length) addCitations(citations);
    if (live) rememberLive([...citations, ...(out.tool === "catalog.search" ? out.hits : out.tool === "catalog.cite" ? [out.item] : [])]);
    switch (out.tool) {
      case "catalog.search":
        addCitations(out.hits);
        return { detail: out.hits[0]?.citation ?? "nessuna voce trovata nel catalogo", cites: out.hits.slice(0, 1) };
      case "catalog.cite":
        addCitations([out.item]);
        return { detail: out.item.citation, cites: [out.item] };
      case "document.read":
        return { detail: out.summary.trim() ? out.summary.trim().slice(0, 160) : `${out.sections.length} sezioni, ${out.fields.length} campi variabili trovati`, cites: [] };
      case "data.extract_table":
        patch({ table: { columns: out.columns, rows: out.rows }, tableSource: workspace.documentTitle ?? "Documento", comparison: undefined, kpis: undefined, chart: undefined });
        return { detail: `${out.rows.length} righe × ${out.columns.length} colonne`, cites: [] };
      case "draft.compose":
        patch({ draft: out.draft });
        return { detail: out.draft.sections.length === 1 ? "testo pronto" : `${out.draft.sections.length} sezioni scritte`, cites: [] };
    }
  }

  /** Esegue un tool locale; null = non applicabile ai dati presenti (si usa il copione). */
  function runLocal(step: MockupAgentStep, input: Record<string, string>): string | LocalOutcome | null {
    switch (step.tool) {
      case "rule.check_threshold": {
        const wanted = input.item_key ? itemKeyOf(input.item_key) : undefined;
        const pick = <T,>(entries: [string, T][]) => entries.find(([k]) => k === wanted) ?? entries[0];
        const live = pick([...liveFacts.entries()]);
        const prepared = live ? undefined : pick(workspace.citations.filter((c) => c.facts?.length).map((c) => [c.itemKey, c.facts as ThresholdFact[]] as [string, ThresholdFact[]]));
        const facts = live?.[1] ?? prepared?.[1] ?? [];
        const source = workspace.citations.find((c) => c.itemKey === (live ?? prepared)?.[0]);
        // Importo e categoria: dall'input del passo; se mancano, dalle variabili del manifesto.
        const amount = input.amount?.trim() ? input.amount : variableLike("importo", /import|valore/);
        const category = input.category?.trim() ? input.category : variableLike("scelta", /tipolog|categor|tipo/);
        const check = checkThreshold({ amount, category, factKey: input.fact_key, facts, citation: source?.citation });
        patch({ threshold: check.ok ? { ...check.result, ...(prepared ? { prepared: true } : {}) } : undefined });
        if (!check.ok) return { detail: check.message, error: true };
        return prepared ? { detail: `${check.result.text} (${preparedThresholdNote})`, prepared: true } : check.result.text;
      }
      case "data.compute": {
        const table = tableFor(input);
        if (!table) return null;
        const op = input.op ?? "";
        if (op === "rotation") {
          const result = checkRotation(table, {
            operator: input.operator ?? "",
            operatorColumn: input.operator_column ?? "Operatore",
            dateColumn: input.date_column ?? "Data",
            months: parseNumber(input.months) ?? 12,
            today: now(),
          });
          if (!result) return null;
          patch({ rotation: result });
          return result.text;
        }
        if (op === "month_over_month") {
          const result = monthOverMonth(groupByMonth(table, input.date_column ?? "Mese", input.column));
          return result ? result.text : null;
        }
        if (op === "normalize") {
          const result = normalizeTable(table);
          patch({ table: result.table, tableSource: workspace.tableSource ?? "Documento" });
          return result.changed ? `${result.changed} celle uniformate` : "dati già uniformi";
        }
        if (op === "compare") {
          const rows = compareRows(table, {
            column: input.column,
            compareColumn: input.compare_column,
            baseColumn: input.base_column,
            labelColumns: (input.label_columns ?? "").split(",").map((s) => s.trim()).filter(Boolean),
          });
          if (!rows.length) return null;
          const anomalies = findAnomalies(rows, parseNumber(input.delta_limit) ?? 0.1, parseNumber(input.capacity_limit) ?? 0.7);
          useTable(table, "Documento di esempio");
          patch({ comparison: { rows, anomalies, column: input.column, compareColumn: input.compare_column, baseColumn: input.base_column } });
          const total = input.compare_column ? evaluateFormula(table, { op: "delta_pct", column: input.column, compareColumn: input.compare_column }) : null;
          return `${total === null ? "" : formatSignedPct(total) + " sul totale, "}${anomalies.length} anomali${anomalies.length === 1 ? "a" : "e"} da guardare`;
        }
        const value = evaluateFormula(table, { op, column: input.column, compareColumn: input.compare_column, dateColumn: input.date_column, threshold: parseNumber(input.threshold) ?? undefined });
        if (value === null) return null;
        const unit = input.unit === "€" || input.unit === "%" || input.unit === "n" ? input.unit : ["ratio", "delta_pct", "threshold", "distinct_pct"].includes(op) ? "%" : "n";
        return `${input.label ?? input.column}: ${op === "delta_pct" ? formatSignedPct(value) : formatValue(value, unit)}`;
      }
      case "kpi.show": {
        const table = tableFor(input);
        if (!table) return null;
        const kpis: KpiView[] = manifest.kpis.map((k) => {
          const raw = kpiValue(k.formula, table);
          // Uno scostamento si legge con il segno: «+1,9%».
          const value = k.formula.op === "delta_pct" && raw !== null ? formatSignedPct(raw) : formatValue(raw, k.unit);
          return { id: k.id, label: k.label, raw, unit: k.unit, value };
        });
        if (input.distinct_column) {
          const share = evaluateFormula(table, { op: "distinct_pct", column: input.distinct_column });
          if (share !== null) kpis.push({ id: "rotazione", label: input.distinct_label ?? "Rotazione degli operatori", raw: share, unit: "%", value: `${formatPct(share)} operatori diversi` });
        }
        if (input.group_column && input.column && input.compare_column) {
          const worst = worstGroup(table, input.group_column, input.column, input.compare_column);
          if (worst) kpis.push({ id: "peggiore", label: input.worst_label ?? "Andamento peggiore", raw: worst.delta, unit: "testo", value: `${worst.group} ${formatSignedPct(worst.delta)}` });
        }
        if (!kpis.some((k) => k.raw !== null)) return null;
        useTable(table, "Documento di esempio");
        patch({ kpis });
        return `${kpis.filter((k) => k.raw !== null).length} indicatori calcolati sui dati`;
      }
      case "chart.show": {
        const table = tableFor(input);
        if (!table) return null;
        const points = groupByMonth(table, input.date_column ?? "Mese", input.column, input.compare_column);
        if (points.length < 2) return null;
        patch({ chart: { title: input.title ?? input.column, points, column: input.column, compareColumn: input.compare_column } });
        return `${points.length} mesi nel grafico`;
      }
      case "file.xlsx": {
        const table = tableFor(input);
        if (!table) return null;
        useTable(table, "Documento di esempio");
        if (input.sheet?.trim()) patch({ sheet: input.sheet.trim().slice(0, 31) });
        offer("xlsx");
        return "Excel con formule pronto da scaricare";
      }
      case "file.docx":
        offer("docx", "pdf");
        return "Word e PDF pronti";
      default:
        return null;
    }
  }

  async function runStep(pathId: string, step: MockupAgentStep) {
    const started = now();
    setStep(step.id, { status: "running", startedAt: started, model: isModelTool(step.tool) });
    const input = resolveInput(canonicalStepInput(step.tool, step.input), workspace.variables);
    const fallback = fillFallback(step.fallback, workspace.variables);
    let detail: string | null = null;
    let usedFallback = false;
    let failed = false;
    let cites: Citation[] = [];
    if (isGatewayTool(step.tool)) {
      const documentId = workspace.documentId;
      const sample: Record<string, string> = !documentId && workspace.sampleDocumentId && !input.sample_document_id ? { sample_document_id: workspace.sampleDocumentId } : {};
      const result = client
        ? await client.run({ pathId, stepId: step.id, input: { ...input, ...sample }, ...(documentId ? { documentId } : {}), requestKey: requestKey() })
        : null;
      let output = result?.ok ? validToolOutput(step.tool, result.data.output) : null;
      const live = !!output;
      const citations = result?.ok && output && Array.isArray(result.data.citations) ? result.data.citations.filter(isCitation).map(cleanCitation) : [];
      if (result && !result.ok) fail(result);
      if (result?.ok && live && result.data.mode === "model") countLive(result.data.remaining);
      if (!output) {
        output = validToolOutput(step.tool, fallback);
        usedFallback = true;
      }
      if (output) ({ detail, cites } = applyGateway(output, citations, live));
    } else {
      const outcome = runLocal(step, input);
      if (outcome !== null && typeof outcome === "object") {
        detail = outcome.detail;
        failed = !!outcome.error;
        usedFallback = !!outcome.prepared;
      } else detail = outcome;
      if (detail === null && step.tool !== "rule.check_threshold") {
        const fb = parseJson(fallback) as { detail?: unknown; text?: unknown } | null;
        detail = typeof fb?.detail === "string" ? fb.detail : typeof fb?.text === "string" ? fb.text : null;
        usedFallback = detail !== null;
      }
    }
    // Anche un passo istantaneo (locale o dal copione) resta «in corso…» quanto basta per leggerlo.
    const elapsed = now() - started;
    if (elapsed < minStepMs) await sleep(minStepMs - elapsed);
    setStep(step.id, {
      status: detail === null || failed ? "error" : "done",
      detail: detail ?? "passo non disponibile, proseguo",
      ms: Math.max(0, now() - started),
      prepared: usedFallback,
      ...(cites.length ? { citations: cites } : {}),
    });
  }

  return {
    get workspace() { return workspace; },
    get run() { return run; },
    setVariables(values: Record<string, string>) { patch({ variables: { ...workspace.variables, ...values } }); },
    useSample(id: string) {
      const doc = manifest.sampleDocuments.find((d) => d.id === id);
      if (!doc) return;
      patch({ sampleDocumentId: id, documentId: undefined, documentTitle: doc.title, ...(doc.table ? { table: doc.table, tableSource: doc.title, comparison: undefined, kpis: undefined, chart: undefined } : {}) });
    },
    /** Documento ammesso dal gateway (già filtrato per la privacy). */
    useDocument(documentId: string, title: string, sampleId?: string) {
      patch({ documentId, documentTitle: title, sampleDocumentId: sampleId });
    },
    async runPath(pathId: string) {
      const path = manifest.paths.find((p) => p.id === pathId);
      if (!path || run.running) return;
      run = { pathId, running: true, steps: path.steps.map((s) => ({ id: s.id, label: s.label, tool: s.tool, status: "pending" })) };
      emit();
      for (const step of path.steps) await runStep(pathId, step);
      run = { ...run, running: false };
      emit();
    },
    async ask(message: string): Promise<AskOutcome> {
      const text = message.trim().slice(0, mockupAgentPolicy.messageMaxLength);
      if (!text) return { kind: "hint", text: "Scrivi una domanda o scegli un percorso guidato." };
      const local = answerDataQuestion(text, workspace.table);
      if (local) return { kind: "answer", text: local, citations: [], local: true };
      const result = client ? await client.ask({ message: text, requestKey: requestKey() }) : null;
      if (result?.ok) {
        const data = result.data;
        const known = new Set(manifest.variables.map((m) => m.key));
        const useVariables = (values: unknown) => {
          const vars = Object.fromEntries(Object.entries(values && typeof values === "object" ? values : {}).filter(([k, v]) => typeof v === "string" && known.has(k)));
          if (Object.keys(vars).length) patch({ variables: { ...workspace.variables, ...vars } });
        };
        if (data.kind === "answer" && typeof data.answer === "string" && data.answer.trim()) {
          countLive(data.remaining);
          const suggested = typeof data.suggestedPathId === "string" && manifest.paths.some((p) => p.id === data.suggestedPathId) ? data.suggestedPathId : undefined;
          if (suggested) useVariables(data.variables);
          const citations = Array.isArray(data.citations) ? data.citations.filter(isCitation).map(cleanCitation) : [];
          return { kind: "answer", text: data.answer, citations, ...(suggested ? { suggest: { pathId: suggested } } : {}) };
        }
        if (data.kind === "path" && manifest.paths.some((p) => p.id === data.pathId)) {
          countLive(data.remaining);
          useVariables(data.variables);
          // Una domanda riceve una risposta, non un avvio: il percorso resta una proposta.
          if (isQuestion(text)) return { kind: "hint", text: questionPathHint(manifest.paths.find((p) => p.id === data.pathId)!.label), suggest: { pathId: data.pathId } };
          return { kind: "path", pathId: data.pathId };
        }
      } else if (result) {
        // 422: tentativo di aggiramento o richiesta non ammessa: si mostra il messaggio, niente copione.
        if (result.error === "rejected_abuse") return { kind: "hint", text: result.message ?? "Richiesta non ammessa.", error: true };
        fail(result);
      }
      // Copione: instradamento per parole chiave sulle frasi di esempio dei percorsi.
      const words = (s: string) => s.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").split(/[^a-z0-9]+/).filter((w) => w.length > 3);
      const asked = new Set(words(text));
      let best: { id: string; score: number } | null = null;
      for (const p of manifest.paths) {
        const score = new Set([...words(p.label), ...p.examples.flatMap(words)].filter((w) => asked.has(w))).size;
        if (score > (best?.score ?? 0)) best = { id: p.id, score };
      }
      if (best && isQuestion(text)) return { kind: "hint", text: preparedQuestionHint(manifest.paths.find((p) => p.id === best.id)!.label), suggest: { pathId: best.id, prepared: true } };
      if (best) return { kind: "path", pathId: best.id, prepared: true };
      return { kind: "hint", text: mockupAgentMessages.freeTextFallback };
    },
  };
}
export type AgentRuntime = ReturnType<typeof createAgentRuntime>;
