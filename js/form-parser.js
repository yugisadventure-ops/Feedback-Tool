/**
 * Parse Excel feedback templates into structured form definitions.
 * Supports images via URL in the "image" column.
 */

const QUESTION_TYPES = new Set([
  "text", "email", "textarea", "select", "radio", "checkbox",
  "rating", "scale", "nps", "info",
]);

function normalizeType(raw) {
  const t = String(raw || "text").trim().toLowerCase();
  return QUESTION_TYPES.has(t) ? t : "text";
}

function parseBool(val) {
  if (typeof val === "boolean") return val;
  const s = String(val || "").trim().toLowerCase();
  return s === "true" || s === "yes" || s === "1";
}

function parseOptions(raw) {
  if (!raw) return [];
  const str = String(raw).trim();
  if (!str) return [];
  return str.split("|").map((o) => o.trim()).filter(Boolean);
}

function parseNumber(val, fallback = null) {
  if (val === "" || val == null) return fallback;
  const n = Number(val);
  return Number.isFinite(n) ? n : fallback;
}

function sheetToRows(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows;
}

function parseMetaSheet(workbook) {
  const metaSheet = workbook.Sheets["Meta"] || workbook.Sheets["meta"];
  if (!metaSheet) return {};
  const rows = XLSX.utils.sheet_to_json(metaSheet, { header: 1, defval: "" });
  const meta = {};
  for (const row of rows) {
    if (row[0]) meta[String(row[0]).trim()] = row[1] != null ? String(row[1]).trim() : "";
  }
  return meta;
}

function detectImageColumn(row) {
  const keys = Object.keys(row);
  const imageKey = keys.find((k) => /^image$/i.test(k) || /^image_url$/i.test(k) || /^image url$/i.test(k));
  return imageKey || "image";
}

function rowToQuestion(row, index) {
  const imageKey = detectImageColumn(row);
  const id = String(row.id || row.ID || row.Id || `q_${index + 1}`).trim();
  const section = String(row.section || row.Section || "General").trim();
  const type = normalizeType(row.type || row.Type);
  const question = String(row.question || row.Question || "").trim();
  const optionsRaw = row.options ?? row.Options ?? "";
  let image = String(row[imageKey] || row.image || row.Image || "").trim();
  if (image === "true" || image === "false") image = "";
  const required = parseBool(row.required ?? row.Required ?? true);
  const min = parseNumber(row.min ?? row.Min, type === "nps" ? 0 : 1);
  const max = parseNumber(row.max ?? row.Max, type === "nps" ? 10 : type === "scale" ? 10 : 5);
  const placeholder = String(row.placeholder || row.Placeholder || "").trim();

  if (!question && type !== "info") return null;

  return {
    id,
    section,
    type,
    question,
    options: parseOptions(optionsRaw),
    image: image || null,
    required,
    min,
    max,
    placeholder,
  };
}

export function parseWorkbook(workbook) {
  const sheetName =
    workbook.SheetNames.find((n) => /^feedback/i.test(n)) ||
    workbook.SheetNames.find((n) => !/^meta$/i.test(n)) ||
    workbook.SheetNames[0];

  const sheet = workbook.Sheets[sheetName];
  const rawRows = sheetToRows(sheet);
  const questions = rawRows
    .map((row, i) => rowToQuestion(row, i))
    .filter(Boolean);

  const sections = [...new Set(questions.map((q) => q.section))];
  const meta = parseMetaSheet(workbook);

  return {
    meta: {
      eventName: meta.event_name || meta.eventName || "S2 Launch Event",
      eventDate: meta.event_date || meta.eventDate || "",
      version: meta.version || "1.0",
      description: meta.description || "",
      ...meta,
    },
    sections,
    questions,
    parsedAt: new Date().toISOString(),
    sourceSheet: sheetName,
  };
}

export async function parseExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  return parseWorkbook(workbook);
}

export function groupBySection(questions) {
  const map = new Map();
  for (const q of questions) {
    if (!map.has(q.section)) map.set(q.section, []);
    map.get(q.section).push(q);
  }
  return map;
}
