/**
 * Parse Excel feedback templates into structured form definitions.
 * Supports image URLs, local asset paths, and embedded Excel images.
 */

import { extractEmbeddedImages, resolveImagePath } from "./image-extractor.js";

const QUESTION_TYPES = new Set([
  "text", "email", "textarea", "select", "radio", "checkbox",
  "rating", "scale", "nps", "info",
]);

const COLUMN_ALIASES = {
  id: ["id", "question_id", "questionid", "key", "field_id", "fieldid", "name"],
  section: ["section", "group", "category", "page", "block"],
  type: ["type", "question_type", "questiontype", "input_type", "inputtype", "format"],
  question: ["question", "label", "title", "question text", "question_text", "prompt", "text"],
  options: ["options", "choices", "answers", "values", "option_list"],
  image: ["image", "image_url", "image url", "img", "picture", "photo", "illustration"],
  required: ["required", "mandatory", "is_required"],
  min: ["min", "minimum", "min_value"],
  max: ["max", "maximum", "max_value"],
  placeholder: ["placeholder", "hint", "help_text", "example"],
};

function normalizeKey(key) {
  return String(key || "").trim().toLowerCase().replace(/[_\s]+/g, " ");
}

function pickField(row, aliases) {
  const keys = Object.keys(row);
  const normalized = Object.fromEntries(keys.map((k) => [normalizeKey(k), k]));
  for (const alias of aliases) {
    const match = normalized[normalizeKey(alias)];
    if (match != null && row[match] !== "") return row[match];
  }
  return "";
}

function normalizeType(raw) {
  const t = String(raw || "text").trim().toLowerCase();
  const aliases = {
    "single choice": "radio",
    "multiple choice": "checkbox",
    "multi select": "checkbox",
    "multiselect": "checkbox",
    "dropdown": "select",
    "stars": "rating",
    "star": "rating",
    "likert": "scale",
    "number": "scale",
    "long text": "textarea",
    "short text": "text",
    "description": "info",
    "header": "info",
  };
  const mapped = aliases[t] || t;
  return QUESTION_TYPES.has(mapped) ? mapped : "text";
}

function parseBool(val, fallback = true) {
  if (typeof val === "boolean") return val;
  const s = String(val ?? "").trim().toLowerCase();
  if (!s) return fallback;
  return s === "true" || s === "yes" || s === "1" || s === "y";
}

function parseOptions(raw) {
  if (!raw) return [];
  const str = String(raw).trim();
  if (!str) return [];
  const sep = str.includes("|") ? "|" : str.includes(";") ? ";" : ",";
  return str.split(sep).map((o) => o.trim()).filter(Boolean);
}

function parseNumber(val, fallback = null) {
  if (val === "" || val == null) return fallback;
  const n = Number(val);
  return Number.isFinite(n) ? n : fallback;
}

function sheetToRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
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

function rowToQuestion(row, index, embeddedImage = null) {
  const id = String(pickField(row, COLUMN_ALIASES.id) || `q_${index + 1}`).trim();
  const section = String(pickField(row, COLUMN_ALIASES.section) || "General").trim();
  const type = normalizeType(pickField(row, COLUMN_ALIASES.type));
  const question = String(pickField(row, COLUMN_ALIASES.question)).trim();
  const optionsRaw = pickField(row, COLUMN_ALIASES.options);
  let image = String(pickField(row, COLUMN_ALIASES.image) || "").trim();
  if (image === "true" || image === "false") image = "";
  if (!image && embeddedImage) image = embeddedImage;

  const required = parseBool(pickField(row, COLUMN_ALIASES.required), type !== "info");
  const min = parseNumber(pickField(row, COLUMN_ALIASES.min), type === "nps" ? 0 : 1);
  const max = parseNumber(pickField(row, COLUMN_ALIASES.max), type === "nps" ? 10 : type === "scale" ? 10 : 5);
  const placeholder = String(pickField(row, COLUMN_ALIASES.placeholder) || "").trim();

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
    _rowIndex: index,
  };
}

export function parseWorkbook(workbook, embeddedImages = {}) {
  const sheetName =
    workbook.SheetNames.find((n) => /^feedback/i.test(n)) ||
    workbook.SheetNames.find((n) => /form|survey|questions/i.test(n)) ||
    workbook.SheetNames.find((n) => !/^meta$/i.test(n)) ||
    workbook.SheetNames[0];

  const sheet = workbook.Sheets[sheetName];
  const rawRows = sheetToRows(sheet);
  const questions = rawRows
    .map((row, i) => rowToQuestion(row, i, embeddedImages[i + 1] || embeddedImages[i] || null))
    .filter(Boolean);

  const sections = [...new Set(questions.map((q) => q.section))];
  const meta = parseMetaSheet(workbook);

  return {
    meta: {
      eventName: meta.event_name || meta.eventName || meta.event || "S2 Launch Event",
      eventDate: meta.event_date || meta.eventDate || "",
      version: meta.version || "1.0",
      description: meta.description || "",
      ...meta,
    },
    sections,
    questions: questions.map(({ _rowIndex, ...q }) => q),
    parsedAt: new Date().toISOString(),
    sourceSheet: sheetName,
  };
}

async function resolveQuestionImages(config) {
  for (const q of config.questions) {
    if (q.image) {
      q.image = await resolveImagePath(q.image);
    }
  }
  return config;
}

export async function parseExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const embeddedImages = await extractEmbeddedImages(buffer);
  const workbook = XLSX.read(buffer, { type: "array" });
  const config = parseWorkbook(workbook, embeddedImages);
  return resolveQuestionImages(config);
}

export function groupBySection(questions) {
  const map = new Map();
  for (const q of questions) {
    if (!map.has(q.section)) map.set(q.section, []);
    map.get(q.section).push(q);
  }
  return map;
}

export async function fetchTemplatePaths() {
  const paths = [];
  try {
    const resp = await fetch("templates/manifest.json");
    if (resp.ok) {
      const manifest = await resp.json();
      if (manifest.preferred) paths.push(manifest.preferred);
      if (Array.isArray(manifest.templates)) paths.push(...manifest.templates);
    }
  } catch {
    /* use defaults */
  }

  paths.push(
    "templates/feedback-form.xlsx",
    "templates/S2-Launch-Feedback.xlsx",
    "templates/s2-launch-feedback-template.xlsx",
  );

  return [...new Set(paths)];
}

export async function loadTemplateFromRepo() {
  const paths = await fetchTemplatePaths();
  for (const path of paths) {
    try {
      const resp = await fetch(path);
      if (!resp.ok) continue;
      const blob = await resp.blob();
      const file = new File([blob], path.split("/").pop());
      const config = await parseExcelFile(file);
      config.fileName = path.split("/").pop();
      config.sourcePath = path;
      return config;
    } catch (e) {
      console.warn(`Could not load template ${path}:`, e);
    }
  }
  return null;
}
