/**
 * Parse Excel feedback templates into structured form definitions.
 * Supports standard column format and Vivo S2 survey-style layouts.
 */

import { extractEmbeddedImages, resolveImagePath } from "./image-extractor.js";

const QUESTION_TYPES = new Set([
  "text", "email", "textarea", "select", "radio", "checkbox",
  "rating", "scale", "nps", "info",
]);

const SURVEY_IMAGE_MAP = {
  "photo zone 1": "assets/images/photo-zone-1.jpg",
  "photo zone 2": "assets/images/photo-zone-2.jpg",
};

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

function slugify(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_|_$/g, "")
    .slice(0, 50) || "question";
}

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

function sheetToArray(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
}

function isSurveyFormat(sheet) {
  const rows = sheetToArray(sheet);
  return rows.some((row) => {
    const q = String(row[0] || "").trim();
    const r = String(row[1] || "").trim();
    return q === "Question" && /rating/i.test(r);
  });
}

function inferSurveyQuestionType(question, section) {
  const q = question.toLowerCase();
  if (q.includes("which zone did you prefer")) {
    return { type: "textarea", placeholder: "Share which zone you preferred and why…" };
  }
  if (q.includes("what worked best") || q.includes("what should be improved")) {
    return { type: "textarea", placeholder: "Your thoughts…" };
  }
  if (q.includes("would you recommend")) {
    return { type: "radio", options: ["Yes", "No"] };
  }
  if (q.includes("name") && q.includes("optional")) {
    return { type: "text", required: false, placeholder: "Your name" };
  }
  if (q.includes("department") && q.includes("optional")) {
    return { type: "text", required: false, placeholder: "Your department" };
  }
  return { type: "rating", min: 1, max: 5 };
}

function surveySectionImage(section) {
  const s = section.toLowerCase();
  for (const [key, path] of Object.entries(SURVEY_IMAGE_MAP)) {
    if ( s.includes(key)) return path;
  }
  return null;
}

function parseSurveySheet(sheet, embeddedImages = {}) {
  const rows = sheetToArray(sheet);
  const questions = [];
  let currentSection = "General";
  let pastHeader = false;
  const usedIds = new Set();

  const meta = {
    eventName: String(rows[0]?.[0] || "Vivo S2 Launch Event").trim(),
    description: [rows[1]?.[0], rows[2]?.[0]].filter(Boolean).join(" · "),
    venue: String(rows[1]?.[0] || "").replace(/^Venue:\s*/i, "").trim(),
  };

  if (rows[0]?.[0]) {
    questions.push({
      id: "info_welcome",
      section: "Welcome",
      type: "info",
      question: `${meta.eventName}${meta.venue ? ` — ${meta.venue}` : ""}. ${rows[2]?.[0] || "Please share your feedback below."}`,
      options: [],
      image: null,
      required: false,
      min: 1,
      max: 5,
      placeholder: "",
    });
  }

  const photoZoneImageUsed = { zone1: false, zone2: false };

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const col0 = String(row[0] || "").trim();
    if (!col0 || col0 === "#VALUE!") continue;

    if (col0 === "Question") {
      pastHeader = true;
      continue;
    }

    if (/^\d+\.\s/.test(col0)) {
      currentSection = col0;
      continue;
    }

    if (!pastHeader) {
      if (/name/i.test(col0) && /optional/i.test(col0)) {
        questions.push({
          id: "participant_name",
          section: "Your Details",
          type: "text",
          question: col0,
          options: [],
          image: null,
          required: false,
          min: 1,
          max: 5,
          placeholder: "Your name",
        });
      } else if (/department/i.test(col0) && /optional/i.test(col0)) {
        questions.push({
          id: "participant_department",
          section: "Your Details",
          type: "text",
          question: col0,
          options: [],
          image: null,
          required: false,
          min: 1,
          max: 5,
          placeholder: "Your department",
        });
      }
      continue;
    }

    const inferred = inferSurveyQuestionType(col0, currentSection);
    let image = null;

    const sectionLower = currentSection.toLowerCase();
    if (sectionLower.includes("photo zone 1") && col0.toLowerCase().includes("visual appeal") && !photoZoneImageUsed.zone1) {
      image = SURVEY_IMAGE_MAP["photo zone 1"];
      photoZoneImageUsed.zone1 = true;
    } else if (sectionLower.includes("photo zone 2") && col0.toLowerCase().includes("visual appeal") && !photoZoneImageUsed.zone2) {
      image = SURVEY_IMAGE_MAP["photo zone 2"];
      photoZoneImageUsed.zone2 = true;
    }

    if (!image) {
      const embedded = embeddedImages[i + 1] || embeddedImages[i];
      if (embedded) image = embedded;
    }

    let id = slugify(col0);
    if (usedIds.has(id)) id = `${id}_${i}`;
    usedIds.add(id);

    const section = col0.toLowerCase().includes("which zone")
      ? "Photo Zones"
      : col0.toLowerCase().includes("what worked") || col0.toLowerCase().includes("what should be improved")
        ? "Open Feedback"
        : currentSection;

    questions.push({
      id,
      section,
      type: inferred.type,
      question: col0,
      options: inferred.options || [],
      image,
      required: inferred.required ?? false,
      min: inferred.min ?? 1,
      max: inferred.max ?? 5,
      placeholder: inferred.placeholder || "",
    });
  }

  const sections = [...new Set(questions.map((q) => q.section))];

  return {
    meta: {
      eventName: meta.eventName,
      eventDate: "2026",
      version: "1.0",
      description: meta.description,
      venue: meta.venue,
    },
    sections,
    questions,
    parsedAt: new Date().toISOString(),
    sourceSheet: sheet.name || "Feedback Form",
    format: "survey",
  };
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

function parseColumnFormat(sheet, embeddedImages, workbook) {
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
    sourceSheet: sheet.name || "Feedback Form",
    format: "column",
  };
}

export function parseWorkbook(workbook, embeddedImages = {}) {
  const sheetName =
    workbook.SheetNames.find((n) => /^feedback/i.test(n)) ||
    workbook.SheetNames.find((n) => /form|survey|questions/i.test(n)) ||
    workbook.SheetNames.find((n) => !/^(meta|response|tally)$/i.test(n)) ||
    workbook.SheetNames[0];

  const sheet = workbook.Sheets[sheetName];
  sheet.name = sheetName;

  if (isSurveyFormat(sheet)) {
    return parseSurveySheet(sheet, embeddedImages);
  }

  return parseColumnFormat(sheet, embeddedImages, workbook);
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
    "S2_Launch_Event_Feedback_Form.xlsx",
    "templates/S2_Launch_Event_Feedback_Form.xlsx",
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
