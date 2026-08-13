/**
 * S2 Launch Feedback Tool — Main Application
 */

import { firebaseConfig, COLLECTIONS } from "./firebase-config.js";
import { parseExcelFile, groupBySection, loadTemplateFromRepo } from "./form-parser.js";
import { computeSummary, renderDashboard } from "./dashboard.js";
import {
  exportResponsesCSV,
  exportResponsesJSON,
  exportDashboardSummary,
  exportDashboardCSV,
} from "./export.js";

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  collection,
  addDoc,
  getDocs,
  query,
  orderBy,
  serverTimestamp,
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

// ─── State ───────────────────────────────────────────────────────────────────
let db = null;
let firebaseReady = false;
let formConfig = null;
let responses = [];
let currentSection = 0;
let sections = [];

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// ─── Firebase Init ───────────────────────────────────────────────────────────
function initFirebase() {
  if (firebaseConfig.apiKey === "YOUR_API_KEY") {
    showToast("Firebase not configured — using local storage fallback", "warning");
    return false;
  }
  try {
    const app = initializeApp(firebaseConfig);
    db = getFirestore(app);
    firebaseReady = true;
    return true;
  } catch (e) {
    console.error("Firebase init failed:", e);
    showToast("Firebase connection failed", "error");
    return false;
  }
}

// ─── Storage Layer ───────────────────────────────────────────────────────────
async function loadFormConfig() {
  if (firebaseReady) {
    const snap = await getDoc(doc(db, COLLECTIONS.ROOT, COLLECTIONS.CONFIG));
    if (snap.exists()) return snap.data();
  }
  const local = localStorage.getItem("s2_form_config");
  return local ? JSON.parse(local) : null;
}

async function saveFormConfig(config) {
  config.updatedAt = new Date().toISOString();
  if (firebaseReady) {
    await setDoc(doc(db, COLLECTIONS.ROOT, COLLECTIONS.CONFIG), config);
  }
  localStorage.setItem("s2_form_config", JSON.stringify(config));
  formConfig = config;
}

async function loadResponses() {
  if (firebaseReady) {
    const q = query(
      collection(db, COLLECTIONS.ROOT, COLLECTIONS.CONFIG, COLLECTIONS.RESPONSES),
      orderBy("submittedAt", "desc")
    );
    const snap = await getDocs(q);
    return snap.docs.map((d) => ({ id: d.id, ...d.data() }));
  }
  const local = localStorage.getItem("s2_responses");
  return local ? JSON.parse(local) : [];
}

async function saveResponse(response) {
  if (firebaseReady) {
    const ref = await addDoc(
      collection(db, COLLECTIONS.ROOT, COLLECTIONS.CONFIG, COLLECTIONS.RESPONSES),
      { ...response, createdAt: serverTimestamp() }
    );
    return ref.id;
  }
  const all = await loadResponses();
  const id = "local_" + Date.now();
  all.unshift({ id, ...response });
  localStorage.setItem("s2_responses", JSON.stringify(all));
  return id;
}

// ─── Navigation ──────────────────────────────────────────────────────────────
function showView(viewId) {
  $$(".view").forEach((v) => v.classList.remove("active"));
  $(`#view-${viewId}`)?.classList.add("active");
  $$(".nav-link").forEach((l) => l.classList.toggle("active", l.dataset.view === viewId));
  if (viewId === "dashboard") refreshDashboard();
}

// ─── Form Rendering ──────────────────────────────────────────────────────────
function renderForm() {
  if (!formConfig?.questions?.length) {
    $("#form-sections").innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📋</div>
        <h3>No feedback form loaded</h3>
        <p>Upload an Excel template from the Admin panel to get started.</p>
        <button class="btn btn-primary" onclick="document.querySelector('[data-view=admin]').click()">Go to Admin</button>
      </div>`;
    return;
  }

  const grouped = groupBySection(formConfig.questions);
  sections = [...grouped.keys()];
  currentSection = 0;

  const eventName = formConfig.meta?.eventName || "S2 Launch Event";
  $("#form-event-title").textContent = eventName;
  $("#form-event-desc").textContent = formConfig.meta?.description || "Share your experience with us";

  renderSection(grouped);
  updateProgress();
}

function renderSection(grouped) {
  const container = $("#form-sections");
  container.innerHTML = "";

  const sectionNames = [...grouped.keys()];
  const sectionName = sectionNames[currentSection];
  const questions = grouped.get(sectionName);

  const sectionEl = document.createElement("div");
  sectionEl.className = "form-section active";
  sectionEl.innerHTML = `<h2 class="section-title">${escapeHtml(sectionName)}</h2>`;

  for (const q of questions) {
    sectionEl.appendChild(createQuestionEl(q));
  }

  container.appendChild(sectionEl);

  const nav = document.createElement("div");
  nav.className = "form-nav";
  nav.innerHTML = `
    <button type="button" class="btn btn-ghost" id="btn-prev" ${currentSection === 0 ? "disabled" : ""}>
      ← Previous
    </button>
    <span class="section-indicator">${currentSection + 1} / ${sectionNames.length}</span>
    ${
      currentSection < sectionNames.length - 1
        ? `<button type="button" class="btn btn-primary" id="btn-next">Next →</button>`
        : `<button type="button" class="btn btn-primary btn-glow" id="btn-submit">Submit Feedback</button>`
    }
  `;
  container.appendChild(nav);

  $("#btn-prev")?.addEventListener("click", () => {
    if (currentSection > 0) {
      currentSection--;
      renderSection(grouped);
      updateProgress();
      container.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  });

  $("#btn-next")?.addEventListener("click", () => {
    if (!validateCurrentSection(questions)) return;
    currentSection++;
    renderSection(grouped);
    updateProgress();
    container.scrollIntoView({ behavior: "smooth", block: "start" });
  });

  $("#btn-submit")?.addEventListener("click", () => submitForm(questions, grouped));
}

function createQuestionEl(q) {
  const wrap = document.createElement("div");
  wrap.className = `question-block type-${q.type}`;
  wrap.dataset.qid = q.id;

  if (q.type === "info") {
    wrap.innerHTML = `<div class="info-block"><p>${escapeHtml(q.question)}</p></div>`;
    return wrap;
  }

  let imageHtml = "";
  if (q.image) {
    imageHtml = `
      <div class="question-image">
        <img src="${escapeAttr(q.image)}" alt="Question illustration" loading="lazy"
             onerror="this.parentElement.classList.add('image-error')" />
      </div>`;
  }

  let inputHtml = "";
  const req = q.required ? "required" : "";
  const name = q.id;

  switch (q.type) {
    case "text":
    case "email":
      inputHtml = `<input type="${q.type}" name="${name}" id="${name}" placeholder="${escapeAttr(q.placeholder)}" ${req} class="input-field" />`;
      break;
    case "textarea":
      inputHtml = `<textarea name="${name}" id="${name}" placeholder="${escapeAttr(q.placeholder)}" ${req} class="input-field textarea-field" rows="4"></textarea>`;
      break;
    case "select":
      inputHtml = `<select name="${name}" id="${name}" ${req} class="input-field select-field">
        <option value="">Select an option</option>
        ${q.options.map((o) => `<option value="${escapeAttr(o)}">${escapeHtml(o)}</option>`).join("")}
      </select>`;
      break;
    case "radio":
      inputHtml = `<div class="option-group radio-group">
        ${q.options.map((o, i) => `
          <label class="option-card">
            <input type="radio" name="${name}" value="${escapeAttr(o)}" ${req && i === 0 ? "" : ""} />
            <span class="option-label">${escapeHtml(o)}</span>
          </label>`).join("")}
      </div>`;
      break;
    case "checkbox":
      inputHtml = `<div class="option-group checkbox-group">
        ${q.options.map((o) => `
          <label class="option-card">
            <input type="checkbox" name="${name}" value="${escapeAttr(o)}" />
            <span class="option-label">${escapeHtml(o)}</span>
          </label>`).join("")}
      </div>`;
      break;
    case "rating":
      inputHtml = renderStarRating(name, q.min, q.max, req);
      break;
    case "scale":
      inputHtml = renderScale(name, q.min, q.max, req);
      break;
    case "nps":
      inputHtml = renderNps(name, q.min, q.max);
      break;
    default:
      inputHtml = `<input type="text" name="${name}" id="${name}" class="input-field" ${req} />`;
  }

  wrap.innerHTML = `
    ${imageHtml}
    <label class="question-label" for="${name}">
      ${escapeHtml(q.question)}
      ${q.required ? '<span class="required-mark">*</span>' : ""}
    </label>
    ${inputHtml}
    <div class="field-error" id="error-${name}"></div>
  `;

  if (q.type === "rating") bindStarRating(wrap, name);
  if (q.type === "scale") bindScale(wrap, name);
  if (q.type === "nps") bindNps(wrap, name);

  return wrap;
}

function renderStarRating(name, min, max, req) {
  const count = max || 5;
  return `<div class="star-rating" data-name="${name}" data-required="${req}">
    ${Array.from({ length: count }, (_, i) => `
      <button type="button" class="star-btn" data-value="${i + 1}" aria-label="${i + 1} stars">
        <svg viewBox="0 0 24 24"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>
      </button>`).join("")}
    <input type="hidden" name="${name}" id="${name}" ${req} value="" />
    <span class="rating-label"></span>
  </div>`;
}

function bindStarRating(wrap, name) {
  const container = wrap.querySelector(".star-rating");
  const hidden = wrap.querySelector(`#${name}`);
  const label = wrap.querySelector(".rating-label");
  const btns = container.querySelectorAll(".star-btn");

  btns.forEach((btn) => {
    btn.addEventListener("click", () => {
      const val = Number(btn.dataset.value);
      hidden.value = val;
      btns.forEach((b, i) => b.classList.toggle("active", i < val));
      label.textContent = `${val} / ${btns.length}`;
    });
  });
}

function renderScale(name, min, max) {
  const lo = min ?? 1;
  const hi = max ?? 10;
  return `<div class="scale-rating" data-name="${name}">
    <div class="scale-labels"><span>${lo}</span><span>${hi}</span></div>
    <div class="scale-buttons">
      ${Array.from({ length: hi - lo + 1 }, (_, i) => {
        const v = lo + i;
        return `<button type="button" class="scale-btn" data-value="${v}">${v}</button>`;
      }).join("")}
    </div>
    <input type="hidden" name="${name}" id="${name}" value="" />
  </div>`;
}

function bindScale(wrap, name) {
  const btns = wrap.querySelectorAll(".scale-btn");
  const hidden = wrap.querySelector(`#${name}`);
  btns.forEach((btn) => {
    btn.addEventListener("click", () => {
      hidden.value = btn.dataset.value;
      btns.forEach((b) => b.classList.toggle("active", b === btn));
    });
  });
}

function renderNps(name, min, max) {
  const lo = min ?? 0;
  const hi = max ?? 10;
  return `<div class="nps-rating" data-name="${name}">
    <div class="nps-labels"><span>Not likely</span><span>Very likely</span></div>
    <div class="nps-buttons">
      ${Array.from({ length: hi - lo + 1 }, (_, i) => {
        const v = lo + i;
        const cls = v <= 6 ? "detractor" : v <= 8 ? "passive" : "promoter";
        return `<button type="button" class="nps-btn ${cls}" data-value="${v}">${v}</button>`;
      }).join("")}
    </div>
    <input type="hidden" name="${name}" id="${name}" value="" required />
  </div>`;
}

function bindNps(wrap, name) {
  const btns = wrap.querySelectorAll(".nps-btn");
  const hidden = wrap.querySelector(`#${name}`);
  btns.forEach((btn) => {
    btn.addEventListener("click", () => {
      hidden.value = btn.dataset.value;
      btns.forEach((b) => b.classList.toggle("selected", b === btn));
    });
  });
}

function updateProgress() {
  const pct = sections.length ? ((currentSection + 1) / sections.length) * 100 : 0;
  $("#progress-fill").style.width = `${pct}%`;
  $("#progress-text").textContent = `Section ${currentSection + 1} of ${sections.length}`;
}

function validateCurrentSection(questions) {
  let valid = true;
  for (const q of questions) {
    if (!q.required || q.type === "info") continue;
    const errEl = $(`#error-${q.id}`);
    if (errEl) errEl.textContent = "";

    const val = getAnswerValue(q);
    if (val === "" || val == null || (Array.isArray(val) && !val.length)) {
      if (errEl) errEl.textContent = "This field is required";
      valid = false;
    }
  }
  if (!valid) showToast("Please complete all required fields", "error");
  return valid;
}

function getAnswerValue(q) {
  if (q.type === "checkbox") {
    return [...document.querySelectorAll(`input[name="${q.id}"]:checked`)].map((el) => el.value);
  }
  if (q.type === "radio") {
    const checked = document.querySelector(`input[name="${q.id}"]:checked`);
    return checked ? checked.value : "";
  }
  const el = document.getElementById(q.id);
  return el ? el.value : "";
}

function collectAllAnswers() {
  const answers = {};
  for (const q of formConfig.questions) {
    if (q.type === "info") continue;
    answers[q.id] = getAnswerValue(q);
  }
  return answers;
}

async function submitForm(currentQuestions, grouped) {
  if (!validateCurrentSection(currentQuestions)) return;

  const allQuestions = formConfig.questions.filter((q) => q.type !== "info");
  for (const q of allQuestions) {
    if (!q.required) continue;
    const val = getAnswerValue(q);
    if (val === "" || val == null || (Array.isArray(val) && !val.length)) {
      showToast(`Please complete: ${q.question}`, "error");
      return;
    }
  }

  const answers = collectAllAnswers();
  const btn = $("#btn-submit");
  btn.disabled = true;
  btn.textContent = "Submitting…";

  const response = {
    answers,
    participantName: answers.participant_name || "",
    participantEmail: answers.participant_email || "",
    submittedAt: new Date().toISOString(),
    eventName: formConfig.meta?.eventName || "S2 Launch Event",
    userAgent: navigator.userAgent.slice(0, 200),
  };

  try {
    await saveResponse(response);
    showView("success");
    showToast("Thank you! Your feedback has been submitted.", "success");
  } catch (e) {
    console.error(e);
    showToast("Failed to submit. Please try again.", "error");
    btn.disabled = false;
    btn.textContent = "Submit Feedback";
  }
}

// ─── Admin / Template Upload ─────────────────────────────────────────────────
async function handleTemplateUpload(file) {
  if (!file) return;
  const ext = file.name.split(".").pop().toLowerCase();
  if (!["xlsx", "xls", "csv"].includes(ext)) {
    showToast("Please upload an .xlsx, .xls, or .csv file", "error");
    return;
  }

  try {
    showToast("Parsing template…", "info");
    const config = await parseExcelFile(file);
    if (!config.questions.length) throw new Error("No questions found in template");

    config.fileName = file.name;
    config.uploadedAt = new Date().toISOString();
    await saveFormConfig(config);

    $("#admin-template-info").innerHTML = `
      <div class="template-info-card">
        <div class="ti-row"><span>Event</span><strong>${escapeHtml(config.meta.eventName)}</strong></div>
        <div class="ti-row"><span>Questions</span><strong>${config.questions.length}</strong></div>
        <div class="ti-row"><span>Sections</span><strong>${config.sections.length}</strong></div>
        <div class="ti-row"><span>With Images</span><strong>${config.questions.filter((q) => q.image).length}</strong></div>
        <div class="ti-row"><span>File</span><strong>${escapeHtml(file.name)}</strong></div>
        <div class="ti-row"><span>Updated</span><strong>${new Date().toLocaleString()}</strong></div>
      </div>`;

    renderForm();
    showToast(`Template loaded: ${config.questions.length} questions`, "success");
  } catch (e) {
    console.error(e);
    showToast("Failed to parse template: " + e.message, "error");
  }
}

// ─── Repo Asset Check ────────────────────────────────────────────────────────
const REPO_PATHS_TO_CHECK = [
  "templates/feedback-form.xlsx",
  "templates/S2-Launch-Feedback.xlsx",
  "feedback-form.xlsx",
  "templates/s2-launch-feedback-template.xlsx",
  "assets/images/manifest.json",
];

const IMAGE_PATHS_TO_CHECK = async () => {
  const paths = [];
  try {
    const resp = await fetch("assets/images/manifest.json");
    if (resp.ok) {
      const data = await resp.json();
      if (Array.isArray(data.images)) paths.push(...data.images);
    }
  } catch { /* ignore */ }
  return paths;
};

async function checkRepoPath(path) {
  try {
    const resp = await fetch(path, { method: "HEAD" });
    return resp.ok;
  } catch {
    return false;
  }
}

async function checkRepoAssets() {
  const results = [];
  for (const path of REPO_PATHS_TO_CHECK) {
    const found = await checkRepoPath(path);
    results.push({ path, found, type: path.endsWith(".xlsx") ? "template" : "config" });
  }

  const imagePaths = await IMAGE_PATHS_TO_CHECK();
  for (const path of imagePaths) {
    results.push({ path, found: await checkRepoPath(path), type: "image" });
  }

  return results;
}

function renderRepoStatus(results) {
  const el = $("#repo-status");
  if (!el) return;

  const templates = results.filter((r) => r.type === "template");
  const images = results.filter((r) => r.type === "image");
  const foundTemplates = templates.filter((r) => r.found);
  const foundImages = images.filter((r) => r.found);

  el.innerHTML = `
    <div class="repo-status-grid">
      <div class="repo-status-section">
        <h4>Excel Templates</h4>
        ${templates.map((r) => `
          <div class="repo-status-row ${r.found ? "found" : "missing"}">
            <span class="status-dot"></span>
            <code>${escapeHtml(r.path)}</code>
            <span>${r.found ? "Found" : "Not found"}</span>
          </div>`).join("")}
      </div>
      <div class="repo-status-section">
        <h4>Images (${foundImages.length} found)</h4>
        ${images.length
          ? images.map((r) => `
            <div class="repo-status-row ${r.found ? "found" : "missing"}">
              <span class="status-dot"></span>
              <code>${escapeHtml(r.path)}</code>
              <span>${r.found ? "Found" : "Not found"}</span>
            </div>`).join("")
          : `<p class="repo-hint">Add photos to <code>assets/images/</code> and run <code>python3 scripts/build-image-manifest.py</code></p>`}
      </div>
    </div>
    <p class="repo-summary ${foundTemplates.some((t) => t.path.includes("feedback-form")) ? "ok" : "warn"}">
      ${foundTemplates.length
        ? `Using: <strong>${escapeHtml(foundTemplates[0].path)}</strong>`
        : "Your custom template (<code>templates/feedback-form.xlsx</code>) was not found in the repo yet."}
    </p>`;
}

async function reloadFromRepo() {
  showToast("Checking repo for templates…", "info");
  const results = await checkRepoAssets();
  renderRepoStatus(results);

  const config = await loadTemplateFromRepo();
  if (!config) {
    showToast("No template found in repo. Upload to templates/feedback-form.xlsx", "warning");
    return;
  }

  config.reloadedAt = new Date().toISOString();
  await saveFormConfig(config);

  $("#admin-template-info").innerHTML = `
    <div class="template-info-card">
      <div class="ti-row"><span>Event</span><strong>${escapeHtml(config.meta?.eventName || "S2 Launch")}</strong></div>
      <div class="ti-row"><span>Questions</span><strong>${config.questions.length}</strong></div>
      <div class="ti-row"><span>Sections</span><strong>${config.sections.length}</strong></div>
      <div class="ti-row"><span>With Images</span><strong>${config.questions.filter((q) => q.image).length}</strong></div>
      <div class="ti-row"><span>File</span><strong>${escapeHtml(config.fileName || config.sourcePath || "Repo template")}</strong></div>
      <div class="ti-row"><span>Reloaded</span><strong>${new Date().toLocaleString()}</strong></div>
    </div>`;

  formConfig = config;
  renderForm();
  showToast(`Loaded ${config.questions.length} questions from repo`, "success");
}

async function refreshDashboard() {
  responses = await loadResponses();
  const summary = computeSummary(responses, formConfig);
  renderDashboard(summary, $("#dashboard-content"));

  $("#dash-total").textContent = summary.totalResponses;

  window._lastSummary = summary;
}

// ─── Utilities ───────────────────────────────────────────────────────────────
function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s ?? "";
  return d.innerHTML;
}

function escapeAttr(s) {
  return String(s ?? "").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

function showToast(msg, type = "info") {
  const toast = document.createElement("div");
  toast.className = `toast toast-${type}`;
  toast.textContent = msg;
  $("#toast-container").appendChild(toast);
  requestAnimationFrame(() => toast.classList.add("show"));
  setTimeout(() => {
    toast.classList.remove("show");
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

// ─── Boot ────────────────────────────────────────────────────────────────────
async function boot() {
  initFirebase();

  // Navigation
  $$(".nav-link").forEach((link) => {
    link.addEventListener("click", (e) => {
      e.preventDefault();
      showView(link.dataset.view);
    });
  });

  // Template upload
  const uploadInput = $("#template-upload");
  const uploadZone = $("#upload-zone");

  uploadZone?.addEventListener("click", () => uploadInput.click());
  uploadZone?.addEventListener("dragover", (e) => {
    e.preventDefault();
    uploadZone.classList.add("drag-over");
  });
  uploadZone?.addEventListener("dragleave", () => uploadZone.classList.remove("drag-over"));
  uploadZone?.addEventListener("drop", (e) => {
    e.preventDefault();
    uploadZone.classList.remove("drag-over");
    const file = e.dataTransfer.files[0];
    if (file) handleTemplateUpload(file);
  });
  uploadInput?.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) handleTemplateUpload(file);
  });

  // Download template
  $("#btn-download-template")?.addEventListener("click", () => {
    const a = document.createElement("a");
    a.href = "templates/s2-launch-feedback-template.xlsx";
    a.download = "s2-launch-feedback-template.xlsx";
    a.click();
  });

  // Export buttons
  $("#btn-export-csv")?.addEventListener("click", async () => {
    responses = await loadResponses();
    if (!responses.length) return showToast("No responses to export", "warning");
    exportResponsesCSV(responses, formConfig);
  });

  $("#btn-export-json")?.addEventListener("click", async () => {
    responses = await loadResponses();
    if (!responses.length) return showToast("No responses to export", "warning");
    exportResponsesJSON(responses, formConfig);
  });

  $("#btn-export-summary-json")?.addEventListener("click", async () => {
    responses = await loadResponses();
    if (!responses.length) return showToast("No data to export", "warning");
    const summary = computeSummary(responses, formConfig);
    exportDashboardSummary(summary, formConfig);
  });

  $("#btn-export-summary-csv")?.addEventListener("click", async () => {
    responses = await loadResponses();
    if (!responses.length) return showToast("No data to export", "warning");
    const summary = computeSummary(responses, formConfig);
    exportDashboardCSV(summary);
  });

  $("#btn-refresh-dashboard")?.addEventListener("click", refreshDashboard);

  $("#btn-reload-repo")?.addEventListener("click", reloadFromRepo);

  $("#btn-new-response")?.addEventListener("click", () => {
    currentSection = 0;
    renderForm();
    showView("form");
  });

  // Load config
  formConfig = await loadFormConfig();

  if (!formConfig) {
    try {
      formConfig = await loadTemplateFromRepo();
      if (formConfig) await saveFormConfig(formConfig);
    } catch (e) {
      console.warn("Could not auto-load template from repo:", e);
    }
  }

  if (formConfig) {
    $("#admin-template-info").innerHTML = `
      <div class="template-info-card">
        <div class="ti-row"><span>Event</span><strong>${escapeHtml(formConfig.meta?.eventName || "S2 Launch")}</strong></div>
        <div class="ti-row"><span>Questions</span><strong>${formConfig.questions?.length || 0}</strong></div>
        <div class="ti-row"><span>Sections</span><strong>${formConfig.sections?.length || 0}</strong></div>
        <div class="ti-row"><span>With Images</span><strong>${formConfig.questions?.filter((q) => q.image).length || 0}</strong></div>
        <div class="ti-row"><span>File</span><strong>${escapeHtml(formConfig.fileName || "Default template")}</strong></div>
      </div>`;
  }

  renderForm();
  responses = await loadResponses();

  const repoResults = await checkRepoAssets();
  renderRepoStatus(repoResults);
}

document.addEventListener("DOMContentLoaded", boot);
