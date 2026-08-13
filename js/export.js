/**
 * Export responses and dashboard summaries to downloadable files.
 */

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function timestampSlug() {
  const d = new Date();
  return d.toISOString().slice(0, 19).replace(/[:T]/g, "-");
}

export function exportResponsesCSV(responses, formConfig) {
  if (!responses.length) return;

  const questionIds = formConfig.questions
    .filter((q) => q.type !== "info")
    .map((q) => q.id);

  const questionMap = Object.fromEntries(
    formConfig.questions.map((q) => [q.id, q.question])
  );

  const headers = [
    "response_id",
    "submitted_at",
    "participant_name",
    "participant_email",
    ...questionIds.map((id) => questionMap[id] || id),
  ];

  const rows = responses.map((r) => {
    const answers = r.answers || {};
    return [
      r.id,
      r.submittedAt || "",
      r.participantName || answers.participant_name || "",
      r.participantEmail || answers.participant_email || "",
      ...questionIds.map((id) => {
        const val = answers[id];
        if (Array.isArray(val)) return val.join("; ");
        return val != null ? String(val) : "";
      }),
    ];
  });

  const escape = (v) => {
    const s = String(v ?? "");
    if (s.includes(",") || s.includes('"') || s.includes("\n")) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  };

  const csv = [headers, ...rows].map((row) => row.map(escape).join(",")).join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  downloadBlob(blob, `s2-feedback-responses-${timestampSlug()}.csv`);
}

export function exportResponsesJSON(responses, formConfig) {
  const payload = {
    exportedAt: new Date().toISOString(),
    event: formConfig.meta?.eventName || "S2 Launch Event",
    totalResponses: responses.length,
    responses,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  downloadBlob(blob, `s2-feedback-responses-${timestampSlug()}.json`);
}

export function exportDashboardSummary(summary, formConfig) {
  const payload = {
    exportedAt: new Date().toISOString(),
    event: formConfig.meta?.eventName || "S2 Launch Event",
    ...summary,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  downloadBlob(blob, `s2-dashboard-summary-${timestampSlug()}.json`);
}

export function exportDashboardCSV(summary) {
  const lines = ["Metric,Value"];
  lines.push(`Total Responses,${summary.totalResponses}`);
  lines.push(`Average Overall Rating,${summary.avgOverallRating ?? "N/A"}`);
  lines.push(`Average NPS,${summary.avgNps ?? "N/A"}`);
  lines.push(`Promoters (9-10),${summary.npsPromoters ?? 0}`);
  lines.push(`Passives (7-8),${summary.npsPassives ?? 0}`);
  lines.push(`Detractors (0-6),${summary.npsDetractors ?? 0}`);
  lines.push(`NPS Score,${summary.npsScore ?? "N/A"}`);
  lines.push("");

  if (summary.ratingBreakdown) {
    lines.push("Rating Question,Average");
    for (const [q, avg] of Object.entries(summary.ratingBreakdown)) {
      lines.push(`"${q}",${avg}`);
    }
    lines.push("");
  }

  if (summary.optionCounts) {
    for (const [question, counts] of Object.entries(summary.optionCounts)) {
      lines.push(`"${question}"`);
      lines.push("Option,Count");
      for (const [opt, count] of Object.entries(counts)) {
        lines.push(`"${opt}",${count}`);
      }
      lines.push("");
    }
  }

  const blob = new Blob(["\uFEFF" + lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  downloadBlob(blob, `s2-dashboard-summary-${timestampSlug()}.csv`);
}
