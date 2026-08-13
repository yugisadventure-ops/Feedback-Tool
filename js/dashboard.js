/**
 * Dashboard analytics and Chart.js rendering.
 */

let chartInstances = [];

function destroyCharts() {
  chartInstances.forEach((c) => c.destroy());
  chartInstances = [];
}

function avg(nums) {
  const valid = nums.filter((n) => Number.isFinite(n));
  if (!valid.length) return null;
  return Math.round((valid.reduce((a, b) => a + b, 0) / valid.length) * 10) / 10;
}

function countOptions(responses, questionId, options) {
  const counts = {};
  for (const opt of options) counts[opt] = 0;
  counts["Other"] = 0;

  for (const r of responses) {
    const val = r.answers?.[questionId];
    if (Array.isArray(val)) {
      for (const v of val) {
        counts[v] = (counts[v] ?? 0) + 1;
      }
    } else if (val != null && val !== "") {
      counts[val] = (counts[val] ?? 0) + 1;
    }
  }
  delete counts["Other"];
  if (Object.values(counts).every((v) => v === 0)) return null;
  return counts;
}

export function computeSummary(responses, formConfig) {
  const questions = formConfig?.questions || [];
  const ratingBreakdown = {};
  const optionCounts = {};

  for (const q of questions) {
    if (q.type === "rating" || q.type === "scale") {
      const vals = responses
        .map((r) => Number(r.answers?.[q.id]))
        .filter((n) => Number.isFinite(n));
      if (vals.length) ratingBreakdown[q.question] = avg(vals);
    }
    if (q.type === "radio" || q.type === "select") {
      const counts = countOptions(responses, q.id, q.options);
      if (counts) optionCounts[q.question] = counts;
    }
    if (q.type === "checkbox") {
      const counts = countOptions(responses, q.id, q.options);
      if (counts) optionCounts[q.question] = counts;
    }
  }

  const npsVals = responses
    .map((r) => Number(r.answers?.nps_score))
    .filter((n) => Number.isFinite(n));

  let npsPromoters = 0, npsPassives = 0, npsDetractors = 0;
  for (const n of npsVals) {
    if (n >= 9) npsPromoters++;
    else if (n >= 7) npsPassives++;
    else npsDetractors++;
  }

  const npsTotal = npsVals.length;
  const npsScore = npsTotal
    ? Math.round(((npsPromoters - npsDetractors) / npsTotal) * 100)
    : null;

  const overallVals = responses
    .map((r) => Number(r.answers?.overall_rating))
    .filter((n) => Number.isFinite(n));

  const recent = [...responses]
    .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt))
    .slice(0, 10);

  return {
    totalResponses: responses.length,
    avgOverallRating: avg(overallVals),
    avgNps: avg(npsVals),
    npsPromoters,
    npsPassives,
    npsDetractors,
    npsScore,
    ratingBreakdown,
    optionCounts,
    responsesOverTime: groupByDay(responses),
    recentResponses: recent,
  };
}

function groupByDay(responses) {
  const map = {};
  for (const r of responses) {
    const day = (r.submittedAt || "").slice(0, 10);
    if (day) map[day] = (map[day] || 0) + 1;
  }
  return Object.entries(map).sort(([a], [b]) => a.localeCompare(b));
}

const CHART_COLORS = [
  "#6366f1", "#8b5cf6", "#a78bfa", "#c4b5fd",
  "#22d3ee", "#34d399", "#fbbf24", "#f472b6",
];

function makeChart(ctx, config) {
  const chart = new Chart(ctx, config);
  chartInstances.push(chart);
  return chart;
}

export function renderDashboard(summary, container) {
  destroyCharts();
  container.innerHTML = "";

  const kpiRow = document.createElement("div");
  kpiRow.className = "kpi-grid";
  kpiRow.innerHTML = `
    <div class="kpi-card" data-animate>
      <span class="kpi-label">Total Responses</span>
      <span class="kpi-value">${summary.totalResponses}</span>
    </div>
    <div class="kpi-card" data-animate>
      <span class="kpi-label">Avg. Overall Rating</span>
      <span class="kpi-value">${summary.avgOverallRating ?? "—"}<small>/5</small></span>
    </div>
    <div class="kpi-card" data-animate>
      <span class="kpi-label">NPS Score</span>
      <span class="kpi-value ${summary.npsScore >= 50 ? "positive" : summary.npsScore >= 0 ? "neutral" : "negative"}">${summary.npsScore ?? "—"}</span>
    </div>
    <div class="kpi-card" data-animate>
      <span class="kpi-label">Avg. NPS Rating</span>
      <span class="kpi-value">${summary.avgNps ?? "—"}<small>/10</small></span>
    </div>
  `;
  container.appendChild(kpiRow);

  const chartsRow = document.createElement("div");
  chartsRow.className = "charts-grid";

  if (summary.responsesOverTime.length) {
    const card = document.createElement("div");
    card.className = "chart-card";
    card.innerHTML = `<h3>Responses Over Time</h3><canvas id="chart-timeline"></canvas>`;
    chartsRow.appendChild(card);
  }

  if (Object.keys(summary.ratingBreakdown).length) {
    const card = document.createElement("div");
    card.className = "chart-card";
    card.innerHTML = `<h3>Rating Averages</h3><canvas id="chart-ratings"></canvas>`;
    chartsRow.appendChild(card);
  }

  if (summary.npsPromoters + summary.npsPassives + summary.npsDetractors > 0) {
    const card = document.createElement("div");
    card.className = "chart-card";
    card.innerHTML = `<h3>NPS Distribution</h3><canvas id="chart-nps"></canvas>`;
    chartsRow.appendChild(card);
  }

  container.appendChild(chartsRow);

  const optionKeys = Object.keys(summary.optionCounts);
  if (optionKeys.length) {
    const optGrid = document.createElement("div");
    optGrid.className = "charts-grid";
    optionKeys.slice(0, 4).forEach((question, i) => {
      const card = document.createElement("div");
      card.className = "chart-card";
      card.innerHTML = `<h3>${escapeHtml(question)}</h3><canvas id="chart-opt-${i}"></canvas>`;
      optGrid.appendChild(card);
    });
    container.appendChild(optGrid);
  }

  requestAnimationFrame(() => {
    if (summary.responsesOverTime.length) {
      const ctx = document.getElementById("chart-timeline");
      if (ctx) {
        makeChart(ctx, {
          type: "line",
          data: {
            labels: summary.responsesOverTime.map(([d]) => d),
            datasets: [{
              label: "Responses",
              data: summary.responsesOverTime.map(([, c]) => c),
              borderColor: "#8b5cf6",
              backgroundColor: "rgba(139, 92, 246, 0.15)",
              fill: true,
              tension: 0.4,
            }],
          },
          options: chartOptions(),
        });
      }
    }

    if (Object.keys(summary.ratingBreakdown).length) {
      const ctx = document.getElementById("chart-ratings");
      if (ctx) {
        const labels = Object.keys(summary.ratingBreakdown);
        const data = Object.values(summary.ratingBreakdown);
        makeChart(ctx, {
          type: "bar",
          data: {
            labels: labels.map((l) => truncate(l, 30)),
            datasets: [{
              label: "Average",
              data,
              backgroundColor: CHART_COLORS,
              borderRadius: 8,
            }],
          },
          options: { ...chartOptions(), indexAxis: "y" },
        });
      }
    }

    const npsTotal = summary.npsPromoters + summary.npsPassives + summary.npsDetractors;
    if (npsTotal > 0) {
      const ctx = document.getElementById("chart-nps");
      if (ctx) {
        makeChart(ctx, {
          type: "doughnut",
          data: {
            labels: ["Promoters (9-10)", "Passives (7-8)", "Detractors (0-6)"],
            datasets: [{
              data: [summary.npsPromoters, summary.npsPassives, summary.npsDetractors],
              backgroundColor: ["#34d399", "#fbbf24", "#f87171"],
              borderWidth: 0,
            }],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: "65%",
            plugins: {
              legend: { labels: { color: "#94a3b8", font: { family: "'DM Sans', sans-serif" } } },
            },
          },
        });
      }
    }

    optionKeys.slice(0, 4).forEach((question, i) => {
      const ctx = document.getElementById(`chart-opt-${i}`);
      if (!ctx) return;
      const counts = summary.optionCounts[question];
      makeChart(ctx, {
        type: "bar",
        data: {
          labels: Object.keys(counts).map((l) => truncate(l, 20)),
          datasets: [{
            label: "Count",
            data: Object.values(counts),
            backgroundColor: CHART_COLORS,
            borderRadius: 6,
          }],
        },
        options: chartOptions(),
      });
    });
  });

  if (summary.recentResponses.length) {
    const recentCard = document.createElement("div");
    recentCard.className = "recent-card";
    recentCard.innerHTML = `<h3>Recent Submissions</h3>`;
    const table = document.createElement("div");
    table.className = "recent-table";
    for (const r of summary.recentResponses) {
      const row = document.createElement("div");
      row.className = "recent-row";
      row.innerHTML = `
        <span class="recent-name">${escapeHtml(r.participantName || r.answers?.participant_name || "Anonymous")}</span>
        <span class="recent-rating">${r.answers?.overall_rating ? `★ ${r.answers.overall_rating}/5` : ""}</span>
        <span class="recent-date">${formatDate(r.submittedAt)}</span>
      `;
      table.appendChild(row);
    }
    recentCard.appendChild(table);
    container.appendChild(recentCard);
  }
}

function chartOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { labels: { color: "#94a3b8", font: { family: "'DM Sans', sans-serif" } } },
    },
    scales: {
      x: { ticks: { color: "#64748b" }, grid: { color: "rgba(148,163,184,0.1)" } },
      y: { ticks: { color: "#64748b" }, grid: { color: "rgba(148,163,184,0.1)" } },
    },
  };
}

function truncate(s, n) {
  return s.length > n ? s.slice(0, n) + "…" : s;
}

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s;
  return d.innerHTML;
}

function formatDate(iso) {
  if (!iso) return "";
  return new Date(iso).toLocaleString(undefined, {
    month: "short", day: "numeric", hour: "2-digit", minute: "2-digit",
  });
}
