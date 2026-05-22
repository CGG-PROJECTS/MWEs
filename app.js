// =============================================
// Complete Optimized app.js
// =============================================

let table1, table2;
let rawData = {
  sheet1: [],
  sheet2: []
};
let vizChart = null;

const DECIMAL_PLACES = 2;

// Age filter state
let ageFilterState = {
  dataTable1: { ageColumn: "", minAge: "", maxAge: "", useRoundedAge: true },
  dataTable2: { ageColumn: "", minAge: "", maxAge: "", useRoundedAge: true }
};

let ageRangeFilterRegistered = false;

// -------------------------
// Utilities
// -------------------------
function debounce(fn, delay = 280) {
  let timeout;
  return function (...args) {
    clearTimeout(timeout);
    timeout = setTimeout(() => fn.apply(this, args), delay);
  };
}

function isNumericValue(value) {
  if (value === null || value === undefined || value === "") return false;
  const cleaned = String(value).replace(/,/g, "").trim();
  return cleaned !== "" && !Number.isNaN(Number(cleaned));
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return NaN;
  const cleaned = String(value).replace(/,/g, "").replace(/[^\d.-]/g, "").trim();
  return cleaned === "" ? NaN : Number(cleaned);
}

function formatNumberForDisplay(value) {
  if (!isNumericValue(value)) return value;
  const num = Number(String(value).replace(/,/g, "").trim());
  return Number.isNaN(num) ? value : Number(num.toFixed(DECIMAL_PLACES));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

// -------------------------
// Data Helpers
// -------------------------
function getTableById(tableId) {
  return tableId === "dataTable1" ? table1 : tableId === "dataTable2" ? table2 : null;
}

function getAllColumns(data) {
  return data && data.length > 0 ? Object.keys(data[0]) : [];
}

function getNumericColumns(data) {
  const columns = getAllColumns(data);
  return columns.filter(col => {
    let numericCount = 0, checkedCount = 0;
    for (const row of data.slice(0, 500)) {
      const value = row[col];
      if (value !== "" && value !== null && value !== undefined) {
        checkedCount++;
        if (!Number.isNaN(parseNumber(value))) numericCount++;
      }
    }
    return checkedCount > 0 && numericCount / checkedCount > 0.7;
  });
}

function getPreferredAgeColumn(numericColumns) {
  if (!numericColumns.length) return "";
  return numericColumns.find(c => c.toLowerCase().includes("aoa")) ||
         numericColumns.find(c => c.toLowerCase().includes("age")) ||
         numericColumns.find(c => c.toLowerCase().includes("year")) ||
         numericColumns[0];
}

// -------------------------
// Age Filter UI
// -------------------------
function createAgeFilterControls(tableId, jsonData) {
  const tableElement = document.getElementById(tableId);
  if (!tableElement) return;

  const existing = document.getElementById(`ageFilterPanel-${tableId}`);
  if (existing) existing.remove();

  const numericColumns = getNumericColumns(jsonData);
  const preferred = getPreferredAgeColumn(numericColumns);

  ageFilterState[tableId] = { ageColumn: preferred, minAge: "", maxAge: "", useRoundedAge: true };

  const panel = document.createElement("div");
  panel.className = "age-filter-panel";
  panel.id = `ageFilterPanel-${tableId}`;

  panel.innerHTML = `
    <div class="age-filter-title">Age / AoA Filter</div>
    <div class="age-filter-controls">
      <label>Age/AoA column
        <select id="ageColumn-${tableId}">
          ${numericColumns.map(col => `
            <option value="${escapeHtml(col)}" ${col === preferred ? "selected" : ""}>${escapeHtml(col)}</option>
          `).join("")}
        </select>
      </label>
      <label>Min age <input type="number" id="ageMin-${tableId}" placeholder="e.g. 8" step="1"></label>
      <label>Max age <input type="number" id="ageMax-${tableId}" placeholder="e.g. 12" step="1"></label>
      <label class="age-checkbox-label">
        <input type="checkbox" id="ageRounded-${tableId}" checked> Use rounded ages
      </label>
      <button type="button" id="applyAgeFilter-${tableId}">Apply Filter</button>
      <button type="button" id="clearAgeFilter-${tableId}">Clear</button>
      <button type="button" id="sortAgeAsc-${tableId}">Sort ↑</button>
      <button type="button" id="sortAgeDesc-${tableId}">Sort ↓</button>
    </div>
    <div class="age-filter-help">
      Example: min <strong>8</strong> – max <strong>12</strong><br>
      Rounded mode: 8.6 → 9
    </div>
    <div class="age-filter-status" id="ageFilterStatus-${tableId}"></div>
  `;

  tableElement.parentNode.insertBefore(panel, tableElement);

  if (numericColumns.length === 0) {
    panel.querySelectorAll("input,select,button").forEach(el => el.disabled = true);
  }
}

// -------------------------
// Age Filter Logic
// -------------------------
function setupAgeFilterEvents(tableId) {
  const els = {
    column: document.getElementById(`ageColumn-${tableId}`),
    min: document.getElementById(`ageMin-${tableId}`),
    max: document.getElementById(`ageMax-${tableId}`),
    rounded: document.getElementById(`ageRounded-${tableId}`),
    apply: document.getElementById(`applyAgeFilter-${tableId}`),
    clear: document.getElementById(`clearAgeFilter-${tableId}`),
    asc: document.getElementById(`sortAgeAsc-${tableId}`),
    desc: document.getElementById(`sortAgeDesc-${tableId}`)
  };

  if (!els.column) return;

  const debouncedRedraw = debounce(() => redrawTable(tableId), 280);

  function updateState() {
    ageFilterState[tableId].ageColumn = els.column.value;
    ageFilterState[tableId].minAge = els.min.value;
    ageFilterState[tableId].maxAge = els.max.value;
    ageFilterState[tableId].useRoundedAge = els.rounded.checked;
  }

  function redrawTable(id) {
    const table = getTableById(id);
    if (!table) return;
    updateState();
    table.rows().invalidate("data");
    table.draw();
    updateAgeFilterStatus(id);
  }

  els.column.addEventListener("change", () => redrawTable(tableId));
  els.min.addEventListener("input", debouncedRedraw);
  els.max.addEventListener("input", debouncedRedraw);
  els.rounded.addEventListener("change", () => redrawTable(tableId));
  if (els.apply) els.apply.addEventListener("click", () => redrawTable(tableId));

  if (els.clear) {
    els.clear.addEventListener("click", () => {
      els.min.value = ""; els.max.value = ""; els.rounded.checked = true;
      updateState();
      const table = getTableById(tableId);
      if (table) {
        table.rows().invalidate("data");
        table.search("");
        table.draw();
      }
      updateAgeFilterStatus(tableId);
    });
  }

  if (els.asc) els.asc.addEventListener("click", () => sortBySelectedAgeColumn(tableId, "asc"));
  if (els.desc) els.desc.addEventListener("click", () => sortBySelectedAgeColumn(tableId, "desc"));

  updateState();
  updateAgeFilterStatus(tableId);
}

function sortBySelectedAgeColumn(tableId, direction) {
  const table = getTableById(tableId);
  const state = ageFilterState[tableId];
  if (!table || !state.ageColumn) return;

  const cols = table.settings().init().columns || [];
  const idx = cols.findIndex(c => c.data === state.ageColumn);
  if (idx === -1) return;

  table.rows().invalidate("data");
  table.order([idx, direction]).draw();
  updateAgeFilterStatus(tableId);
}

function registerAgeRangeFilter() {
  if (ageRangeFilterRegistered) return;

  $.fn.dataTable.ext.search.push(function (settings, data, dataIndex) {
    const tableId = settings.nTable.id;
    if (tableId !== "dataTable1" && tableId !== "dataTable2") return true;

    const state = ageFilterState[tableId];
    if (!state?.ageColumn) return true;

    const min = state.minAge === "" ? null : Number(state.minAge);
    const max = state.maxAge === "" ? null : Number(state.maxAge);
    if (min === null && max === null) return true;

    const rowData = new $.fn.dataTable.Api(settings).row(dataIndex).data();
    if (!rowData) return true;

    let age = parseNumber(rowData[state.ageColumn]);
    if (Number.isNaN(age)) return false;
    if (state.useRoundedAge) age = Math.round(age);

    if (min !== null && age < min) return false;
    if (max !== null && age > max) return false;
    return true;
  });

  ageRangeFilterRegistered = true;
}

function updateAgeFilterStatus(tableId) {
  const statusEl = document.getElementById(`ageFilterStatus-${tableId}`);
  const table = getTableById(tableId);
  const state = ageFilterState[tableId];
  if (!statusEl || !table || !state) return;

  const visible = table.rows({ filter: "applied" }).count();
  const total = table.rows().count();

  if (state.minAge === "" && state.maxAge === "") {
    statusEl.innerHTML = `Showing <strong>${visible}</strong> of <strong>${total}</strong> rows. No age filter.`;
  } else {
    statusEl.innerHTML = `
      Showing <strong>${visible}</strong> of <strong>${total}</strong> rows.<br>
      Filter on <strong>${escapeHtml(state.ageColumn)}</strong>: 
      <strong>${state.minAge || '—'}</strong> to <strong>${state.maxAge || '—'}</strong> 
      (${state.useRoundedAge ? 'rounded' : 'exact'})
    `;
  }
}

// -------------------------
// Excel Loading
// -------------------------
async function loadExcel(filename, tableId) {
  try {
    const res = await fetch(filename);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    
    const arrayBuffer = await res.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });

    if (jsonData.length === 0) return;

    if (tableId === "dataTable1") rawData.sheet1 = jsonData;
    else rawData.sheet2 = jsonData;

    createAgeFilterControls(tableId, jsonData);

    const columns = Object.keys(jsonData[0]).map(key => ({
      title: key,
      data: key,
      render: function (data, type, row) {
        if (!isNumericValue(data)) return data;
        const num = parseNumber(data);

        if (type === "sort" || type === "type") {
          const state = ageFilterState[tableId];
          if (state?.useRoundedAge && state.ageColumn === key) return Math.round(num);
          return num;
        }
        if (type === "display" || type === "filter") return formatNumberForDisplay(data);
        return num;
      }
    }));

    const dtConfig = {
      data: jsonData,
      columns: columns,
      pageLength: 50,
      scrollX: true,
      scrollY: "65vh",
      scrollCollapse: true,
      deferRender: true,
      order: []
    };

    if (tableId === "dataTable1") {
      if (table1) table1.destroy();
      table1 = $("#dataTable1").DataTable(dtConfig);
      table1.on("draw", () => updateAgeFilterStatus("dataTable1"));
      setupAgeFilterEvents("dataTable1");
    } else {
      if (table2) table2.destroy();
      table2 = $("#dataTable2").DataTable(dtConfig);
      table2.on("draw", () => updateAgeFilterStatus("dataTable2"));
      setupAgeFilterEvents("dataTable2");
    }
  } catch (err) {
    console.error(err);
    alert(`Failed to load ${filename}`);
  }
}

// -------------------------
// Tabs
// -------------------------
function openTab(tabIndex) {
  document.querySelectorAll(".tab-content").forEach(t => t.style.display = "none");
  document.querySelectorAll(".tab-button").forEach(b => b.classList.remove("active"));

  document.getElementById(`tab${tabIndex}`).style.display = "block";
  document.querySelectorAll(".tab-button")[tabIndex].classList.add("active");

  setTimeout(() => {
    if (tabIndex === 0 && table1) { table1.columns.adjust().draw(false); updateAgeFilterStatus("dataTable1"); }
    if (tabIndex === 1 && table2) { table2.columns.adjust().draw(false); updateAgeFilterStatus("dataTable2"); }
    if (tabIndex === 2) renderVisualization();
  }, 150);
}

// -------------------------
// Visualization
// -------------------------
function setSelectOptions(selectId, options, preferred = null) {
  const select = document.getElementById(selectId);
  if (!select) return;
  select.innerHTML = "";
  options.forEach(opt => {
    const el = document.createElement("option");
    el.value = opt;
    el.textContent = opt;
    select.appendChild(el);
  });
  if (preferred && options.includes(preferred)) select.value = preferred;
}

function updateVisualizationControls() {
  const vizSheet = document.getElementById("vizSheet");
  if (!vizSheet) return;
  const data = rawData[vizSheet.value];
  if (!data?.length) return;

  const allCols = getAllColumns(data);
  const numCols = getNumericColumns(data);

  const prefAoA = numCols.find(c => c.toLowerCase().includes("aoa")) ||
                  numCols.find(c => c.toLowerCase().includes("age")) || numCols[0];

  const prefLabel = allCols.find(c => /mwe|phrase|expression|item/i.test(c)) || allCols[0];

  setSelectOptions("xColumn", numCols, prefAoA);
  setSelectOptions("yColumn", numCols, numCols[1] || prefAoA);
  setSelectOptions("labelColumn", allCols, prefLabel);
}

function makeHistogram(values, binCount = 20) {
  if (!values.length) return { labels: [], counts: [] };
  const min = Math.min(...values);
  const max = Math.max(...values);
  if (min === max) return { labels: [formatNumberForDisplay(min)], counts: [values.length] };

  const binSize = (max - min) / binCount;
  const counts = new Array(binCount).fill(0);
  const labels = [];

  for (let i = 0; i < binCount; i++) {
    const start = min + i * binSize;
    const end = start + binSize;
    labels.push(`${formatNumberForDisplay(start)}–${formatNumberForDisplay(end)}`);
  }

  values.forEach(v => {
    let idx = Math.floor((v - min) / binSize);
    if (idx >= binCount) idx = binCount - 1;
    counts[idx]++;
  });

  return { labels, counts };
}

function renderVisualization() {
  const sheetEl = document.getElementById("vizSheet");
  const typeEl = document.getElementById("chartType");
  const xEl = document.getElementById("xColumn");
  const yEl = document.getElementById("yColumn");
  const labelEl = document.getElementById("labelColumn");
  const canvas = document.getElementById("vizChart");

  if (!sheetEl || !typeEl || !xEl || !canvas || typeof Chart === "undefined") return;

  if (vizChart) vizChart.destroy();

  const data = rawData[sheetEl.value];
  if (!data?.length) return;

  const chartType = typeEl.value;
  const xCol = xEl.value;
  const yCol = yEl.value;
  const labelCol = labelEl.value;

  if (chartType === "histogram") {
    const values = data.map(r => parseNumber(r[xCol])).filter(v => !isNaN(v));
    const hist = makeHistogram(values);

    vizChart = new Chart(canvas, {
      type: "bar",
      data: { labels: hist.labels, datasets: [{ label: xCol, data: hist.counts, backgroundColor: "rgba(54,162,235,0.7)" }] },
      options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true } } }
    });
  } 
  else if (chartType === "bar") {
    const rows = data.map(r => ({ label: r[labelCol] || "", value: parseNumber(r[xCol]) }))
                    .filter(r => !isNaN(r.value))
                    .sort((a,b) => b.value - a.value)
                    .slice(0, 25);

    vizChart = new Chart(canvas, {
      type: "bar",
      data: {
        labels: rows.map(r => String(r.label).slice(0,40)),
        datasets: [{ label: xCol, data: rows.map(r => r.value), backgroundColor: "rgba(255,159,64,0.7)" }]
      },
      options: { responsive: true, maintainAspectRatio: false, indexAxis: "y" }
    });
  } 
  else if (chartType === "scatter") {
    let points = data.map(r => ({ x: parseNumber(r[xCol]), y: parseNumber(r[yCol]) }))
                    .filter(p => !isNaN(p.x) && !isNaN(p.y));

    if (points.length > 3000) points = points.sort(() => 0.5 - Math.random()).slice(0, 3000);

    vizChart = new Chart(canvas, {
      type: "scatter",
      data: { datasets: [{ label: `${yCol} vs ${xCol}`, data: points, backgroundColor: "rgba(75,192,192,0.6)" }] },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }
}

function setupVisualizationEvents() {
  const els = ["vizSheet", "chartType", "xColumn", "yColumn", "labelColumn"].map(id => document.getElementById(id));
  if (!els[0]) return;

  els.forEach(el => {
    if (el) el.addEventListener("change", () => {
      if (el.id === "vizSheet") updateVisualizationControls();
      renderVisualization();
    });
  });

  const updateBtn = document.getElementById("updatePlot");
  if (updateBtn) updateBtn.addEventListener("click", renderVisualization);
}

// -------------------------
// Initialize
// -------------------------
window.onload = async () => {
  registerAgeRangeFilter();
  openTab(0);

  await loadExcel("sheet1.xlsx", "dataTable1");
  await loadExcel("sheet2.xlsx", "dataTable2");

  setupVisualizationEvents();
  updateVisualizationControls();
  openTab(0);
};
