// =============================================
// Complete app.js - Improved Version
// =============================================

let table1, table2;
let rawData = { sheet1: [], sheet2: [] };
let vizChart = null;
const DECIMAL_PLACES = 2;

let ageFilterState = {
  dataTable1: { ageColumn: "", minAge: "", maxAge: "", useRoundedAge: true },
  dataTable2: { ageColumn: "", minAge: "", maxAge: "", useRoundedAge: true }
};

let ageRangeFilterRegistered = false;

// -------------------------
// Utilities
// -------------------------
function debounce(fn, delay = 300) {
  let timeout;
  return (...args) => {
    clearTimeout(timeout);
    timeout = setTimeout(() => fn(...args), delay);
  };
}

function isNumericValue(value) {
  if (value == null || value === "") return false;
  return !Number.isNaN(Number(String(value).replace(/,/g, "").trim()));
}

function parseNumber(value) {
  if (value == null || value === "") return NaN;
  const cleaned = String(value).replace(/,/g, "").replace(/[^\d.-]/g, "").trim();
  return cleaned === "" ? NaN : Number(cleaned);
}

function formatNumberForDisplay(value) {
  if (!isNumericValue(value)) return value;
  const num = parseNumber(value);
  return Number.isNaN(num) ? value : Number(num.toFixed(DECIMAL_PLACES));
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'})[m]);
}

// -------------------------
// Data Helpers
// -------------------------
function getTableById(tableId) {
  return tableId === "dataTable1" ? table1 : tableId === "dataTable2" ? table2 : null;
}

function getAllColumns(data) { 
  return data?.length ? Object.keys(data[0]) : []; 
}

function getNumericColumns(data) {
  const columns = getAllColumns(data);
  return columns.filter(col => {
    let num = 0, checked = 0;
    for (const row of data.slice(0, 500)) {
      if (row[col] != null && row[col] !== "") {
        checked++;
        if (!Number.isNaN(parseNumber(row[col]))) num++;
      }
    }
    return checked > 0 && num / checked > 0.7;
  });
}

function getPreferredAgeColumn(cols) {
  if (!cols.length) return "";
  return cols.find(c => /aoa/i.test(c)) ||
         cols.find(c => /age/i.test(c)) ||
         cols.find(c => /year/i.test(c)) || cols[0];
}

// -------------------------
// Age Filter UI
// -------------------------
function createAgeFilterControls(tableId, jsonData) {
  const tableElement = document.getElementById(tableId);
  if (!tableElement) return;

  document.getElementById(`ageFilterPanel-${tableId}`)?.remove();

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
        <select id="ageColumn-${tableId}"></select>
      </label>
      <label>Min age <input type="number" id="ageMin-${tableId}" placeholder="e.g. 8" step="1"></label>
      <label>Max age <input type="number" id="ageMax-${tableId}" placeholder="e.g. 12" step="1"></label>
      <label class="age-checkbox-label">
        <input type="checkbox" id="ageRounded-${tableId}" checked> Use rounded ages
      </label>
      <button type="button" id="applyAgeFilter-${tableId}">Apply Filter</button>
      <button type="button" id="clearAgeFilter-${tableId}">Clear</button>
      <button type="button" id="sortAgeAsc-${tableId}">Sort Age ↑</button>
      <button type="button" id="sortAgeDesc-${tableId}">Sort Age ↓</button>
    </div>
    <div class="age-filter-help">
      Enter Min and Max age, then click <strong>Apply Filter</strong>
    </div>
    <div class="age-filter-status" id="ageFilterStatus-${tableId}"></div>
  `;

  tableElement.parentNode.insertBefore(panel, tableElement);

  // Populate select
  const select = panel.querySelector(`#ageColumn-${tableId}`);
  numericColumns.forEach(col => {
    const opt = document.createElement("option");
    opt.value = col;
    opt.textContent = col;
    if (col === preferred) opt.selected = true;
    select.appendChild(opt);
  });

  if (numericColumns.length === 0) {
    panel.querySelectorAll("input,select,button").forEach(el => el.disabled = true);
  }
}

// -------------------------
// Filter Logic
// -------------------------
function setupAgeFilterEvents(tableId) {
  const column = document.getElementById(`ageColumn-${tableId}`);
  const minInput = document.getElementById(`ageMin-${tableId}`);
  const maxInput = document.getElementById(`ageMax-${tableId}`);
  const rounded = document.getElementById(`ageRounded-${tableId}`);
  const applyBtn = document.getElementById(`applyAgeFilter-${tableId}`);
  const clearBtn = document.getElementById(`clearAgeFilter-${tableId}`);
  const ascBtn = document.getElementById(`sortAgeAsc-${tableId}`);
  const descBtn = document.getElementById(`sortAgeDesc-${tableId}`);

  if (!column) return;

  function updateState() {
    ageFilterState[tableId].ageColumn = column.value;
    ageFilterState[tableId].minAge = minInput.value;
    ageFilterState[tableId].maxAge = maxInput.value;
    ageFilterState[tableId].useRoundedAge = rounded.checked;
  }

  function applyFilter() {
    updateState();
    const table = getTableById(tableId);
    if (table) {
      table.rows().invalidate("data");
      table.draw();
    }
    updateAgeFilterStatus(tableId);
  }

  column.addEventListener("change", applyFilter);
  rounded.addEventListener("change", applyFilter);
  applyBtn?.addEventListener("click", applyFilter);
  clearBtn?.addEventListener("click", () => {
    minInput.value = "";
    maxInput.value = "";
    rounded.checked = true;
    updateState();
    const table = getTableById(tableId);
    if (table) {
      table.rows().invalidate("data");
      table.search("");
      table.draw();
    }
    updateAgeFilterStatus(tableId);
  });

  ascBtn?.addEventListener("click", () => sortBySelectedAgeColumn(tableId, "asc"));
  descBtn?.addEventListener("click", () => sortBySelectedAgeColumn(tableId, "desc"));

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
    let age = parseNumber(rowData?.[state.ageColumn]);
    if (Number.isNaN(age)) return false;

    if (state.useRoundedAge) age = Math.round(age);

    if (min !== null && age < min) return false;
    if (max !== null && age > max) return false;

    return true;
  });

  ageRangeFilterRegistered = true;
}

function updateAgeFilterStatus(tableId) {
  const el = document.getElementById(`ageFilterStatus-${tableId}`);
  const table = getTableById(tableId);
  const state = ageFilterState[tableId];
  if (!el || !table) return;

  const visible = table.rows({ filter: "applied" }).count();
  const total = table.rows().count();

  if (!state.minAge && !state.maxAge) {
    el.innerHTML = `Showing <strong>${visible}</strong> of <strong>${total}</strong> rows.`;
  } else {
    el.innerHTML = `Showing <strong>${visible}</strong> of <strong>${total}</strong> rows | Filter: ${state.minAge || '—'} to ${state.maxAge || '—'} on <strong>${escapeHtml(state.ageColumn)}</strong>`;
  }
}

// -------------------------
// Load Excel
// -------------------------
async function loadExcel(filename, tableId) {
  try {
    console.log(`Loading ${filename}...`);
    const res = await fetch(filename);
    if (!res.ok) throw new Error(`HTTP ${res.status} - ${res.statusText}`);

    const arrayBuffer = await res.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const jsonData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });

    if (jsonData.length === 0) throw new Error("Sheet is empty");

    if (tableId === "dataTable1") rawData.sheet1 = jsonData;
    else rawData.sheet2 = jsonData;

    createAgeFilterControls(tableId, jsonData);

    const columns = Object.keys(jsonData[0]).map(key => ({
      title: key,
      data: key,
      render: function(data, type) {
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

    const config = {
      data: jsonData,
      columns,
      pageLength: 50,
      scrollX: true,
      scrollY: "65vh",
      scrollCollapse: true,
      deferRender: true,
      order: []
    };

    if (tableId === "dataTable1") {
      table1?.destroy();
      table1 = $("#dataTable1").DataTable(config);
      table1.on("draw", () => updateAgeFilterStatus("dataTable1"));
      setupAgeFilterEvents("dataTable1");
    } else {
      table2?.destroy();
      table2 = $("#dataTable2").DataTable(config);
      table2.on("draw", () => updateAgeFilterStatus("dataTable2"));
      setupAgeFilterEvents("dataTable2");
    }

    console.log(`✅ ${filename} loaded successfully (${jsonData.length} rows)`);
    return jsonData;

  } catch (e) {
    console.error(`Failed to load ${filename}:`, e);
    throw e;
  }
}

// -------------------------
// Tab Handling
// -------------------------
function openTab(tabIndex) {
  document.querySelectorAll(".tab-button").forEach((btn, i) => {
    btn.classList.toggle("active", i === tabIndex);
  });

  document.getElementById("tab0").style.display = tabIndex === 0 ? "block" : "none";
  document.getElementById("tab1").style.display = tabIndex === 1 ? "block" : "none";

  // Redraw DataTable when switching to hidden tab
  setTimeout(() => {
    if (tabIndex === 1 && table2) {
      table2.draw();
    }
  }, 100);
}

// -------------------------
// Initialization
// -------------------------
window.onload = async () => {
  const baseUrl = "https://raw.githubusercontent.com/cgg-projects/MWEs/main/";

  registerAgeRangeFilter();
  openTab(0);

  try {
    await loadExcel(baseUrl + "sheet1.xlsx", "dataTable1");
  } catch (e) {
    console.error("Sheet 1 failed to load", e);
  }

  // Load Sheet 2 (non-blocking)
  loadExcel(baseUrl + "sheet2.xlsx", "dataTable2")
    .then(() => {
      console.log("✅ Sheet 2 loaded");
    })
    .catch(err => {
      console.error("Sheet 2 failed", err);
      const errorDiv = document.createElement("div");
      errorDiv.style.cssText = "color: #d32f2f; background: #ffebee; padding: 15px; margin: 20px; border-radius: 8px; border-left: 5px solid #d32f2f;";
      errorDiv.innerHTML = `<strong>⚠️ Sheet 2 failed to load</strong><br>${err.message}<br><small>Check browser console (F12) for more details.</small>`;
      document.getElementById("tab1").prepend(errorDiv);
    });

  // Optional: Call your visualization functions here if you have them
  // setupVisualizationEvents();
  // updateVisualizationControls();

  openTab(0); // Ensure first tab is active
};
