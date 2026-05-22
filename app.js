let table1, table2;

let rawData = {
  sheet1: [],
  sheet2: []
};

let vizChart = null;

// Change this if you want fewer or more decimals in the displayed table.
// 2 = 12.35
// 1 = 12.3
// 0 = 12
const DECIMAL_PLACES = 2;

// Age range filter settings for each table
let ageFilterState = {
  dataTable1: {
    ageColumn: "",
    minAge: "",
    maxAge: "",
    useRoundedAge: true
  },
  dataTable2: {
    ageColumn: "",
    minAge: "",
    maxAge: "",
    useRoundedAge: true
  }
};

let ageRangeFilterRegistered = false;


// -------------------------
// Number formatting helpers
// -------------------------

function isNumericValue(value) {
  if (value === null || value === undefined || value === "") return false;

  const cleaned = String(value)
    .replace(/,/g, "")
    .trim();

  return cleaned !== "" && !Number.isNaN(Number(cleaned));
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return NaN;

  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "")
    .trim();

  if (cleaned === "") return NaN;

  return Number(cleaned);
}

function formatNumberForDisplay(value) {
  if (!isNumericValue(value)) return value;

  const num = Number(String(value).replace(/,/g, "").trim());

  if (Number.isNaN(num)) return value;

  return Number(num.toFixed(DECIMAL_PLACES));
}

function roundAgeValue(value) {
  const num = parseNumber(value);

  if (Number.isNaN(num)) return NaN;

  return Math.round(num);
}


// -------------------------
// Data/table helpers
// -------------------------

function getTableById(tableId) {
  if (tableId === "dataTable1") return table1;
  if (tableId === "dataTable2") return table2;
  return null;
}

function getAllColumns(data) {
  if (!data || data.length === 0) return [];
  return Object.keys(data[0]);
}

function getNumericColumns(data) {
  const columns = getAllColumns(data);

  return columns.filter(col => {
    let numericCount = 0;
    let checkedCount = 0;

    for (const row of data.slice(0, 500)) {
      const value = row[col];

      if (value !== "" && value !== null && value !== undefined) {
        checkedCount++;

        if (!Number.isNaN(parseNumber(value))) {
          numericCount++;
        }
      }
    }

    return checkedCount > 0 && numericCount / checkedCount > 0.7;
  });
}

function getPreferredAgeColumn(numericColumns) {
  if (!numericColumns || numericColumns.length === 0) return "";

  return (
    numericColumns.find(c => c.toLowerCase().includes("aoa")) ||
    numericColumns.find(c => c.toLowerCase().includes("age")) ||
    numericColumns.find(c => c.toLowerCase().includes("year")) ||
    numericColumns[0]
  );
}


// -------------------------
// Age filter controls
// -------------------------

function createAgeFilterControls(tableId, jsonData) {
  const tableElement = document.getElementById(tableId);

  if (!tableElement) return;

  const existingPanel = document.getElementById(`ageFilterPanel-${tableId}`);

  if (existingPanel) {
    existingPanel.remove();
  }

  const numericColumns = getNumericColumns(jsonData);
  const preferredAgeColumn = getPreferredAgeColumn(numericColumns);

  ageFilterState[tableId].ageColumn = preferredAgeColumn;
  ageFilterState[tableId].minAge = "";
  ageFilterState[tableId].maxAge = "";
  ageFilterState[tableId].useRoundedAge = true;

  const panel = document.createElement("div");
  panel.className = "age-filter-panel";
  panel.id = `ageFilterPanel-${tableId}`;

  panel.innerHTML = `
    <div class="age-filter-title">
      Age / AoA filter
    </div>

    <div class="age-filter-controls">
      <label>
        Age/AoA column
        <select id="ageColumn-${tableId}">
          ${
            numericColumns.length > 0
              ? numericColumns.map(col => `
                  <option value="${escapeHtml(col)}" ${col === preferredAgeColumn ? "selected" : ""}>
                    ${escapeHtml(col)}
                  </option>
                `).join("")
              : `<option value="">No numeric columns found</option>`
          }
        </select>
      </label>

      <label>
        Min age
        <input
          type="number"
          id="ageMin-${tableId}"
          placeholder="e.g. 8"
          step="1"
        >
      </label>

      <label>
        Max age
        <input
          type="number"
          id="ageMax-${tableId}"
          placeholder="e.g. 12"
          step="1"
        >
      </label>

      <label class="age-checkbox-label">
        <input
          type="checkbox"
          id="ageRounded-${tableId}"
          checked
        >
        Use rounded ages
      </label>

      <button type="button" id="applyAgeFilter-${tableId}">
        Apply age filter
      </button>

      <button type="button" id="clearAgeFilter-${tableId}">
        Clear
      </button>

      <button type="button" id="sortAgeAsc-${tableId}">
        Sort age ↑
      </button>

      <button type="button" id="sortAgeDesc-${tableId}">
        Sort age ↓
      </button>
    </div>

    <div class="age-filter-help">
      Example: enter min <strong>8</strong> and max <strong>12</strong> to show all items with AoA between 8–12 years.
      If rounded ages is enabled, <strong>8.6</strong> is treated as <strong>9</strong>.
    </div>

    <div class="age-filter-status" id="ageFilterStatus-${tableId}">
    </div>
  `;

  tableElement.parentNode.insertBefore(panel, tableElement);

  if (numericColumns.length === 0) {
    panel.querySelectorAll("input, button, select").forEach(el => {
      el.disabled = true;
    });
  }
}

function setupAgeFilterEvents(tableId) {
  const ageColumn = document.getElementById(`ageColumn-${tableId}`);
  const ageMin = document.getElementById(`ageMin-${tableId}`);
  const ageMax = document.getElementById(`ageMax-${tableId}`);
  const ageRounded = document.getElementById(`ageRounded-${tableId}`);
  const applyButton = document.getElementById(`applyAgeFilter-${tableId}`);
  const clearButton = document.getElementById(`clearAgeFilter-${tableId}`);
  const sortAscButton = document.getElementById(`sortAgeAsc-${tableId}`);
  const sortDescButton = document.getElementById(`sortAgeDesc-${tableId}`);

  if (!ageColumn || !ageMin || !ageMax || !ageRounded) return;

  function updateStateFromControls() {
    ageFilterState[tableId].ageColumn = ageColumn.value;
    ageFilterState[tableId].minAge = ageMin.value;
    ageFilterState[tableId].maxAge = ageMax.value;
    ageFilterState[tableId].useRoundedAge = ageRounded.checked;
  }

  function redrawTable() {
    const table = getTableById(tableId);

    if (!table) return;

    updateStateFromControls();

    // Invalidate forces DataTables to re-check sort values,
    // including rounded sort values for the selected age column.
    table.rows().invalidate("data");
    table.draw();
    updateAgeFilterStatus(tableId);
  }

  ageColumn.addEventListener("change", redrawTable);
  ageMin.addEventListener("input", redrawTable);
  ageMax.addEventListener("input", redrawTable);
  ageRounded.addEventListener("change", redrawTable);

  if (applyButton) {
    applyButton.addEventListener("click", redrawTable);
  }

  if (clearButton) {
    clearButton.addEventListener("click", () => {
      ageMin.value = "";
      ageMax.value = "";
      ageRounded.checked = true;

      updateStateFromControls();

      const table = getTableById(tableId);

      if (table) {
        table.rows().invalidate("data");
        table.search("");
        table.draw();
      }

      updateAgeFilterStatus(tableId);
    });
  }

  if (sortAscButton) {
    sortAscButton.addEventListener("click", () => {
      sortBySelectedAgeColumn(tableId, "asc");
    });
  }

  if (sortDescButton) {
    sortDescButton.addEventListener("click", () => {
      sortBySelectedAgeColumn(tableId, "desc");
    });
  }

  updateStateFromControls();
  updateAgeFilterStatus(tableId);
}

function sortBySelectedAgeColumn(tableId, direction) {
  const table = getTableById(tableId);
  const state = ageFilterState[tableId];

  if (!table || !state || !state.ageColumn) return;

  const columns = table.settings().init().columns || [];
  const columnIndex = columns.findIndex(col => col.data === state.ageColumn);

  if (columnIndex === -1) return;

  table.rows().invalidate("data");
  table.order([columnIndex, direction]).draw();

  updateAgeFilterStatus(tableId);
}

function registerAgeRangeFilter() {
  if (ageRangeFilterRegistered) return;

  $.fn.dataTable.ext.search.push(function (settings, data, dataIndex) {
    const tableId = settings.nTable.id;

    if (tableId !== "dataTable1" && tableId !== "dataTable2") {
      return true;
    }

    const state = ageFilterState[tableId];

    if (!state || !state.ageColumn) {
      return true;
    }

    const minAge = state.minAge === "" ? null : Number(state.minAge);
    const maxAge = state.maxAge === "" ? null : Number(state.maxAge);

    // If both are empty, do not filter by age.
    if (minAge === null && maxAge === null) {
      return true;
    }

    const tableApi = new $.fn.dataTable.Api(settings);
    const rowData = tableApi.row(dataIndex).data();

    if (!rowData) return true;

    let ageValue = parseNumber(rowData[state.ageColumn]);

    if (Number.isNaN(ageValue)) {
      return false;
    }

    if (state.useRoundedAge) {
      ageValue = Math.round(ageValue);
    }

    if (minAge !== null && ageValue < minAge) {
      return false;
    }

    if (maxAge !== null && ageValue > maxAge) {
      return false;
    }

    return true;
  });

  ageRangeFilterRegistered = true;
}

function updateAgeFilterStatus(tableId) {
  const status = document.getElementById(`ageFilterStatus-${tableId}`);
  const table = getTableById(tableId);
  const state = ageFilterState[tableId];

  if (!status || !table || !state) return;

  const visibleRows = table.rows({ filter: "applied" }).count();
  const totalRows = table.rows().count();

  const minText = state.minAge === "" ? "no min" : state.minAge;
  const maxText = state.maxAge === "" ? "no max" : state.maxAge;
  const roundedText = state.useRoundedAge ? "rounded to nearest year" : "exact decimals";

  if (state.minAge === "" && state.maxAge === "") {
    status.innerHTML = `
      Showing <strong>${visibleRows}</strong> of <strong>${totalRows}</strong> rows.
      No age range filter active.
    `;
  } else {
    status.innerHTML = `
      Showing <strong>${visibleRows}</strong> of <strong>${totalRows}</strong> rows.
      Age filter on <strong>${escapeHtml(state.ageColumn)}</strong>:
      <strong>${minText}</strong> to <strong>${maxText}</strong>
      using <strong>${roundedText}</strong>.
    `;
  }
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
// Excel loading
// -------------------------

async function loadExcel(filename, tableId) {
  try {
    console.log(`Loading ${filename}...`);

    const response = await fetch(filename);

    if (!response.ok) {
      throw new Error(`HTTP ${response.status} while loading ${filename}`);
    }

    const arrayBuffer = await response.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
      type: "array"
    });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
      defval: ""
    });

    if (jsonData.length === 0) {
      console.error(`No data found in ${filename}`);
      return;
    }

    // Save original data for visualisations.
    if (tableId === "dataTable1") {
      rawData.sheet1 = jsonData;
    } else {
      rawData.sheet2 = jsonData;
    }

    createAgeFilterControls(tableId, jsonData);

    const columns = Object.keys(jsonData[0]).map(key => ({
      title: key,
      data: key,

      // This controls how values appear in the table.
      // Numeric values are rounded for display.
      // Sorting remains numeric.
      // If this is the selected age/AoA column and "Use rounded ages" is checked,
      // the sort also uses nearest-year rounded values.
      render: function (data, type, row) {
        if (isNumericValue(data)) {
          const num = parseNumber(data);

          if (type === "sort" || type === "type") {
            const state = ageFilterState[tableId];

            if (
              state &&
              state.useRoundedAge &&
              state.ageColumn === key
            ) {
              return Math.round(num);
            }

            return num;
          }

          if (type === "display" || type === "filter") {
            return formatNumberForDisplay(data);
          }

          return num;
        }

        return data;
      }
    }));

    if (tableId === "dataTable1") {
      if (table1) {
        table1.destroy();
        $("#dataTable1").empty();
      }

      table1 = $("#dataTable1").DataTable({
        data: jsonData,
        columns: columns,
        pageLength: 25,
        scrollX: true,
        scrollY: "65vh",
        scrollCollapse: true,
        order: []
      });

      table1.on("draw", () => {
        updateAgeFilterStatus("dataTable1");
      });

      setupAgeFilterEvents("dataTable1");

      console.log(`Loaded ${filename} into dataTable1`);
    } else {
      if (table2) {
        table2.destroy();
        $("#dataTable2").empty();
      }

      table2 = $("#dataTable2").DataTable({
        data: jsonData,
        columns: columns,
        pageLength: 25,
        scrollX: true,
        scrollY: "65vh",
        scrollCollapse: true,
        order: []
      });

      table2.on("draw", () => {
        updateAgeFilterStatus("dataTable2");
      });

      setupAgeFilterEvents("dataTable2");

      console.log(`Loaded ${filename} into dataTable2`);
    }

  } catch (error) {
    console.error(error);
    alert(`Error loading ${filename}: ${error.message}`);
  }
}


// -------------------------
// Tab control
// -------------------------

function openTab(tabIndex) {
  document.querySelectorAll(".tab-content").forEach(tab => {
    tab.style.display = "none";
  });

  document.querySelectorAll(".tab-button").forEach(btn => {
    btn.classList.remove("active");
  });

  const selectedTab = document.getElementById(`tab${tabIndex}`);
  const selectedButton = document.querySelectorAll(".tab-button")[tabIndex];

  if (selectedTab) {
    selectedTab.style.display = "block";
  }

  if (selectedButton) {
    selectedButton.classList.add("active");
  }

  setTimeout(() => {
    if (tabIndex === 0 && table1) {
      table1.columns.adjust().draw(false);
      updateAgeFilterStatus("dataTable1");
    }

    if (tabIndex === 1 && table2) {
      table2.columns.adjust().draw(false);
      updateAgeFilterStatus("dataTable2");
    }

    if (tabIndex === 2) {
      renderVisualization();
    }
  }, 100);
}


// -------------------------
// Visualisation helpers
// -------------------------

function setSelectOptions(selectId, options, preferredValue = null) {
  const select = document.getElementById(selectId);

  if (!select) return;

  select.innerHTML = "";

  options.forEach(option => {
    const el = document.createElement("option");
    el.value = option;
    el.textContent = option;
    select.appendChild(el);
  });

  if (preferredValue && options.includes(preferredValue)) {
    select.value = preferredValue;
  }
}

function updateVisualizationControls() {
  const vizSheet = document.getElementById("vizSheet");

  if (!vizSheet) return;

  const sheetName = vizSheet.value;
  const data = rawData[sheetName];

  if (!data || data.length === 0) return;

  const allColumns = getAllColumns(data);
  const numericColumns = getNumericColumns(data);

  const preferredAoA =
    numericColumns.find(c => c.toLowerCase().includes("aoa")) ||
    numericColumns.find(c => c.toLowerCase().includes("age")) ||
    numericColumns[0];

  const preferredLabel =
    allColumns.find(c => c.toLowerCase().includes("mwe")) ||
    allColumns.find(c => c.toLowerCase().includes("phrase")) ||
    allColumns.find(c => c.toLowerCase().includes("expression")) ||
    allColumns.find(c => c.toLowerCase().includes("item")) ||
    allColumns[0];

  setSelectOptions("xColumn", numericColumns, preferredAoA);
  setSelectOptions("yColumn", numericColumns, numericColumns[1] || preferredAoA);
  setSelectOptions("labelColumn", allColumns, preferredLabel);
}

function makeHistogram(values, binCount = 20) {
  if (values.length === 0) {
    return {
      labels: [],
      counts: []
    };
  }

  const min = Math.min(...values);
  const max = Math.max(...values);

  if (min === max) {
    return {
      labels: [String(formatNumberForDisplay(min))],
      counts: [values.length]
    };
  }

  const binSize = (max - min) / binCount;
  const counts = new Array(binCount).fill(0);
  const labels = [];

  for (let i = 0; i < binCount; i++) {
    const start = min + i * binSize;
    const end = start + binSize;

    labels.push(
      `${formatNumberForDisplay(start)}–${formatNumberForDisplay(end)}`
    );
  }

  values.forEach(value => {
    let index = Math.floor((value - min) / binSize);

    if (index >= binCount) {
      index = binCount - 1;
    }

    counts[index]++;
  });

  return {
    labels,
    counts
  };
}


// -------------------------
// Visualisation rendering
// -------------------------

function renderVisualization() {
  const vizSheet = document.getElementById("vizSheet");
  const chartTypeEl = document.getElementById("chartType");
  const xColumnEl = document.getElementById("xColumn");
  const yColumnEl = document.getElementById("yColumn");
  const labelColumnEl = document.getElementById("labelColumn");
  const chartCanvas = document.getElementById("vizChart");

  // If the visualisation tab has not been added to index.html yet,
  // this prevents errors.
  if (!vizSheet || !chartTypeEl || !xColumnEl || !yColumnEl || !labelColumnEl || !chartCanvas) {
    return;
  }

  if (typeof Chart === "undefined") {
    console.warn("Chart.js is not loaded.");
    return;
  }

  const sheetName = vizSheet.value;
  const chartType = chartTypeEl.value;
  const xColumn = xColumnEl.value;
  const yColumn = yColumnEl.value;
  const labelColumn = labelColumnEl.value;

  const data = rawData[sheetName];

  if (!data || data.length === 0) {
    console.warn("No data available for visualisation.");
    return;
  }

  if (!xColumn) {
    updateVisualizationControls();
    return;
  }

  if (vizChart) {
    vizChart.destroy();
  }

  const ctx = chartCanvas;

  if (chartType === "histogram") {
    const values = data
      .map(row => parseNumber(row[xColumn]))
      .filter(value => !Number.isNaN(value));

    const histogram = makeHistogram(values, 20);

    vizChart = new Chart(ctx, {
      type: "bar",
      data: {
        labels: histogram.labels,
        datasets: [{
          label: `Distribution of ${xColumn}`,
          data: histogram.counts,
          backgroundColor: "rgba(54, 162, 235, 0.65)",
          borderColor: "rgba(54, 162, 235, 1)",
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          title: {
            display: true,
            text: `Histogram of ${xColumn}`
          }
        },
        scales: {
          x: {
            title: {
              display: true,
              text: xColumn
            }
          },
          y: {
            title: {
              display: true,
              text: "Count"
            }
          }
        }
      }
    });
  }

  if (chartType === "bar") {
    const rows = data
      .map(row => ({
        label: row[labelColumn] || "",
        value: parseNumber(row[xColumn])
      }))
      .filter(row => !Number.isNaN(row.value))
      .sort((a, b) => b.value - a.value)
      .slice(0, 25);

    vizChart = new Chart(ctx, {
      type: "bar",
      data: {
        labels: rows.map(row => String(row.label).slice(0, 40)),
        datasets: [{
          label: xColumn,
          data: rows.map(row => Number(row.value.toFixed(DECIMAL_PLACES))),
          backgroundColor: "rgba(255, 159, 64, 0.65)",
          borderColor: "rgba(255, 159, 64, 1)",
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: "y",
        plugins: {
          title: {
            display: true,
            text: `Top 25 values by ${xColumn}`
          }
        },
        scales: {
          x: {
            title: {
              display: true,
              text: xColumn
            }
          }
        }
      }
    });
  }

  if (chartType === "scatter") {
    let points = data
      .map(row => ({
        x: parseNumber(row[xColumn]),
        y: parseNumber(row[yColumn])
      }))
      .filter(point => !Number.isNaN(point.x) && !Number.isNaN(point.y));

    // Limit scatterplots for browser performance.
    if (points.length > 3000) {
      points = points.sort(() => 0.5 - Math.random()).slice(0, 3000);
    }

    points = points.map(point => ({
      x: Number(point.x.toFixed(DECIMAL_PLACES)),
      y: Number(point.y.toFixed(DECIMAL_PLACES))
    }));

    vizChart = new Chart(ctx, {
      type: "scatter",
      data: {
        datasets: [{
          label: `${yColumn} vs ${xColumn}`,
          data: points,
          backgroundColor: "rgba(75, 192, 192, 0.55)"
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          title: {
            display: true,
            text: `${yColumn} vs ${xColumn}`
          }
        },
        scales: {
          x: {
            title: {
              display: true,
              text: xColumn
            }
          },
          y: {
            title: {
              display: true,
              text: yColumn
            }
          }
        }
      }
    });
  }
}


// -------------------------
// Visualisation events
// -------------------------

function setupVisualizationEvents() {
  const vizSheet = document.getElementById("vizSheet");
  const chartType = document.getElementById("chartType");
  const xColumn = document.getElementById("xColumn");
  const yColumn = document.getElementById("yColumn");
  const labelColumn = document.getElementById("labelColumn");
  const updateButton = document.getElementById("updatePlot");

  // If visualisation controls do not exist yet, do nothing.
  if (!vizSheet || !chartType || !xColumn || !yColumn || !labelColumn) {
    return;
  }

  vizSheet.addEventListener("change", () => {
    updateVisualizationControls();
    renderVisualization();
  });

  chartType.addEventListener("change", () => {
    renderVisualization();
  });

  xColumn.addEventListener("change", () => {
    renderVisualization();
  });

  yColumn.addEventListener("change", () => {
    renderVisualization();
  });

  labelColumn.addEventListener("change", () => {
    renderVisualization();
  });

  if (updateButton) {
    updateButton.addEventListener("click", () => {
      renderVisualization();
    });
  }
}


// -------------------------
// Load everything
// -------------------------

window.onload = async () => {
  openTab(0);

  registerAgeRangeFilter();

  await loadExcel("sheet1.xlsx", "dataTable1");
  await loadExcel("sheet2.xlsx", "dataTable2");

  setupVisualizationEvents();
  updateVisualizationControls();

  openTab(0);
};
