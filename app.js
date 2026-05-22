let table1 = null;
let table2 = null;

let rawData = {
  sheet1: [],
  sheet2: []
};

let vizChart = null;

// Change this for table display rounding.
// 2 = 12.35
// 1 = 12.3
// 0 = 12
const DECIMAL_PLACES = 2;

// If you want fixed names instead of the Excel sheet names,
// change displayName below.
// For example:
// dataTable2: { filename: "sheet2.xlsx", rawKey: "sheet2", displayName: "My Original Sheet 2 Name" }
const FILE_CONFIG = {
  dataTable1: {
    filename: "sheet1.xlsx",
    rawKey: "sheet1",
    displayName: "Sheet 1"
  },
  dataTable2: {
    filename: "sheet2.xlsx",
    rawKey: "sheet2",
    displayName: "Sheet 2"
  }
};

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

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return NaN;

  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "")
    .trim();

  if (cleaned === "" || cleaned === "-" || cleaned === "." || cleaned === "-.") {
    return NaN;
  }

  return Number(cleaned);
}

function isNumericValue(value) {
  return !Number.isNaN(parseNumber(value));
}

function formatNumberForDisplay(value) {
  if (!isNumericValue(value)) return value;

  const num = parseNumber(value);

  if (Number.isNaN(num)) return value;

  return Number(num.toFixed(DECIMAL_PLACES));
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
// General table helpers
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

function updateSheetLabels(tableId, excelSheetName) {
  const config = FILE_CONFIG[tableId];

  let label = config.displayName || excelSheetName || tableId;

  // If the display name is generic, use the Excel sheet name where possible.
  if (
    excelSheetName &&
    config.displayName &&
    ["Sheet 1", "Sheet 2", "sheet1", "sheet2"].includes(config.displayName)
  ) {
    label = excelSheetName;
  }

  const titleEl = document.getElementById(`tableTitle-${tableId}`);
  const tabButtonEl = document.getElementById(`tabButton-${tableId}`);

  if (titleEl) {
    titleEl.textContent = label;
  }

  if (tabButtonEl) {
    tabButtonEl.textContent = label;
  }
}


// -------------------------
// Age / AoA filter controls
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

  const optionsHtml = numericColumns.length > 0
    ? numericColumns.map(col => {
      const selected = col === preferredAgeColumn ? "selected" : "";
      return `<option value="${escapeHtml(col)}" ${selected}>${escapeHtml(col)}</option>`;
    }).join("")
    : `<option value="">No numeric columns found</option>`;

  panel.innerHTML = `
    <div class="age-filter-title">Age / AoA filter</div>

    <div class="age-filter-controls">
      <label>
        Age/AoA column
        <select id="ageColumn-${tableId}">
          ${optionsHtml}
        </select>
      </label>

      <label>
        Min age
        <input id="ageMin-${tableId}" type="number" step="0.1" placeholder="e.g. 8">
      </label>

      <label>
        Max age
        <input id="ageMax-${tableId}" type="number" step="0.1" placeholder="e.g. 12">
      </label>

      <label class="age-checkbox-label">
        <input id="ageRounded-${tableId}" type="checkbox" checked>
        Use rounded ages
      </label>

      <button id="applyAgeFilter-${tableId}" type="button">Apply age filter</button>
      <button id="clearAgeFilter-${tableId}" type="button">Clear age filter</button>
      <button id="sortAgeAsc-${tableId}" type="button">Sort age ↑</button>
      <button id="sortAgeDesc-${tableId}" type="button">Sort age ↓</button>
    </div>

    <div class="age-filter-help">
      Example: min <strong>8</strong> and max <strong>12</strong> shows AoA values between 8–12.
      If rounded ages is enabled, 8.6 is treated as 9.
      This filter now updates only when you press <strong>Apply</strong> or press <strong>Enter</strong>, so typing should not freeze.
    </div>

    <div id="ageFilterStatus-${tableId}" class="age-filter-status"></div>
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

  function applyAgeFilter() {
    const table = getTableById(tableId);

    if (!table) return;

    updateStateFromControls();

    // Only redraw when Apply is pressed.
    // This avoids freezing while typing.
    table.rows().invalidate("data");
    table.draw();

    updateAgeFilterStatus(tableId);
  }

  function clearAgeFilter() {
    const table = getTableById(tableId);

    ageMin.value = "";
    ageMax.value = "";
    ageRounded.checked = true;

    updateStateFromControls();

    if (table) {
      table.rows().invalidate("data");
      table.draw();
    }

    updateAgeFilterStatus(tableId);
  }

  function applyOnEnter(event) {
    if (event.key === "Enter") {
      applyAgeFilter();
    }
  }

  // Do NOT redraw on every keystroke.
  // Only update when Apply is clicked, Enter is pressed, column changes, or rounded checkbox changes.
  ageMin.addEventListener("keydown", applyOnEnter);
  ageMax.addEventListener("keydown", applyOnEnter);

  ageColumn.addEventListener("change", applyAgeFilter);
  ageRounded.addEventListener("change", applyAgeFilter);

  if (applyButton) {
    applyButton.addEventListener("click", applyAgeFilter);
  }

  if (clearButton) {
    clearButton.addEventListener("click", clearAgeFilter);
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

    const minAge = state.minAge === "" ? null : parseNumber(state.minAge);
    const maxAge = state.maxAge === "" ? null : parseNumber(state.maxAge);

    // If both are empty, do not filter.
    if (
      (minAge === null || Number.isNaN(minAge)) &&
      (maxAge === null || Number.isNaN(maxAge))
    ) {
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

    if (minAge !== null && !Number.isNaN(minAge) && ageValue < minAge) {
      return false;
    }

    if (maxAge !== null && !Number.isNaN(maxAge) && ageValue > maxAge) {
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
  const roundedText = state.useRoundedAge ? "rounded to nearest year" : "exact decimal age";

  if (state.minAge === "" && state.maxAge === "") {
    status.innerHTML = `
      Showing <strong>${visibleRows}</strong> of <strong>${totalRows}</strong> rows.
      No age range filter active.
    `;
  } else {
    status.innerHTML = `
      Showing <strong>${visibleRows}</strong> of <strong>${totalRows}</strong> rows.
      Age filter on <strong>${escapeHtml(state.ageColumn)}</strong>:
      <strong>${minText}</strong> to <strong>${maxText}</strong>,
      using <strong>${roundedText}</strong>.
    `;
  }
}


// -------------------------
// Excel loading
// -------------------------

async function loadExcel(tableId) {
  const config = FILE_CONFIG[tableId];

  if (!config) {
    console.error(`No config found for ${tableId}`);
    return;
  }

  try {
    console.log(`Loading ${config.filename}...`);

    const response = await fetch(config.filename);

    if (!response.ok) {
      throw new Error(`HTTP ${response.status} while loading ${config.filename}`);
    }

    const arrayBuffer = await response.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
      type: "array"
    });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    updateSheetLabels(tableId, sheetName);

    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
      defval: ""
    });

    if (jsonData.length === 0) {
      console.error(`No data found in ${config.filename}`);
      return;
    }

    rawData[config.rawKey] = jsonData;

    createAgeFilterControls(tableId, jsonData);

    const columns = Object.keys(jsonData[0]).map(key => ({
      title: key,
      data: key,
      render: function (data, type) {
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

        // Performance improvements
        deferRender: true,
        processing: true,
        searchDelay: 500,
        autoWidth: false,

        pageLength: 25,
        lengthMenu: [10, 25, 50, 100],
        scrollX: true,
        scrollY: "65vh",
        scrollCollapse: true,
        order: []
      });

      table1.on("draw", () => {
        updateAgeFilterStatus("dataTable1");
      });

      setupAgeFilterEvents("dataTable1");

      console.log(`Loaded ${config.filename} into dataTable1`);
    }

    if (tableId === "dataTable2") {
      if (table2) {
        table2.destroy();
        $("#dataTable2").empty();
      }

      table2 = $("#dataTable2").DataTable({
        data: jsonData,
        columns: columns,

        // Performance improvements
        deferRender: true,
        processing: true,
        searchDelay: 500,
        autoWidth: false,

        pageLength: 25,
        lengthMenu: [10, 25, 50, 100],
        scrollX: true,
        scrollY: "65vh",
        scrollCollapse: true,
        order: []
      });

      table2.on("draw", () => {
        updateAgeFilterStatus("dataTable2");
      });

      setupAgeFilterEvents("dataTable2");

      console.log(`Loaded ${config.filename} into dataTable2`);
    }

  } catch (error) {
    console.error(error);
    alert(`Error loading ${config.filename}: ${error.message}`);
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
      table1.columns.adjust();
      updateAgeFilterStatus("dataTable1");
    }

    if (tabIndex === 1 && table2) {
      table2.columns.adjust();
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
    allColumns.find(c => c.toLowerCase().includes("word")) ||
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

    labels.push(`${formatNumberForDisplay(start)}–${formatNumberForDisplay(end)}`);
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
        animation: false,
        plugins: {
          title: {
            display: true,
            text: `Histogram of ${xColumn}`,
            font: {
              size: 18
            }
          },
          legend: {
            labels: {
              font: {
                size: 14
              }
            }
          }
        },
        scales: {
          x: {
            ticks: {
              font: {
                size: 12
              }
            },
            title: {
              display: true,
              text: xColumn,
              font: {
                size: 15
              }
            }
          },
          y: {
            ticks: {
              font: {
                size: 12
              }
            },
            title: {
              display: true,
              text: "Count",
              font: {
                size: 15
              }
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
      .slice(0, 30);

    vizChart = new Chart(ctx, {
      type: "bar",
      data: {
        labels: rows.map(row => String(row.label).slice(0, 60)),
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
        animation: false,
        indexAxis: "y",
        plugins: {
          title: {
            display: true,
            text: `Top 30 values by ${xColumn}`,
            font: {
              size: 18
            }
          },
          legend: {
            labels: {
              font: {
                size: 14
              }
            }
          }
        },
        scales: {
          x: {
            ticks: {
              font: {
                size: 12
              }
            },
            title: {
              display: true,
              text: xColumn,
              font: {
                size: 15
              }
            }
          },
          y: {
            ticks: {
              font: {
                size: 12
              }
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

    // Keep chart responsive by limiting plotted points.
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
        animation: false,
        plugins: {
          title: {
            display: true,
            text: `${yColumn} vs ${xColumn}`,
            font: {
              size: 18
            }
          },
          legend: {
            labels: {
              font: {
                size: 14
              }
            }
          }
        },
        scales: {
          x: {
            ticks: {
              font: {
                size: 12
              }
            },
            title: {
              display: true,
              text: xColumn,
              font: {
                size: 15
              }
            }
          },
          y: {
            ticks: {
              font: {
                size: 12
              }
            },
            title: {
              display: true,
              text: yColumn,
              font: {
                size: 15
              }
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

  await loadExcel("dataTable1");
  await loadExcel("dataTable2");

  setupVisualizationEvents();
  updateVisualizationControls();

  openTab(0);
};
