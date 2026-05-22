let table1, table2;

let rawData = {
  sheet1: [],
  sheet2: []
};

let vizChart = null;

// Change this if you want fewer or more decimals.
// 2 = 12.35
// 1 = 12.3
// 0 = 12
const DECIMAL_PLACES = 2;


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

  // This rounds to the chosen number of decimal places
  // and removes unnecessary trailing zeros.
  return Number(num.toFixed(DECIMAL_PLACES));
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

    // Save original data for visualisations
    if (tableId === "dataTable1") {
      rawData.sheet1 = jsonData;
    } else {
      rawData.sheet2 = jsonData;
    }

    const columns = Object.keys(jsonData[0]).map(key => ({
      title: key,
      data: key,

      // This controls how values appear in the table.
      // Numeric values are rounded for display,
      // but still sort numerically.
      render: function (data, type, row) {
        if (isNumericValue(data)) {
          const num = parseNumber(data);

          if (type === "sort" || type === "type") {
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
    }

    if (tabIndex === 1 && table2) {
      table2.columns.adjust().draw(false);
    }

    if (tabIndex === 2) {
      renderVisualization();
    }
  }, 100);
}


// -------------------------
// Visualisation helpers
// -------------------------

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
  const min = Math.min(...values);
  const max = Math.max(...values);

  if (values.length === 0) {
    return {
      labels: [],
      counts: []
    };
  }

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

  await loadExcel("sheet1.xlsx", "dataTable1");
  await loadExcel("sheet2.xlsx", "dataTable2");

  setupVisualizationEvents();
  updateVisualizationControls();

  openTab(0);
};
