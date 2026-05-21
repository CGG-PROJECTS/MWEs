let table1, table2;

async function loadExcel(filename, tableId) {
  try {
    const response = await fetch(filename);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    if (jsonData.length === 0) {
      console.error("No data in sheet");
      return;
    }

    // Initialize or reload DataTable
    if (tableId === 'dataTable1') {
      if (table1) table1.destroy();
      table1 = $('#dataTable1').DataTable({
        data: jsonData,
        columns: Object.keys(jsonData[0]).map(key => ({ title: key, data: key })),
        pageLength: 25,
        scrollX: true,
        scrollY: "65vh",
        scrollCollapse: true,
        order: [],
      });
    } else {
      if (table2) table2.destroy();
      table2 = $('#dataTable2').DataTable({
        data: jsonData,
        columns: Object.keys(jsonData[0]).map(key => ({ title: key, data: key })),
        pageLength: 25,
        scrollX: true,
        scrollY: "65vh",
        scrollCollapse: true,
        order: [],
      });
    }

  } catch (error) {
    console.error(error);
    alert(`Error loading ${filename}`);
  }
}

// Tab control
function openTab(tabIndex) {
  document.querySelectorAll('.tab-content').forEach(tab => tab.style.display = 'none');
  document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));

  document.getElementById(`tab${tabIndex}`).style.display = 'block';
  document.querySelectorAll('.tab-button')[tabIndex].classList.add('active');
}

// Load everything
window.onload = () => {
  loadExcel('sheet1.xlsx', 'dataTable1');
  loadExcel('sheet2.xlsx', 'dataTable2');
  openTab(0);
};
