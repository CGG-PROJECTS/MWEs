let table1, table2;

async function loadExcel(filename, tableId) {
  try {
    console.log(`Loading ${filename}...`);

    const response = await fetch(filename);
    if (!response.ok) {
      throw new Error(`HTTP ${response.status} while loading ${filename}`);
    }

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
      defval: ""
    });

    if (jsonData.length === 0) {
      console.error(`No data found in ${filename}`);
      return;
    }

    const columns = Object.keys(jsonData[0]).map(key => ({
      title: key,
      data: key
    }));

    if (tableId === 'dataTable1') {
      if (table1) {
        table1.destroy();
        $('#dataTable1').empty();
      }

      table1 = $('#dataTable1').DataTable({
        data: jsonData,
        columns: columns,
        pageLength: 25,
        scrollX: true,
        scrollY: "65vh",
        scrollCollapse: true,
        order: [],
      });

      console.log(`Loaded ${filename} into dataTable1`);
    } else {
      if (table2) {
        table2.destroy();
        $('#dataTable2').empty();
      }

      table2 = $('#dataTable2').DataTable({
        data: jsonData,
        columns: columns,
        pageLength: 25,
        scrollX: true,
        scrollY: "65vh",
        scrollCollapse: true,
        order: [],
      });

      console.log(`Loaded ${filename} into dataTable2`);
    }

  } catch (error) {
    console.error(error);
    alert(`Error loading ${filename}: ${error.message}`);
  }
}

// Tab control
function openTab(tabIndex) {
  document.querySelectorAll('.tab-content').forEach(tab => {
    tab.style.display = 'none';
  });

  document.querySelectorAll('.tab-button').forEach(btn => {
    btn.classList.remove('active');
  });

  document.getElementById(`tab${tabIndex}`).style.display = 'block';
  document.querySelectorAll('.tab-button')[tabIndex].classList.add('active');

  setTimeout(() => {
    if (tabIndex === 0 && table1) {
      table1.columns.adjust().draw(false);
    }

    if (tabIndex === 1 && table2) {
      table2.columns.adjust().draw(false);
    }
  }, 100);
}

// Load everything
window.onload = async () => {
  openTab(0);

  await loadExcel('sheet1.xlsx', 'dataTable1');
  await loadExcel('sheet2.xlsx', 'dataTable2');

  openTab(0);
};
