async function loadExcel(filename, containerId) {
  try {
    const response = await fetch(filename);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    if (jsonData.length === 0) {
      document.getElementById(containerId).innerHTML = "<p>No data found.</p>";
      return;
    }

    let html = '<table><thead><tr>';
    const headers = Object.keys(jsonData[0]);

    headers.forEach(header => {
      html += `<th>${header}</th>`;
    });
    html += '</tr></thead><tbody>';

    jsonData.forEach(row => {
      html += '<tr>';
      headers.forEach(header => {
        html += `<td>${row[header] ?? ''}</td>`;
      });
      html += '</tr>';
    });

    html += '</tbody></table>';
    document.getElementById(containerId).innerHTML = html;

  } catch (error) {
    console.error(error);
    document.getElementById(containerId).innerHTML = 
      `<p style="color:red;">Error loading ${filename}</p>`;
  }
}

// Tab functionality
function openTab(tabIndex) {
  document.querySelectorAll('.tab-content').forEach(tab => tab.style.display = 'none');
  document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));

  document.getElementById(`tab${tabIndex}`).style.display = 'block';
  document.querySelectorAll('.tab-button')[tabIndex].classList.add('active');
}

// Load both sheets when page loads
window.onload = () => {
  loadExcel('sheet1.xlsx', 'table1');
  loadExcel('sheet2.xlsx', 'table2');
  openTab(0); // Show first tab by default
};
