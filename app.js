async function loadExcel(filename, containerId) {
  try {
    const response = await fetch(filename);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    let html = '<table><thead><tr>';
    const headers = Object.keys(jsonData[0] || {});
    
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
    console.error(`Error loading ${filename}:`, error);
    document.getElementById(containerId).innerHTML = `<p style="color:red;">Error loading ${filename}</p>`;
  }
}

// Load both Excel files
loadExcel('sheet1.xlsx', 'table1');
loadExcel('sheet2.xlsx', 'table2');
