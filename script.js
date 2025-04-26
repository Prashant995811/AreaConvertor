let measurements = [];

function addMeasurement() {
  document.getElementById('manual-entry').classList.remove('hidden');
}

function cancelMeasurement() {
  document.getElementById('manual-entry').classList.add('hidden');
  clearManualForm();
}

function clearManualForm() {
  document.getElementById('description').value = '';
  document.getElementById('length').value = '';
  document.getElementById('breadth').value = '';
  document.getElementById('unit').value = 'inch';
}

function saveMeasurement() {
  const description = document.getElementById('description').value || '';
  const length = parseFloat(document.getElementById('length').value);
  const breadth = parseFloat(document.getElementById('breadth').value);
  const unit = document.getElementById('unit').value;

  if (isNaN(length) || isNaN(breadth)) {
    alert("Please enter valid numbers for Length and Breadth.");
    return;
  }

  const lengthInInch = convertToInch(length, unit);
  const breadthInInch = convertToInch(breadth, unit);
  const areaSqft = (lengthInInch * breadthInInch) / 144;

  measurements.push({
    description,
    length: lengthInInch.toFixed(2),
    breadth: breadthInInch.toFixed(2),
    area: areaSqft.toFixed(2)
  });

  clearManualForm();
  cancelMeasurement();
  renderMeasurements();
}

function convertToInch(value, unit) {
  switch (unit) {
    case 'cm':
      return value / 2.54;
    case 'mm':
      return value / 25.4;
    case 'inch':
    default:
      return value;
  }
}

function renderMeasurements() {
  const tbody = document.getElementById('table-body');
  tbody.innerHTML = '';
  let totalArea = 0;

  measurements.forEach((m, index) => {
    totalArea += parseFloat(m.area);
    const row = `<tr>
      <td class="border px-4 py-2">${index + 1}</td>
      <td class="border px-4 py-2">${m.description}</td>
      <td class="border px-4 py-2">${m.length}</td>
      <td class="border px-4 py-2">${m.breadth}</td>
      <td class="border px-4 py-2">${m.area}</td>
      <td class="border px-4 py-2">
        <button onclick="deleteMeasurement(${index})" class="bg-red-500 hover:bg-red-600 text-white px-3 py-1 rounded">
          🗑️
        </button>
      </td>
    </tr>`;
    tbody.innerHTML += row;
  });

  document.getElementById('total-area').textContent = `Total Area: ${totalArea.toFixed(2)} sqft`;
}

function deleteMeasurement(index) {
  measurements.splice(index, 1);
  renderMeasurements();
}

function downloadExcel() {
  const filename = document.getElementById('filename').value.trim() || 'Measurements';

  const data = measurements.map((m, i) => ({
    'Serial No.': i + 1,
    'Description': m.description,
    'Length': m.length,
    'Breadth': m.breadth,
    'Area (sqft)': m.area
  }));

  const totalArea = measurements.reduce((sum, m) => sum + parseFloat(m.area), 0);

  // Add Total row
  data.push({
    'Serial No.': '',
    'Description': '',
    'Length': '',
    'Breadth': 'Total',
    'Area (sqft)': totalArea.toFixed(2)
  });

  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Measurements");

  XLSX.writeFile(workbook, `${filename}.xlsx`);
}

// Upload Excel and Populate Table
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet);

    // Clear existing measurements
    measurements = [];

    json.forEach(row => {
      // Skip Total row
      if (row['Breadth'] === 'Total') return;

      measurements.push({
        description: row['Description'] || '',
        length: parseFloat(row['Length']).toFixed(2),
        breadth: parseFloat(row['Breadth']).toFixed(2),
        area: parseFloat(row['Area (sqft)']).toFixed(2)
      });
    });

    renderMeasurements();
  };
  reader.readAsArrayBuffer(file);
}
