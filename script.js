// script.js

let results = [];

function addMeasurement() {
  const measurementsDiv = document.getElementById('measurements');

  const newRow = document.createElement('div');
  newRow.className = 'measurement-row';

  newRow.innerHTML = `
    <input type="text" placeholder="Description" class="description border p-2 rounded-lg">
    <input type="number" placeholder="Length" class="length border p-2 rounded-lg">
    <select class="length-unit border p-2 rounded-lg">
      <option value="mm">mm</option>
      <option value="cm">cm</option>
      <option value="inch">inch</option>
    </select>
    <input type="number" placeholder="Breadth" class="breadth border p-2 rounded-lg">
    <select class="breadth-unit border p-2 rounded-lg">
      <option value="mm">mm</option>
      <option value="cm">cm</option>
      <option value="inch">inch</option>
    </select>
    <button onclick="removeMeasurement(this)" class="text-red-500 hover:text-red-700 font-bold">🗑️</button>
  `;

  measurementsDiv.appendChild(newRow);
  attachAutoCalculate();
}

function removeMeasurement(button) {
  button.parentElement.remove();
  calculateAreas();
}

function convertToInches(value, unit) {
  if (unit === 'mm') return value / 25.4;
  if (unit === 'cm') return value / 2.54;
  return value; // already inches
}

function calculateAreas() {
  const descriptions = document.querySelectorAll('.description');
  const lengths = document.querySelectorAll('.length');
  const lengthUnits = document.querySelectorAll('.length-unit');
  const breadths = document.querySelectorAll('.breadth');
  const breadthUnits = document.querySelectorAll('.breadth-unit');

  results = [];
  let resultsList = "";
  let totalArea = 0;

  for (let i = 0; i < lengths.length; i++) {
    const desc = descriptions[i].value.trim();
    const lengthVal = parseFloat(lengths[i].value);
    const lengthUnitVal = lengthUnits[i].value;
    const breadthVal = parseFloat(breadths[i].value);
    const breadthUnitVal = breadthUnits[i].value;

    if (isNaN(lengthVal) || isNaN(breadthVal)) continue;

    const lengthInInch = convertToInches(lengthVal, lengthUnitVal);
    const breadthInInch = convertToInches(breadthVal, breadthUnitVal);

    const areaSqInch = lengthInInch * breadthInInch;
    const areaSqFt = areaSqInch / 144;

    totalArea += areaSqFt;

    results.push({
      serial: i + 1,
      description: desc,
      length: lengthVal + ' ' + lengthUnitVal,
      breadth: breadthVal + ' ' + breadthUnitVal,
      area: areaSqFt.toFixed(2)
    });

    resultsList += `<li>${desc ? desc : "No Description"}: <span class="font-semibold">${areaSqFt.toFixed(2)} sqft</span></li>`;
  }

  document.getElementById('results-list').innerHTML = resultsList;
  document.getElementById('total-area').innerHTML = `Total Area: ${totalArea.toFixed(2)} sqft`;
}

function attachAutoCalculate() {
  const inputs = document.querySelectorAll('.length, .breadth, .description, .length-unit, .breadth-unit');
  inputs.forEach(input => {
    input.addEventListener('input', calculateAreas);
    input.addEventListener('change', calculateAreas);
  });
}

function downloadExcel() {
    if (results.length === 0) {
      alert('Please add some measurements first!');
      return;
    }
  
    // Get the file name from user input
    const filenameInput = document.getElementById('filename').value.trim();
    const fileName = filenameInput !== "" ? filenameInput + ".xlsx" : "measurements.xlsx";
  
    const wsData = [
      ["Serial No.", "Description", "Length", "Breadth", "Area (sqft)"]
    ];
  
    // Adding rows and data
    results.forEach(r => {
      wsData.push([r.serial, r.description, r.length, r.breadth, r.area]);
    });
  
    // Add total area row
    const total = results.reduce((sum, r) => sum + parseFloat(r.area), 0);
    wsData.push(["", "Total Area", "", "", total.toFixed(2)]);
  
    // Create workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);
  
    // Apply formatting (borders, bold header row, etc.)
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const cell = ws[XLSX.utils.encode_cell({r: R, c: C})];
        
        // Make the first row bold (headers)
        if (R === 0) {
          if (cell) {
            cell.s = {font: {bold: true}};
          }
        }
        
        // Add borders to each cell
        if (cell) {
          if (!cell.s) {
            cell.s = {}; // Ensure the cell has a 'style' property
          }
          cell.s.border = {
            top: {style: 'thin', color: {rgb: "000000"}},
            left: {style: 'thin', color: {rgb: "000000"}},
            bottom: {style: 'thin', color: {rgb: "000000"}},
            right: {style: 'thin', color: {rgb: "000000"}}
          };
        }
      }
    }
  
    // Conditional formatting: Highlight rows with area > 100 sqft (optional)
    wsData.forEach((row, index) => {
      if (index > 0) {  // Skip header row
        const areaCell = ws[XLSX.utils.encode_cell({r: index, c: 4})];
        if (areaCell && parseFloat(areaCell.v) > 100) {
          areaCell.s = {
            fill: { 
              fgColor: { rgb: "FFFF00" } // Yellow fill for large areas
            }
          };
        }
      }
    });
  
    // Set dynamic sheet name based on first description or today's date
    const sheetName = results[0]?.description || new Date().toLocaleDateString();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  
    // Write the Excel file
    XLSX.writeFile(wb, fileName);
  
    alert("✅ Excel file downloaded successfully with enhanced formatting!");
  }
  

// Initialize with one measurement
addMeasurement();
