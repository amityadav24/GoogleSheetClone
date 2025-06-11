let sheetCount = 0;
let cellData = {};
let currentCell = null;

function createSpreadsheet(containerId) {
  const container = document.getElementById(containerId);
  const spreadsheet = document.createElement('div');
  spreadsheet.className = 'spreadsheet';
  spreadsheet.dataset.sheetId = ++sheetCount;

  spreadsheet.appendChild(document.createElement('div')); // Top-left empty corner

  for (let col = 0; col < 20; col++) {
    const colHeader = document.createElement('div');
    colHeader.className = 'header';
    colHeader.textContent = String.fromCharCode(65 + col);
    spreadsheet.appendChild(colHeader);
  }

  for (let row = 1; row <= 50; row++) {
    const rowHeader = document.createElement('div');
    rowHeader.className = 'row-header';
    rowHeader.textContent = row;
    spreadsheet.appendChild(rowHeader);

    for (let col = 0; col < 20; col++) {
      const cell = document.createElement('div');
      cell.className = 'cell';
      cell.contentEditable = true;
      const cellId = `S${sheetCount}_${String.fromCharCode(65 + col)}${row}`;
      cell.dataset.id = cellId;

      cell.addEventListener('click', () => {
        currentCell = cell;
        document.getElementById("formulaInput").value = cell.innerText;
      });

      cell.addEventListener('input', () => {
        const value = cell.innerText;
        cellData[cellId] = {
          value,
          color: cell.style.color,
          bg: cell.style.backgroundColor,
        };
      });

      spreadsheet.appendChild(cell);
    }
  }

  container.appendChild(spreadsheet);
}

function format(command) {
  document.execCommand(command, false, null);
}

function applyColor(color, type) {
  if (currentCell) {
    currentCell.style[type] = color;
    const cellId = currentCell.dataset.id;
    cellData[cellId] = cellData[cellId] || {};
    cellData[cellId][type === 'color' ? 'color' : 'bg'] = color;
  }
}

document.getElementById('fontSelect').addEventListener('change', function () {
  document.execCommand('fontName', false, this.value);
});

document.getElementById('fontSizeSelect').addEventListener('change', function () {
  document.execCommand('fontSize', false, this.value);
});

document.getElementById("formulaInput").addEventListener("keydown", (e) => {
  if (e.key === "Enter" && currentCell) {
    e.preventDefault();
    let input = e.target.value.trim();
    if (input.startsWith("=")) {
      try {
        let result = eval(input.substring(1));
        currentCell.innerText = result;
      } catch {
        currentCell.innerText = "Error";
      }
    } else {
      currentCell.innerText = input;
    }
    currentCell.dispatchEvent(new Event('input')); // update cellData
  }
});

function addNewSheet() {
  createSpreadsheet('sheetContainer');
}

function exportToCSV() {
  const sheet = document.querySelector('.spreadsheet:last-of-type');
  let csv = '';
  let rows = 50;
  let cols = 20;
  for (let row = 1; row <= rows; row++) {
    let rowData = [];
    for (let col = 0; col < cols; col++) {
      let cellId = `S${sheetCount}_${String.fromCharCode(65 + col)}${row}`;
      rowData.push(cellData[cellId]?.value || '');
    }
    csv += rowData.join(',') + '\n';
  }

  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'spreadsheet.csv';
  a.click();
  URL.revokeObjectURL(url);
}

document.getElementById("importFile").addEventListener("change", function () {
  const file = this.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    const rows = e.target.result.split("\n");
    const sheet = document.querySelector('.spreadsheet:last-of-type');
    const cells = sheet.querySelectorAll(".cell");
    let index = 0;
    for (let row of rows) {
      let cols = row.split(",");
      for (let val of cols) {
        if (cells[index]) cells[index].innerText = val.trim();
        index++;
      }
    }
  };
  reader.readAsText(file);
});

function toggleDarkMode() {
  document.body.classList.toggle("dark");
  localStorage.setItem("theme", document.body.classList.contains("dark") ? "dark" : "light");
}

// Load saved theme
window.onload = () => {
  if (localStorage.getItem("theme") === "dark") {
    document.body.classList.add("dark");
  }
  createSpreadsheet("sheetContainer");
};
