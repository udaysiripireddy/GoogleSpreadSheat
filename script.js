let rows = 10;
let cols = 12;
let selectedCell = null;
const sheetBody = document.getElementById("sheet-body");
const headerRow = document.getElementById("header-row");

// Generate initial spreadsheet
function generateSheet() {
    headerRow.innerHTML = "<th></th>"; // Empty corner cell

    for (let i = 0; i < cols; i++) {
        let th = document.createElement("th");
        th.innerText = columnName(i);
        headerRow.appendChild(th);
    }

    sheetBody.innerHTML = "";
    for (let i = 0; i < rows; i++) {
        let tr = document.createElement("tr");

        let th = document.createElement("th");
        th.innerText = i + 1;
        tr.appendChild(th);

        for (let j = 0; j < cols; j++) {
            let td = createEditableCell(i, j);
            tr.appendChild(td);
        }

        sheetBody.appendChild(tr);
    }
}

// Create an editable cell
function createEditableCell(row, col) {
    let td = document.createElement("td");
    td.setAttribute("contenteditable", "true");
    td.dataset.row = row;
    td.dataset.col = col;
    
    td.addEventListener("click", () => selectCell(td));
    return td;
}

// Select a cell
function selectCell(cell) {
    if (selectedCell) {
        selectedCell.style.backgroundColor = "";
    }
    selectedCell = cell;
    selectedCell.style.backgroundColor = "#ffeb3b"; 
}

// Apply formatting
function applyFormatting(action) {
    if (!selectedCell) {
        alert("Please select a cell first.");
        return;
    }

    switch (action) {
        case "BOLD":
            selectedCell.style.fontWeight = selectedCell.style.fontWeight === "bold" ? "normal" : "bold";
            break;
        case "ITALIC":
            selectedCell.style.fontStyle = selectedCell.style.fontStyle === "italic" ? "normal" : "italic";
            break;
        case "COLOR":
            let color = prompt("Enter a color (e.g., red, blue, #ff0000):");
            if (color) selectedCell.style.color = color;
            break;
    }
}

// Perform calculations on a selected row or column
function calculate(operation) {
    if (!selectedCell) {
        alert("Please select a cell to perform calculations.");
        return;
    }

    let selectedRow = selectedCell.dataset.row;
    let selectedColumn = selectedCell.dataset.col;
    let values = [];

    let isRowSelected = confirm("Do you want to calculate based on the selected ROW? Click 'Cancel' for COLUMN calculation.");

    if (isRowSelected) {
        document.querySelectorAll(`[data-row="${selectedRow}"]`).forEach(td => {
            let num = parseFloat(td.innerText);
            if (!isNaN(num)) values.push(num);
        });
    } else {
        document.querySelectorAll(`[data-col="${selectedColumn}"]`).forEach(td => {
            let num = parseFloat(td.innerText);
            if (!isNaN(num)) values.push(num);
        });
    }

    if (values.length === 0) {
        alert(`No numeric values found for ${operation}`);
        return;
    }

    let result;
    switch (operation) {
        case "SUM":
            result = values.reduce((sum, num) => sum + num, 0);
            break;
        case "COUNT":
            result = values.length;
            break;
        case "AVERAGE":
        case "MEAN":
            result = values.reduce((sum, num) => sum + num, 0) / values.length;
            break;
        case "MEDIAN":
            values.sort((a, b) => a - b);
            let mid = Math.floor(values.length / 2);
            result = values.length % 2 === 0 ? (values[mid - 1] + values[mid]) / 2 : values[mid];
            break;
        case "MODE":
            let frequency = {};
            values.forEach(num => frequency[num] = (frequency[num] || 0) + 1);
            let maxFreq = Math.max(...Object.values(frequency));
            result = Object.keys(frequency).filter(num => frequency[num] === maxFreq).join(", ");
            break;
        case "MAX":
            result = Math.max(...values);
            break;
        case "MIN":
            result = Math.min(...values);
            break;
    }

    alert(`${operation}: ${result}`);
}

// Add/Remove Rows & Columns
function addRow() {
    let tr = document.createElement("tr");
    let rowIndex = sheetBody.rows.length;

    let th = document.createElement("th");
    th.innerText = rowIndex + 1;
    tr.appendChild(th);

    for (let i = 0; i < cols; i++) {
        let td = createEditableCell(rowIndex, i);
        tr.appendChild(td);
    }

    sheetBody.appendChild(tr);
    rows++;
}

function addColumn() {
    let th = document.createElement("th");
    th.innerText = columnName(cols);
    headerRow.appendChild(th);

    document.querySelectorAll("#sheet-body tr").forEach((tr, rowIndex) => {
        let td = createEditableCell(rowIndex, cols);
        tr.appendChild(td);
    });

    cols++;
}

function deleteRow() {
    if (rows > 1) {
        sheetBody.removeChild(sheetBody.lastChild);
        rows--;
    }
}

function deleteColumn() {
    if (cols > 1) {
        headerRow.removeChild(headerRow.lastChild);
        document.querySelectorAll("#sheet-body tr").forEach(tr => tr.removeChild(tr.lastChild));
        cols--;
    }
}

// Find and Replace
function findAndReplace() {
    let findText = document.getElementById("find-text").value;
    let replaceText = document.getElementById("replace-text").value;

    if (findText === "") {
        alert("Please enter text to find.");
        return;
    }

    document.querySelectorAll("#sheet-body td").forEach(td => {
        if (td.innerText.includes(findText)) {
            td.innerText = td.innerText.replace(findText, replaceText);
        }
    });
}

// Data Cleaning Functions
function applyDataQuality(action) {
    if (!selectedCell) {
        alert("Please select a cell first.");
        return;
    }

    switch (action) {
        case "TRIM":
            selectedCell.innerText = selectedCell.innerText.trim();
            break;
        case "UPPER":
            selectedCell.innerText = selectedCell.innerText.toUpperCase();
            break;
        case "LOWER":
            selectedCell.innerText = selectedCell.innerText.toLowerCase();
            break;
    }
}

// Remove Duplicates
function removeDuplicates() {
    let selectedColumn = selectedCell ? selectedCell.dataset.col : null;
    if (selectedColumn === null) {
        alert("Please select a column to remove duplicates.");
        return;
    }

    let uniqueValues = new Set();
    document.querySelectorAll(`[data-col="${selectedColumn}"]`).forEach(td => {
        if (uniqueValues.has(td.innerText)) {
            td.innerText = "";
        } else {
            uniqueValues.add(td.innerText);
        }
    });
}

// Column name (A, B, C... AA, AB)
function columnName(index) {
    let result = "";
    while (index >= 0) {
        result = String.fromCharCode((index % 26) + 65) + result;
        index = Math.floor(index / 26) - 1;
    }
    return result;
}
document.addEventListener("DOMContentLoaded", function () {
    const table = document.querySelector("table");
    let selectedCell = null;

    // Track the clicked cell
    table.addEventListener("click", function (event) {
        if (event.target.tagName === "TD") {
            selectedCell = event.target; // Store selected cell
            table.querySelectorAll("td").forEach(td => td.style.outline = ""); // Clear outlines
            selectedCell.style.outline = "2px solid blue"; // Highlight selected cell
        }
    });

    // Handle Paste Event
    document.addEventListener("paste", function (event) {
        if (!selectedCell) {
            alert("Please select a cell before pasting.");
            return;
        }

        event.preventDefault(); // Stop default paste behavior
        
        let clipboardData = event.clipboardData || window.clipboardData;
        let pastedData = clipboardData.getData("Text");

        if (!pastedData) return;

        let rows = pastedData.split("\n").map(row => row.split("\t")); // Split by tabs & new lines
        let startRow = selectedCell.parentElement.rowIndex; // Get selected row index
        let startCol = selectedCell.cellIndex; // Get selected column index

        for (let i = 0; i < rows.length; i++) {
            let rowIndex = startRow + i;
            if (rowIndex >= table.rows.length) break; // Stop if out of table bounds

            let row = table.rows[rowIndex];

            for (let j = 0; j < rows[i].length; j++) {
                let colIndex = startCol + j;
                if (colIndex >= row.cells.length) break; // Stop if out of column bounds

                row.cells[colIndex].textContent = rows[i][j].trim();
            }
        }

        // Clear cell highlight after pasting
        selectedCell.style.outline = "";
        selectedCell = null;
    });
});



// Generate initial spreadsheet
function generateSheet() {
    headerRow.innerHTML = "<th></th>"; // Empty corner cell

    for (let i = 0; i < cols; i++) {
        let th = document.createElement("th");
        th.innerText = columnName(i);
        th.style.position = "relative"; // Needed for resizer
        addColumnResizer(th);
        headerRow.appendChild(th);
    }

    sheetBody.innerHTML = "";
    for (let i = 0; i < rows; i++) {
        let tr = document.createElement("tr");

        let th = document.createElement("th");
        th.innerText = i + 1;
        th.style.position = "relative"; // Needed for resizer
        addRowResizer(th);
        tr.appendChild(th);

        for (let j = 0; j < cols; j++) {
            let td = createEditableCell(i, j);
            tr.appendChild(td);
        }

        sheetBody.appendChild(tr);
    }
}

// Create an editable cell
function createEditableCell(row, col) {
    let td = document.createElement("td");
    td.setAttribute("contenteditable", "true");
    td.dataset.row = row;
    td.dataset.col = col;
    
    td.addEventListener("click", () => selectCell(td));
    return td;
}

// Select a cell
function selectCell(cell) {
    if (selectedCell) {
        selectedCell.style.backgroundColor = "";
    }
    selectedCell = cell;
    selectedCell.style.backgroundColor = "#ffeb3b"; 
}

// Function to add a resizer for columns
function addColumnResizer(th) {
    let resizer = document.createElement("div");
    resizer.style.width = "5px";
    resizer.style.height = "100%";
    resizer.style.position = "absolute";
    resizer.style.right = "0";
    resizer.style.top = "0";
    resizer.style.cursor = "col-resize";
    resizer.addEventListener("mousedown", initColumnResize);
    th.appendChild(resizer);
}

// Function to add a resizer for rows
function addRowResizer(th) {
    let resizer = document.createElement("div");
    resizer.style.width = "100%";
    resizer.style.height = "5px";
    resizer.style.position = "absolute";
    resizer.style.bottom = "0";
    resizer.style.left = "0";
    resizer.style.cursor = "row-resize";
    resizer.addEventListener("mousedown", initRowResize);
    th.appendChild(resizer);
}

// Initialize column resizing
function initColumnResize(e) {
    let th = e.target.parentElement;
    let startX = e.clientX;
    let startWidth = th.offsetWidth;

    function doResize(event) {
        let newWidth = startWidth + (event.clientX - startX);
        th.style.width = newWidth + "px";
    }

    function stopResize() {
        document.removeEventListener("mousemove", doResize);
        document.removeEventListener("mouseup", stopResize);
    }

    document.addEventListener("mousemove", doResize);
    document.addEventListener("mouseup", stopResize);
}

// Initialize row resizing
function initRowResize(e) {
    let th = e.target.parentElement;
    let tr = th.parentElement;
    let startY = e.clientY;
    let startHeight = tr.offsetHeight;

    function doResize(event) {
        let newHeight = startHeight + (event.clientY - startY);
        tr.style.height = newHeight + "px";
    }

    function stopResize() {
        document.removeEventListener("mousemove", doResize);
        document.removeEventListener("mouseup", stopResize);
    }

    document.addEventListener("mousemove", doResize);
    document.addEventListener("mouseup", stopResize);
}

// Column name (A, B, C... AA, AB)
function columnName(index) {
    let result = "";
    while (index >= 0) {
        result = String.fromCharCode((index % 26) + 65) + result;
        index = Math.floor(index / 26) - 1;
    }
    return result;
}
function saveToExcel() {
    let table = document.getElementById("sheet"); // Get the table element
    let wb = XLSX.utils.book_new(); // Create a new workbook
    let ws = XLSX.utils.table_to_sheet(table); // Convert table to worksheet

    XLSX.utils.book_append_sheet(wb, ws, "Sheet1"); // Append worksheet to workbook

    // Trigger the download
    XLSX.writeFile(wb, "Spreadsheet.xlsx");
}

// Initialize spreadsheet
generateSheet();


