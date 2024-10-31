let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0]; // Load the first sheet
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        // Initially display the full sheet
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = ''; // Clear existing content

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    // Create table headers
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null || cell === "" ? 'NULL' : cell; // Print 'NULL' for empty cells
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Function to apply the selected operations and update the table
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const functionSelect = document.getElementById('function-select').value;

    if (!primaryColumn || !functionSelect) {
        alert('Please select the primary column and a function.');
        return;
    }

    // TODO: Implement operations based on selected function
    // For example:
    if (functionSelect === 'sum') {
        // Logic for sum operation
    }

    // Update the displayed table
    displaySheet(filteredData); // Update with new filtered data
}

// Function to open the download modal
function openDownloadModal() {
    document.getElementById('download-modal').style.display = 'flex';
}

// Function to close the download modal
function closeDownloadModal() {
    document.getElementById('download-modal').style.display = 'none';
}

// Function to download filtered data as an Excel file or CSV
function downloadExcel() {
    const filename = document.getElementById('filename').value.trim() || 'download';
    const format = document.getElementById('file-format').value;

    const exportData = filteredData.map(row => {
        return Object.keys(row).reduce((acc, key) => {
            acc[key] = row[key] === null || row[key] === "" ? 'NULL' : row[key]; // Ensure 'NULL' for empty cells
            return acc;
        }, {});
    });

    let worksheet = XLSX.utils.json_to_sheet(exportData);
    let workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Filtered Data');

    if (format === 'xlsx') {
        XLSX.writeFile(workbook, `${filename}.xlsx`);
    } else if (format === 'csv') {
        XLSX.writeFile(workbook, `${filename}.csv`, { bookType: 'csv' });
    }

    closeDownloadModal();
}

// Event Listeners
document.getElementById('function-select').addEventListener('change', function() {
    const selectedFunction = this.value;
    const operationOutput = document.getElementById('operation-type');

    switch (selectedFunction) {
        case 'sum':
            operationOutput.textContent = 'Calculates the total of the selected column.';
            break;
        case 'subtraction':
            operationOutput.textContent = 'Subtracts values from the selected column.';
            break;
        case 'multiplication':
            operationOutput.textContent = 'Multiplies values in the selected column.';
            break;
        case 'division':
            operationOutput.textContent = 'Divides values in the selected column.';
            break;
        case 'average':
            operationOutput.textContent = 'Calculates the average of the selected column.';
            break;
        case 'comparison':
            operationOutput.textContent = 'Compares values in the selected column.';
            break;
        default:
            operationOutput.textContent = 'Please select a function.';
            break;
    }
});

document.getElementById('apply-operation').addEventListener('click', applyOperation);
document.getElementById('download-button').addEventListener('click', openDownloadModal);
document.getElementById('confirm-download').addEventListener('click', downloadExcel);
document.getElementById('close-modal').addEventListener('click', closeDownloadModal);

// Call loadExcelSheet with the provided Google Sheet URL
loadExcelSheet('URL_OF_YOUR_GOOGLE_SHEET'); // Replace with actual URL
