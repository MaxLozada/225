document.addEventListener('DOMContentLoaded', function() {
    // Load the Excel file automatically
    loadExcelFile('decoded_messages_25.xlsx');

    // Set up event listeners
    document.getElementById('toggleAllBtn').addEventListener('click', toggleAllColumns);
    document.getElementById('exportBtn').addEventListener('click', exportToExcel);
    document.getElementById('searchInput').addEventListener('input', searchTable);
});

async function loadExcelFile(filePath) {
    try {
        // Encode the filename to handle special characters like $
        const encodedPath = encodeURIComponent(filePath);
        const response = await fetch(encodedPath);

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const arrayBuffer = await response.arrayBuffer();
        processExcelFile(arrayBuffer);
    } catch (error) {
        console.error('Error loading Excel file:', error);
        alert(`Failed to load ${filePath}. Please ensure:
1. The file exists in the same directory
2. You're using a local server (not file://)
3. The file isn't open in Excel`);
    }
}

function processExcelFile(data) {
    const workbook = XLSX.read(new Uint8Array(data), { type: 'array' });

    // Get first sheet
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const html = XLSX.utils.sheet_to_html(firstSheet);

    // Display the table
    const tableWrapper = document.getElementById('tableWrapper');
    tableWrapper.innerHTML = html;

    // Add some classes to the table for styling
    const table = tableWrapper.querySelector('table');
    if (table) {
        table.classList.add('data-table');

        // Update status bar
        updateStatusBar(table);

        // Set up column toggles
        setupColumnToggles(table);
    }
}

function setupColumnToggles(table) {
    const headerRow = table.rows[0];
    const columnControls = document.getElementById('columnControls');
    columnControls.innerHTML = '';

    for (let i = 0; i < headerRow.cells.length; i++) {
        const th = headerRow.cells[i];
        const btn = document.createElement('button');
        btn.className = 'toggle-btn active';
        btn.innerHTML = `<i class="fas fa-eye"></i> ${th.textContent || `Column ${i+1}`}`;
        btn.dataset.column = i;

        btn.addEventListener('click', function() {
            this.classList.toggle('active');
            const icon = this.querySelector('i');
            icon.className = this.classList.contains('active') ? 'fas fa-eye' : 'fas fa-eye-slash';
            toggleColumnVisibility(this.dataset.column, this.classList.contains('active'));
        });

        columnControls.appendChild(btn);
    }
}

function toggleColumnVisibility(colIndex, show) {
    const table = document.querySelector('table');
    if (!table) return;

    const rows = table.rows;
    for (let j = 0; j < rows.length; j++) {
        const cell = rows[j].cells[colIndex];
        if (cell) {
            cell.style.display = show ? '' : 'none';
        }
    }
}

function toggleAllColumns() {
    const toggleButtons = document.querySelectorAll('.toggle-btn');
    const allActive = Array.from(toggleButtons).every(btn => btn.classList.contains('active'));

    toggleButtons.forEach(btn => {
        const shouldActivate = !allActive;
        btn.classList.toggle('active', shouldActivate);
        const icon = btn.querySelector('i');
        icon.className = shouldActivate ? 'fas fa-eye' : 'fas fa-eye-slash';
        toggleColumnVisibility(btn.dataset.column, shouldActivate);
    });
}

function searchTable() {
    const input = document.getElementById('searchInput');
    const filter = input.value.toLowerCase();
    const table = document.querySelector('table');
    if (!table) return;

    const rows = table.getElementsByTagName('tr');

    for (let i = 1; i < rows.length; i++) { // Skip header row
        let row = rows[i];
        let cells = row.getElementsByTagName('td');
        let rowMatches = false;

        for (let j = 0; j < cells.length; j++) {
            if (cells[j].style.display !== 'none') { // Only search visible columns
                const txtValue = cells[j].textContent || cells[j].innerText;
                if (txtValue.toLowerCase().indexOf(filter) > -1) {
                    rowMatches = true;
                    break;
                }
            }
        }

        row.style.display = rowMatches ? '' : 'none';
    }
}

function updateStatusBar(table) {
    const rowCount = table.rows.length - 1; // Subtract header row
    const columnCount = table.rows[0].cells.length;

    document.getElementById('rowCount').textContent = `${rowCount} rows loaded`;
    document.getElementById('columnCount').textContent = `${columnCount} columns`;
}

function exportToExcel() {
    const table = document.querySelector('table');
    if (!table) return;

    const workbook = XLSX.utils.table_to_book(table);
    XLSX.writeFile(workbook, 'exported_data.xlsx');
}