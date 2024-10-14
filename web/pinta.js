///////////////////////////////////////////////////
// Put a sheet URL here to fetch the table from it!
///////////////////////////////////////////////////

const SHEET_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRQXByO5kM3joHaiWUJQmkt0oI_gfPp7yZNyGRn9Cg_Khink8NuBtQS9BwtX8LF4Fdt06C2-1HvjRKw/pubhtml?gid=500526275&single=true';

///////asd/////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

async function fetchAndParseTable() {
    try {
        const response = await fetch(SHEET_URL);
        const html = await response.text();
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const tableBody = doc.querySelector('tbody');

        if (!tableBody) {
            console.error('No s\'ha trobat el tbody a la pàgina');
            return;
        }
        return tableBody;
    } catch (error) {
        console.error('Error en fer scraping o processar la taula:', error);
    }
}

function getWeekStatus() {
    const currentDate = new Date();
    const startDate = new Date('2024-09-30');
    const weeksPassed = Math.floor((currentDate - startDate) / (7 * 24 * 60 * 60 * 1000));
    return weeksPassed % 2 === 0 ? 'S2' : 'S1';
}

function updateWeekStatus() {
    const weekElement = document.querySelector('.setmana');
    if (weekElement) {
        weekElement.textContent = getWeekStatus();
    }
}

function getTodayColumnIndex() {
    const columnMapping = {
        1: 0,  // Monday
        2: 6,  // Tuesday
        3: 12, // Wednesday
        4: 18, // Thursday
        5: 24  // Friday
    };
    const dayOfWeek = new Date().getDay();
    return columnMapping[dayOfWeek] !== undefined ? columnMapping[dayOfWeek] : -1;
}

function transposeMatrix(matrix) {
    return matrix[0].map((_, colIndex) => matrix.map(row => row[colIndex]));
}

function updateDay() {
    const dayElement = document.querySelector('.dia');
    if (dayElement) {
        dayElement.textContent = new Date().toLocaleDateString('ca-ES', { weekday: 'long' }).toUpperCase();
    }
}

function updateDate() {
    const today = new Date();
    const formattedDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
    const dateElement = document.querySelector('.data');
    if (dateElement) {
        dateElement.textContent = formattedDate; 
    }
}

async function updateTable() {
    const todayColumnIndex = getTodayColumnIndex();
    if (todayColumnIndex === -1) {
        console.error('No es un dia de dilluns a divendres');
        return;
    }

    const updatedTable = await fetchAndParseTable();
    const rows = updatedTable.querySelectorAll('tr');
    const colorMatrix = Array.from(rows).map(row => {
        const cells = row.querySelectorAll('td');
        return Array.from({ length: 6 }, (_, cellIndex) => {
            const cell = cells[todayColumnIndex + cellIndex];
            if (!cell) return '#c47979'; // Default en vermell

            if ((cell.classList.contains('s1') && getWeekStatus() === 'S1') ||
                (cell.classList.contains('s3') && getWeekStatus() === 'S2') ||
                cell.classList.contains('s2')) {
                return '#93c47d'; // Verd
            }
            return '#c47979'; // Vermell
        });
    });

    const table = document.getElementById('horari');
    const transposedMatrix = transposeMatrix(colorMatrix);

    transposedMatrix.forEach((row, rowIndex) => {
        row.forEach((color, cellIndex) => {
            const cell = table.rows[(rowIndex % 6) + 1].cells[cellIndex + 1];
            cell.style.backgroundColor = color;
        });
    });

    updateWeekStatus();
    updateDay();
    updateDate();
}

updateTable();
