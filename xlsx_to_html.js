let xlsx = require('xlsx');
let fs = require('fs');
let inputfile = "Output.xlsx";
let outputfile = "htmltable.html";

// Configuration for your specific table layout
const roomcolumn = 0;
const monthcolumn = 2;
const datecolumn = 3;

let style = `
<style>
    .table-container {
        width: 100%;
        overflow-x: auto;
    }
    h1{
        margin: 0px;
        text-align: left;
        font-size: 30pt;
        color: #003552;
        font-family: Arial, sans-serif;
        background-color: white;
    }
    body {
        background-color: #009dd1;
        margin: 0px;
    }
    table {
        background-color: #e0e0e0;
        border-collapse: collapse; 
        min-width: 800px;
        width: 100%;
        font-family: Arial, sans-serif;
        table-layout: auto;
        margin: 8px;
    }
    th, td {
        background-color: #e0e0e0;
        border: 1px solid #000000; 
        padding: 10px;
        text-align: left;
        color: #003552;
        word-wrap: break-word;
        overflow: hidden;
        width: auto;
    }
    table tr:first-child td {
        font-weight: bold;
        color: black;
    }
    td:first-child {
        font-weight: bold;
        color: black;
        width: fit-content;
    }
    footer {
        margin: 0px;
        text-align: left;
        font-size: 10pt;
        color: #003552;
        font-family: Arial, sans-serif;
        background-color: white;
    }
</style>
`;

let script = `
<script>
    document.addEventListener('DOMContentLoaded', () => {
        const dropdown = document.getElementById('RoomID_dropdown');
        const table = document.querySelector('table');
        const roomcolumn = ${roomcolumn};

        const rows = Array.from(table.querySelectorAll('tr')).slice(1);

        const applyFilter = (selectedroomID) => {
            rows.forEach(row => {
                const roomcellID = row.cells[roomcolumn]; 
                if (roomcellID) {
                    let rowText = roomcellID.textContent.trim();
                    // Strip "Room: " prefix if it exists in the table cell
                    if (rowText.startsWith("Room: ")) {
                        rowText = rowText.substring(6);
                    }
                    
                    if (selectedroomID === "" || rowText === selectedroomID) {
                        row.style.display = ""; 
                    } else {
                        row.style.display = "none"; 
                    }
                }
            });
        };

        dropdown.addEventListener('change', (event) => {
            applyFilter(event.target.value);
        });
    });
</script>
`;

try {
    let wb = xlsx.readFile(inputfile);
    let sheetName = wb.SheetNames[0];
    let ws = wb.Sheets[sheetName];

    // --- NEW: Filter and Sort Logic ---
    let data = xlsx.utils.sheet_to_json(ws, { header: 1 });
    let header = data[0];
    let rows = data.slice(1);

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const months = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];

    let processedRows = rows.map(r => {
        let mIdx = months.indexOf(String(r[monthcolumn]).toLowerCase().trim());
        let dNum = parseInt(String(r[datecolumn]).replace(/\D/g, '')) || 0;
        let rDate = new Date(today.getFullYear(), mIdx, dNum);
        
        // Year rollover check (if today is Dec and reservation is Jan)
        if (mIdx < today.getMonth() && today.getMonth() > 9) rDate.setFullYear(today.getFullYear() + 1);
        return { r, rDate };
    })
    .filter(item => item.rDate >= today && !isNaN(item.rDate.getTime()))
    .sort((a, b) => a.rDate - b.rDate)
    .map(item => item.r);

    // Update the sheet with sorted/filtered data before converting to HTML
    let sortedWs = xlsx.utils.aoa_to_sheet([header, ...processedRows]);
    let rawTableHtml = xlsx.utils.sheet_to_html(sortedWs);
    // ----------------------------------

    const uniqueroomIDs = new Set();
    for (let i = 0; i < processedRows.length; i++) {
        let rawroomID = processedRows[i][roomcolumn];
        if (rawroomID) {
            rawroomID = String(rawroomID).trim();
            if (rawroomID.startsWith("Room: ")) {
                rawroomID = rawroomID.substring("Room: ".length);
            }
            uniqueroomIDs.add(rawroomID);
        }
    }

    let dropdownOptions = '<option value="">Show All Rooms</option>';
    // Added .sort() here so the dropdown list is alphabetical
    Array.from(uniqueroomIDs).sort().forEach(roomID => {
        dropdownOptions += `<option value="${roomID}">${roomID}</option>`;
    });

    const dropdownHtml = `
        <div style="margin-bottom: 20px; margin-top: 20px; font-family: Arial, sans-serif; font-weight: bold; color: white;">
            <label for="RoomID_dropdown">⠀Filter by Room:</label>
            <select id="RoomID_dropdown">
                ${dropdownOptions}
            </select>
        </div>
    `;

    let finalHtmlData = `
<!DOCTYPE html>
<html lang="en"> 
<head>
    <meta charset="UTF-8">
    <title>25Live Room Reservations</title>
    <h1>⠀<a href="https://dsu.edu" style="text-decoration:none"> <img src="DSU_UniversityLogo_Icon_Primary.png" height="30px" alt="DSU Logo"> </a>⠀DSU Campus Reservations</h1>
    ${style}
</head>
<body>
    ${dropdownHtml}
    <div class="table-container">${rawTableHtml}</div>
    ${script}
</body>
<footer>
    <h1>⠀<a href="https://25live.collegenet.com/pro/sdbor" style="text-decoration:none">25Live</a></h1>
</footer>
</html>
`;

    fs.writeFileSync(outputfile, finalHtmlData);
    console.log("Outputted File: " + processedRows.length + " rows processed.");
} catch (e) {
    console.error("Error", e.message);
}