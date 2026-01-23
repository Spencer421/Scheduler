let xlsx = require('xlsx');
let fs = require('fs');
let inputfile = "Output.xlsx";
let outputfile = "htmltable.html";
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
</style>
`;

let script = `
<script>
    document.addEventListener('DOMContentLoaded', () => {
        const dropdown = document.getElementById('RoomID_dropdown');
        const table = document.querySelector('table');
        const roomcolumn = 0;

        const rows = Array.from(table.querySelectorAll('tr')).slice(1);

        const urlParams = new URLSearchParams(window.location.search);
        const urlRoomID = urlParams.get('roomID');
        let initialFilterRoomID = "";

        if (urlRoomID) {
            const optionExists = Array.from(dropdown.options).some(
                option => option.value === urlRoomID
            );
            if (optionExists) {
                dropdown.value = urlRoomID;
                initialFilterRoomID = urlRoomID;
            }
        }
        
        const applyFilter = (selectedroomID) => {
            rows.forEach(row => {
                const roomcellID = row.cells[roomcolumn]; 
                if (roomcellID) {
                    const rowText = roomcellID.textContent.trim();
                    
                    if (selectedroomID === "") {
                        row.style.display = ""; 
                    } else if (rowText === selectedroomID) {
                        row.style.display = ""; 
                    } else {
                        row.style.display = "none"; 
                    }
                }
            });
        };

        applyFilter(initialFilterRoomID);

        dropdown.addEventListener('change', (event) => {
            const selectedroomID = event.target.value;
            applyFilter(selectedroomID);
        });
    });
</script>
`;

try {
    let wb = xlsx.readFile(inputfile);
    let sheetName = wb.SheetNames[0];
    let ws = wb.Sheets[sheetName];

    let rawTableHtml = xlsx.utils.sheet_to_html(ws);
    let data = xlsx.utils.sheet_to_json(ws, { header: 1 });
const uniqueroomIDs = new Set();
const roomcolumn = 0;
    for (let i = 1; i < data.length; i++) {
        let rawroomID = data[i][roomcolumn];

        if (rawroomID) {
            if (rawroomID.startsWith("Room: ")) {
                rawroomID = rawroomID.substring("Room: ".length);
            }
            uniqueroomIDs.add(String(rawroomID).trim());
        }
    }
    let dropdownOptions = '<option value="">Show All Rooms</option>';
    uniqueroomIDs.forEach(roomID => {
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
    <h1>⠀<img src="DSU_UniversityLogo_Icon_Primary.png" height="30px" alt="DSU Logo">⠀DSU Campus Reservations</h1>
    ${style}
</head>
<body>
    ${dropdownHtml}
    <div class="table-container">${rawTableHtml}</div>
    ${script}
</body>
</html>
`;

    fs.writeFileSync(outputfile, finalHtmlData);
    console.log("Outputted File");
} catch (e) {
    console.error("Error", e.message);
}