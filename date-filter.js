(function () {
    const MONTH_COL = 2;
    const DATE_COL  = 3;
 
    const MONTHS = [
        "january", "february", "march", "april", "may", "june",
        "july", "august", "september", "october", "november", "december"
    ];
 
    function parseReservationDate(monthText, dateText) {
        const mIdx = MONTHS.indexOf(String(monthText).toLowerCase().trim());
        const dNum = parseInt(String(dateText).replace(/\D/g, ""), 10);
 
        if (mIdx === -1 || isNaN(dNum)) return null;
 
        //const today = new Date();
        const today = new Date();
        today.setDate(today.getDate() + 1); // TEST 

        let year = today.getFullYear();
 
        if (mIdx < today.getMonth()) year += 1;
 
        return new Date(year, mIdx, dNum);
    }
 
    function filterPastRows() {
        const table = document.querySelector("table");
        if (!table) return;
 
        const today = new Date();
        today.setHours(0, 0, 0, 0); 

        const rows = Array.from(table.querySelectorAll("tr")).slice(1);
        let hiddenCount = 0;
 
        rows.forEach(row => {
            const cells = row.cells;
            if (!cells || cells.length <= DATE_COL) return;
 
            const monthText = cells[MONTH_COL] ? cells[MONTH_COL].textContent.trim() : "";
            const dateText  = cells[DATE_COL]  ? cells[DATE_COL].textContent.trim()  : "";
 
            const reservationDate = parseReservationDate(monthText, dateText);
 
            if (reservationDate && reservationDate < today) {
                row.style.display = "none";
                hiddenCount++;
            }
        });
 
        if (hiddenCount > 0) {
            console.log(`[date-filter] Hid ${hiddenCount} past reservation(s).`);
        }
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", filterPastRows);
    } else {
        filterPastRows(); // Shawn Zwach is the GOAT
    }
})();