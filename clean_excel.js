const ExcelJS = require('exceljs');

async function cleanSpreadsheet() {
    const wb = new ExcelJS.Workbook();
    try {
        await wb.xlsx.readFile('Parent_Orientation_Tracker.xlsx');
        
        let ws1 = wb.getWorksheet('PMO - Plan Sessions') || wb.getWorksheet('PMO – Plan Sessions') || wb.getWorksheet(1);
        if (ws1) {
            console.log("Cleaning PMO sheet...");
            for (let r = 4; r <= 1004; r++) {
                const row = ws1.getRow(r);
                row.values = [];
                row.commit();
            }
        }
        
        let ws2 = wb.getWorksheet('Speaker – Post Activity') || wb.getWorksheet('Speaker - Post Activity');
        if (ws2) {
            console.log("Cleaning Speaker sheet...");
            for (let r = 5; r <= 1005; r++) {
                const row = ws2.getRow(r);
                row.values = [];
                row.commit();
            }
        }
        
        await wb.xlsx.writeFile('Parent_Orientation_Tracker.xlsx');
        console.log("Successfully cleaned Parent_Orientation_Tracker.xlsx!");
    } catch (err) {
        console.error("Error cleaning file:", err);
    }
}

cleanSpreadsheet();
