function main(workbook: ExcelScript.Workbook) {
    // Specify the sheet names
    const sheetNames = ["M2NELab", "M2NW Hallway (Front of 60) -->", "Sheet1", "Sheet2", "M2NE Hallway (Front of 360) -->"];

    // Iterate over the sheets
    for (let sheetName of sheetNames) {
        // Get the worksheet
        let worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            console.log(`Sheet: ${sheetName} not found`);
            continue;
        }

        // Set the first row color to dark blue and font to white bold of size 18
        let firstRow = worksheet.getRange("1:1");
        firstRow.getFormat().getFill().setColor("002060"); // dark blue
        firstRow.getFormat().getFont().setColor("FFFFFF"); // white
        firstRow.getFormat().getFont().setBold(true); // bold
        firstRow.getFormat().getFont().setSize(18); // font size 18

        // Get the used range
        let usedRange = worksheet.getUsedRange();
        if (!usedRange) {
            console.log(`No data in sheet: ${sheetName}`);
            continue;
        }

        // Get the number of rows in the used range
        let rowCount = usedRange.getRowCount();

        // Start from the second row, and alternate the fill color between white and light turquoise
        for (let i = 2; i <= rowCount; i++) {
            let row = worksheet.getRange(`${i}:${i}`);
            let color = (i % 2 == 0) ? "FFFFFF" : "AFEEEE"; // white and light turquoise
            row.getFormat().getFill().setColor(color);
        }
    }
}
