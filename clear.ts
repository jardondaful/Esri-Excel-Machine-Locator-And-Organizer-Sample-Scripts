function main(workbook: ExcelScript.Workbook) {
  // Get the destination worksheet by name
  let destinationSheet = workbook.getWorksheet("Sheet1");

  // If the destination sheet doesn't exist, create it
  if (!destinationSheet) {
    destinationSheet = workbook.addWorksheet("Sheet1");
  }

  // Clear everything within "Sheet1"
  destinationSheet.getUsedRange()?.clear();

  // Get all worksheets
  let allWorksheets = workbook.getWorksheets();

  // Iterate through all worksheets
  allWorksheets.forEach((worksheet) => {
    // Check if the worksheet is created by sortAndCategorize function
    if (worksheet.getName().startsWith("Category_")) {
      // Clear everything within this worksheet
      worksheet.getUsedRange()?.clear();
    }

    // Set the style of the first row
    let firstRow = worksheet.getRange("1:1");
    firstRow.getFormat().getFill().setColor("002060"); // dark blue
    firstRow.getFormat().getFont().setColor("FFFFFF"); // white
    firstRow.getFormat().getFont().setBold(true); // bold
    firstRow.getFormat().getFont().setSize(18); // font size 18

    // Get the used range of the worksheet
    let usedRange = worksheet.getUsedRange();

    // Get the total number of columns
    let columnCount = usedRange.getColumnCount();

    // Iterate over each column and set its width
    for (let i = 0; i < columnCount; i++) {
      // Get range of the current column
      let columnRange = worksheet.getRangeByIndexes(0, i, usedRange.getRowCount(), 1);

      // Set the width of this column to ~250 pixels (approx 31.25 character units)
      columnRange.getFormat().setColumnWidth(200);
    }
  });
}
