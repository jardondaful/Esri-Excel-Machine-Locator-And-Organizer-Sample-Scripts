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
  });
}
