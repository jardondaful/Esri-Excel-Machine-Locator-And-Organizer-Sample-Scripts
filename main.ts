function main(workbook: ExcelScript.Workbook) {
  // Get the source worksheet by name
  let sourceSheet = workbook.getWorksheet("M2NELab");

  // Get the destination worksheet by name (Sheet1)
  let destinationSheet = workbook.getWorksheet("Sheet1");

  // If the destination sheet doesn't exist, create it
  if (!destinationSheet) {
    destinationSheet = workbook.addWorksheet("Sheet1");
  }

  // Clear everything within "Sheet1"
  destinationSheet.getUsedRange()?.clear();

  // Set the style of the first row for destinationSheet
  let firstRow = destinationSheet.getRange("1:1");
  firstRow.getFormat().getFill().setColor("002060"); // dark blue
  firstRow.getFormat().getFont().setColor("FFFFFF"); // white
  firstRow.getFormat().getFont().setBold(true); // bold

  // Get the used range of the source sheet
  let sourceRange = sourceSheet.getUsedRange();

  // Get the values in the source range
  let values = sourceRange.getValues();

  // Filter and copy only columns A to J
  let filteredValues = values.map(row => row.slice(0, 10));

  // Output the filtered values array
  console.log("Filtered Values:");
  console.log(filteredValues);

  // Get the used range of the destination sheet
  let destinationUsedRange = destinationSheet.getUsedRange();

  // Get the last row number with text in the destination sheet
  let lastRowWithText = destinationUsedRange ? destinationUsedRange.getRowCount() : 0;

  // Determine the destination range for the source data
  let destinationRange = destinationSheet.getRangeByIndexes(lastRowWithText, 0, filteredValues.length, filteredValues[0].length);

  // Set the values in the destination sheet from the source
  destinationRange.setValues(filteredValues);

  // Append data from "M2NW Hallway (Front of 60) -->"
  let sourceSheet2 = workbook.getWorksheet("M2NW Hallway (Front of 60) -->");
  appendDataToSheet(sourceSheet2, destinationSheet);

  // Append data from "M2NE Hallway (Front of 360) -->"
  let sourceSheet3 = workbook.getWorksheet("M2NE Hallway (Front of 360) -->");
  appendDataToSheet(sourceSheet3, destinationSheet);
}

function appendDataToSheet(sourceSheet: ExcelScript.Worksheet, destinationSheet: ExcelScript.Worksheet) {
  // Get the used range of the source sheet
  let sourceRange = sourceSheet.getUsedRange();

  // Get the values in the source range
  let values = sourceRange.getValues();

  // Get the used range of the destination sheet
  let destinationUsedRange = destinationSheet.getUsedRange();

  // Get the last row number with text in the destination sheet
  let lastRowWithText = destinationUsedRange ? destinationUsedRange.getRowCount() : 0;

  // Determine the destination range for the source data
  let destinationRange = destinationSheet.getRangeByIndexes(lastRowWithText, 0, values.length, values[0].length);

  // Set the values in the destination sheet from the source
  destinationRange.setValues(values);
}
