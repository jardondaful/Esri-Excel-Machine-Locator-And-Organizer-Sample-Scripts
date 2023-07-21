function main(workbook: ExcelScript.Workbook) {
  // Get the source worksheet by name
  let sourceSheet = workbook.getWorksheet("M2NELab");

  // Get the destination worksheet by name
  let destinationSheet = workbook.getWorksheet("Sheet1");

  // Get the used range of the source sheet
  let sourceRange = sourceSheet.getUsedRange();

  // Output the number of rows and columns in the source range
  let numRows = sourceRange.getRowCount();
  let numCols = sourceRange.getColumnCount();
  console.log("Source Range: " + numRows + " rows, " + numCols + " columns");

  // Get the values in the source range
  let values = sourceRange.getValues();

  // Filter and copy only columns A to J
  let filteredValues = values.map(row => row.slice(0, 10));

  // Output the filtered values array
  console.log("Filtered Values:");
  console.log(filteredValues);

  // Clear the destination sheet before pasting the data
  let destinationRange = destinationSheet.getRange("A1:J" + filteredValues.length);

  // Set the values in the destination sheet
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


