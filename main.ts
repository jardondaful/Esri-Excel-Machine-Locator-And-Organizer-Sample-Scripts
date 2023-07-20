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

  // Get the used range of the destination sheet
  let destinationUsedRange = destinationSheet.getUsedRange();

  // Get the last row number with text in the destination sheet
  let lastRowWithText = destinationUsedRange ? destinationUsedRange.getRowCount() : 0;

  // Now, let's append data from "M2NW Hallway (Front of 60) -->"
  let sourceSheet2 = workbook.getWorksheet("M2NW Hallway (Front of 60) -->");

  // Get the used range of the second source sheet
  let sourceRange2 = sourceSheet2.getUsedRange();

  // Get the values in the second source range
  let values2 = sourceRange2.getValues();

  // Determine the destination range for the second source data
  let destinationRange2 = destinationSheet.getRangeByIndexes(lastRowWithText, 0, values2.length, values2[0].length);

  // Set the values in the destination sheet from the second source
  destinationRange2.setValues(values2);

  // Get the last row number with text in the destination sheet
  let finalRowWithText = destinationSheet.getUsedRange().getRowCount();

  // Get the last column number with text in the destination sheet
  let finalColumnWithText = destinationSheet.getUsedRange().getColumnCount();

  // Output the last row and last column with text
  console.log("Final Row with Text: " + finalRowWithText);
  console.log("Final Column with Text: " + finalColumnWithText);
}
