function main(workbook: ExcelScript.Workbook) {
  // Get the source worksheet by name
  let sourceSheet = workbook.getWorksheet("M2NELab");

  // Get the destination worksheet by name
  let destinationSheet = workbook.getWorksheet("Sheet1");

  // Get the used range of the source sheet
  let sourceRange = sourceSheet.getUsedRange();

  // Get the values in the source range
  let values = sourceRange.getValues();

  // Filter and copy only columns A to J
  let filteredValues = values.map(row => row.slice(0, 10));

  // Clear the destination sheet before pasting the data
  let destinationRange = destinationSheet.getRange("A1:J" + (filteredValues.length));
  destinationRange.clear();

  // Set the values in the destination sheet
  destinationRange.setValues(filteredValues);

  // Get the used range of the destination sheet
  let destinationUsedRange = destinationSheet.getUsedRange();

  // Get the values in the destination used range
  let destinationValues = destinationUsedRange.getValues();

  // Filter out blank rows
  let nonBlankValues = destinationValues.filter(row => row.some(cellValue => cellValue !== ""));

  // Clear the destination sheet
  destinationUsedRange.clear();

  // Set the non-blank values back to the destination sheet
  let nonBlankRange = destinationSheet.getRangeByIndexes(0, 0, nonBlankValues.length, nonBlankValues[0].length);
  nonBlankRange.setValues(nonBlankValues);
}
