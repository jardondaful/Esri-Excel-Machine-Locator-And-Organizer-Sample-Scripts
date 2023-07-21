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

  // Set the style of the first row for destinationSheet
  let firstRowDestinationSheet = destinationSheet.getRange("1:1");
  firstRowDestinationSheet.getFormat().getFill().setColor("002060"); // dark blue
  firstRowDestinationSheet.getFormat().getFont().setColor("FFFFFF"); // white
  firstRowDestinationSheet.getFormat().getFont().setBold(true); // bold
  firstRowDestinationSheet.getFormat().getFont().setSize(18); // font size 18

  // Append data from "M2NW Hallway (Front of 60) -->"
  let sourceSheet2 = workbook.getWorksheet("M2NW Hallway (Front of 60) -->");
  appendDataToSheet(sourceSheet2, destinationSheet);

  // Append data from "M2NE Hallway (Front of 360) -->"
  let sourceSheet3 = workbook.getWorksheet("M2NE Hallway (Front of 360) -->");
  appendDataToSheet(sourceSheet3, destinationSheet);

  // Sort and categorize the destination sheet based on column "I"
  sortAndCategorize(destinationSheet, workbook);

  // After appending data, call the following function
  addExtraRows(destinationSheet);
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

function sortAndCategorize(sheet: ExcelScript.Worksheet, workbook: ExcelScript.Workbook) {
  // Get values from the sheet
  let sheetValues = sheet.getUsedRange().getValues();

  // Sort the data based on column "I" (9th column, index 8)
  sheetValues.sort((a, b) => {
    if (a[8] < b[8]) {
      return -1;
    } else if (a[8] > b[8]) {
      return 1;
    } else {
      return 0;
    }
  });

  // Write sorted data back to the sheet
  sheet.getRangeByIndexes(0, 0, sheetValues.length, sheetValues[0].length).setValues(sheetValues);

  // Get column "I" (index 8) values from the sorted sheet values
  let columnIValues = sheetValues.map(row => row[8]);

  // Remove duplicates
  let uniqueValues = [...new Set(columnIValues)];

  // Define header row
  let headerRow = ["Computer Name", "Owner", "Model", "Serial #", "Asset", "Asset #", "Location", "Location Status", "Model #", "Notes", "Order", "Title"];

  uniqueValues.forEach((value) => {
    // Create new worksheet for each unique value, if it doesn't already exist
    let newSheetName = `Category_${value}`;
    let newSheet = workbook.getWorksheet(newSheetName);
    if (!newSheet) {
      newSheet = workbook.addWorksheet(newSheetName);
    } else {
      // If the worksheet already exists, clear the existing data
      newSheet.getUsedRange()?.clear();
    }

    // Write the header row to the new worksheet
    newSheet.getRangeByIndexes(0, 0, 1, headerRow.length).setValues([headerRow]);

    // Set the style of the first row for newSheet
    let firstRowNewSheet = newSheet.getRange("1:1");
    firstRowNewSheet.getFormat().getFill().setColor("002060"); // dark blue
    firstRowNewSheet.getFormat().getFont().setColor("FFFFFF"); // white
    firstRowNewSheet.getFormat().getFont().setBold(true); // bold
    firstRowNewSheet.getFormat().getFont().setSize(18); // font size 18

    // Filter rows based on the column "I" value and remove blank rows
    let filteredRows = sheetValues.filter(row => row[8] === value && row.join('').trim() !== '');

    // Write the filtered data to the new worksheet
    newSheet.getRangeByIndexes(1, 0, filteredRows.length, filteredRows[0].length).setValues(filteredRows);

    // After setting the values in the new sheet, call the following function
    addExtraRows(newSheet);
  });
}

// Define a new function to add extra rows
function addExtraRows(sheet: ExcelScript.Worksheet) {
  // Determine how many rows in the sheet
  let rowCount = sheet.getUsedRange().getRowCount();
  // The number of rows to skip
  let skipRows = 5;
  // Starting from the sixth row (0-based index)
  for (let i = skipRows; i < rowCount; i += skipRows + 1) {
    // Add a row
    sheet.getRangeByIndexes(i, 0, 1, sheet.getUsedRange().getColumnCount()).getFormat().getFill().setColor("000000");
    // Update rowCount because we added a row
    rowCount++;
  }
}
