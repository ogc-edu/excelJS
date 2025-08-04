const XLSX = require('xlsx');
const fs = require("fs");
const readline = require("readline");
/*
1. Takes all 80 models best in each iteration with average into 8 different sheets, one sheet = 10 models 
*/ 

async function readFileLineByLine(filePath, a, wb, sheetName) {
  const ws = {};

  let startRow = 2;  // Starting at row 2 (0-based index)
  let startCol = 2;  // Starting at column 2 (0-based index)
  let recordCount = 0;

  const fileStream = fs.createReadStream(filePath);
  const rl = readline.createInterface({ input: fileStream });
  const fileStreamAvg = fs.createReadStream(a);
  const r2 = readline.createInterface({input: fileStreamAvg});

  const lineIterator = rl[Symbol.asyncIterator](); // Get async iterator
  const avgIterator = r2[Symbol.asyncIterator]();
  let result;
  const Avg = new Array();
  // Track iteration for both rows and columns within a record
  let rowIteration = 0;
  let colIteration = 0;

  while (!(result = await lineIterator.next()).done) {
    const cellValue = result.value.trim(); // Remove extra spaces

    if (cellValue === "") continue; // Skip empty lines

    const cell = { v: cellValue, t: "n" }; // String type for better compatibility

    const currentRow = startRow + rowIteration;
    const currentCol = startCol + colIteration;

    // Encode cell address
    const cellAddress = XLSX.utils.encode_cell({ 
      r: currentRow - 1, 
      c: currentCol - 1 
    });

    // Write cell to worksheet
    ws[cellAddress] = cell;

    // Increment row iteration
    rowIteration++;

    // When we've filled 10 rows, move to next column
    if (rowIteration === 10) {
      rowIteration = 0;
      colIteration++;
    }
    var avg;
    // When we've filled 10 columns (100 cells total), move to next record
    if (colIteration === 10) {
      recordCount++;
      for(let i = 0; i < 10; i++) {
        const avg = await avgIterator.next();
        Avg[i] = avg.value.trim();
      }
      
      for(let i = 0; i < 10; i++) {
        const cellAddress = XLSX.utils.encode_cell({ 
          r: (recordCount * 12)-1, 
          c: 1 + i
        });
        const cell = { v: Avg[i], t: "n" };
        ws[cellAddress] = cell;
      }
      // Reset column iteration
      colIteration = startCol;
      colIteration = 0;
      // Move start position for next record
      // Each record starts 12 rows below the previous one (10 data rows + 2 empty rows)
      startRow += 12;
      startCol = 2; // Reset column start
    }
  }

  // Set the reference range to cover all written data
  ws["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 }, 
    e: { 
      r: startRow + 11, // Add extra rows to ensure full coverage
      c: startCol + 9  // 10 columns (0-based)
    }
  });

  // Append sheet to workbook
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  console.log(`Processed ${recordCount} records for sheet: ${sheetName}`);
}
async function processAllFiles() {
  const sheets = ["bg", "eg", "og", "tg", "bs", "es", "os", "ts"];
  // const fileNames = ["bg", "eg", "og", "tg", "bs", "es", "os", "ts"].map(fileName => fileName + "_best.txt");
  const fileNames = ["bg", "eg", "og", "tg", "bs", "es", "os", "ts"].map(fileName => fileName + ".txt");
  const fileNamesAvg = ["bg", "eg", "og", "tg", "bs", "es", "os", "ts"].map(fileName => fileName + "Avg.txt");
  
  // Create one workbook for all sheets
  const wb = XLSX.utils.book_new();
  
  // Process each file and create a sheet
  for(let i = 0; i < fileNames.length; i++) {
    await readFileLineByLine(fileNames[i], fileNamesAvg[i], wb, sheets[i]);
  }
  excelName = "set1FV.xlsx";
  // Write the single Excel file with all sheets
  XLSX.writeFile(wb, excelName);
  console.log(`All sheets created in ${excelName}`);
}

// Run the main function
processAllFiles().catch(console.error);