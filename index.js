const XLSX = require('xlsx');
const fs = require("fs");
const readline = require("readline");

async function readFileLineByLine(filePath, a) {
  const wb = XLSX.utils.book_new();
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

  // Append sheet and write file
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, "Exponential3.xlsx");

  console.log(`Processed ${recordCount} records`);
}

// readFileLineByLine("allBinomial.txt", "allBinAvg.txt");
readFileLineByLine("allEx.txt", "allExAvg.txt");