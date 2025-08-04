const ExcelJS = require('exceljs');
const fs = require("fs");
const readline = require("readline");

async function createAllAverageExcel() {
  const names = ["bg", "eg", "og", "tg", "bs", "es", "os", "ts"];
  const fileNames = names.map(name => name + "Avg.txt");
  
  // Read all files
  const allFileData = {};
  
  for (const fileName of fileNames) {
    const fileStream = fs.createReadStream(fileName);
    const rl = readline.createInterface({ input: fileStream });
    
    const data = [];
    const lineIterator = rl[Symbol.asyncIterator]();
    let result;
    
    // Read all 100 lines from each file
    while (!(result = await lineIterator.next()).done) {
      const value = result.value.trim();
      if (value !== "") {
        data.push(parseFloat(value));
      }
    }
    
    allFileData[fileName] = data;
    rl.close();
  }
  
  // Create workbook and worksheet
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('AllAverage');
  
  // Process 10 iterations
  let currentRow = 1;
  
  for (let iteration = 0; iteration < 10; iteration++) {
    // For each file in order
    for (const fileName of fileNames) {
      const rowData = [];
      
      // Get 10 consecutive values for this iteration
      for (let colIdx = 0; colIdx < 10; colIdx++) {
        const dataIndex = iteration * 10 + colIdx;
        rowData.push(allFileData[fileName][dataIndex]);
      }
      
      // Add row to worksheet
      const row = worksheet.addRow(rowData);
      
      // Find the minimum value and its column index
      const minValue = Math.min(...rowData);
      const minColumnIndex = rowData.indexOf(minValue) + 1; // ExcelJS uses 1-based indexing
      
      // Apply blue fill to the minimum value cell
      const minCell = row.getCell(minColumnIndex);
      minCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF87CEEB' } // Light blue
      };
      
      currentRow++;
    }
  }
  
  // Set column widths
  for (let i = 1; i <= 10; i++) {
    worksheet.getColumn(i).width = 15;
  }
  
  // Save the workbook
  await workbook.xlsx.writeFile('allAvg.xlsx');
  
  console.log(`Successfully created Excel file with ${currentRow - 1} rows and 10 columns`);
  console.log("Output file: allAvg.xlsx");
  console.log("Minimum values are highlighted in light blue");
}

createAllAverageExcel().catch(console.error);








