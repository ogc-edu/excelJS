const XLSX = require('xlsx');
const fs = require("fs");
const readline = require("readline");

async function createAverageResultsExcel(filePath) {
  const fileStream = fs.createReadStream(filePath);
  const rl = readline.createInterface({ input: fileStream });
  
  const data = [];
  const lineIterator = rl[Symbol.asyncIterator]();
  let result;
  
  // Read all 100 lines
  while (!(result = await lineIterator.next()).done) {
    const value = result.value.trim();
    if (value !== "") {
      data.push(parseFloat(value));
    }
  }
  
  // Model names in order
  const modelNames = [
    "best/1", "best/2", "best/3", 
    "current-to-best/1", "cur-to-best/2", 
    "cur-to-rand/1", "cur-to-rand/2", 
    "rand/1", "rand/2", "rand/3"
  ];
  
  // Create workbook and worksheet
  const wb = XLSX.utils.book_new();
  const wsData = [];
  
  // Add header row
  const headerRow = ["Benchmark Function", ...modelNames];
  wsData.push(headerRow);
  
  // Organize data: each 10 consecutive values in the same row
  for (let funcIdx = 0; funcIdx < 10; funcIdx++) {
    const row = [`F${funcIdx + 1}`];
    
    // For each column, get consecutive values
    for (let colIdx = 0; colIdx < 10; colIdx++) {
      const dataIndex = funcIdx * 10 + colIdx;
      row.push(data[dataIndex]);
    }
    
    wsData.push(row);
  }
  
  // Create worksheet from data
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  
  // Add some formatting - make header row bold
  ws['!cols'] = [{ wch: 18 }, ...Array(10).fill({ wch: 15 })];
  
  // Append sheet and write file
  XLSX.utils.book_append_sheet(wb, ws, "Average");
  XLSX.writeFile(wb, "bgAvg.xlsx");
  
  console.log(`Successfully created Excel file with ${data.length} values arranged in 10x10 format`);
  console.log("Output file: averaged_results.xlsx");
}

createAverageResultsExcel("bgAvg.txt");
