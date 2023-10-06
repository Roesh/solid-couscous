var XLSX = require("xlsx");

const numberOfEntries = 37
const mainDataTab = "Program Data"

const testData = Array(numberOfEntries).fill({
    programName: 'CON-IT',
    size: 1,
    overallHealth: 'Green',
    Schedule: 'Green',
    Cost: 'Green',
    Resource: 'Green',
    Risk: 'Green',
    Comment: 'string',
    Risks: 'string',
    Achievements: 'string'
})



XLSX.utils.json_to_sheet(testData)
const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(testData);

XLSX.utils.book_append_sheet(workbook, worksheet, mainDataTab);

XLSX.writeFile(workbook, "test.xlsx");


const readWorkbook = XLSX.readFile("test.xlsx");
const readWorksheet = readWorkbook.Sheets[mainDataTab]

// Convert the worksheet to an array of objects
const jsonData = XLSX.utils.sheet_to_json(readWorksheet);
console.log(jsonData)