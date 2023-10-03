var XLSX = require("xlsx");

const testData = [
    {programName: 'CON-IT'}
]

XLSX.utils.json_to_sheet(testData)
const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(testData);

XLSX.utils.book_append_sheet(workbook, worksheet, "Program Data");

XLSX.writeFile(workbook, "test.xlsx");