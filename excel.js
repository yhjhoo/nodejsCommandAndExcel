let xlsx = require('xlsx');
let workbook = xlsx.readFile('output.xlsx');


const sheetNames = workbook.SheetNames;
const worksheet = workbook.Sheets[sheetNames[0]];
console.log(worksheet);

worksheet.A1.v = "hello";
worksheet.E4.v = "hi";
xlsx.writeFile(workbook, 'output.xlsx');
