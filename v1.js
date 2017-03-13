const Excel = require('exceljs');

const arr = [{ a: 1, b: 2}, { a: 3, b: 4}]

const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet("Тестовый лист");
workbook.getWorksheet("Тестовый лист");
fillTable(arr);

workbook.xlsx.writeFile("./test_exceljs.xlsx");

console.log(42);


function fillTable(arr) {
  for (let i = 1; i <= arr.length; i++) {
    const k = i - 1;
    worksheet.getCell(`A${i}`).value = arr[k].a;
    worksheet.getCell(`B${i}`).value = arr[k].b;
    worksheet.getCell(`C${i}`).value = { formula: `A${i}+B${i}`};
  }
}
