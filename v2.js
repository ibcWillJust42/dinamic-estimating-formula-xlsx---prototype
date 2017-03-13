const Excel = require('exceljs');
const _ = require('lodash');

const arr = [{ a: 1, b: 2, c: 'a{{i}}+b{{i}}' }, { a: 3, b: 4, c: 'a{{i}}+b{{i}}' }]

const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet("Тестовый лист");
workbook.getWorksheet("Тестовый лист");

fillTable(arr);

workbook.xlsx.writeFile("./test_exceljs.xlsx");

console.log(42);




function fillTable(arr) {
  for (let i = 1; i <= arr.length; i++) {
    const k = i - 1;

    const string = arr[k];
    const pair = _.toPairs(string);
    console.log(pair);

    for (let j = 0; j < pair.length; j++) {
      const cell = pair[j];
      console.log('${cell}', `${cell}`);
      let key = cell[0];
      const value = cell[1];
      key = key.toUpperCase();  // нужен заглавный регистр букв в ячейках
      const cellPlace = [key, i].join('');
      console.log('cell join', cellPlace);

      const isFormula = _.isString(value) && value.includes('{{i}}');
      console.log(isFormula);
      if (isFormula) {
        const prepareFormula = value.replace(/{{i}}/g, i);
        const finalFormula = prepareFormula.toUpperCase();
        worksheet.getCell(cellPlace).value = { formula: finalFormula };
      } else {
        worksheet.getCell(cellPlace).value = value;
      }
    }

  }
}
