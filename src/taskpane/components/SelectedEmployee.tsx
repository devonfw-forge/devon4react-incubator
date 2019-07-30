import { ERRORS, CALC, WORKSHEET_ERRORS } from "./shared/constant";

export const getSelectedEmployeeData = async (
  context: Excel.RequestContext,
  updateTotal,
  setError,
  setShowTable,
) => {
  const activeSheet = context.workbook.worksheets.getActiveWorksheet(); //Get the active Excel sheet
  const range = activeSheet.context.workbook
  .getSelectedRange()
  .load(['address', 'values', 'rowIndex', 'formulas']); // Get the selected cell location, value and index of its row
  await context.sync();
  
  
  const checkFormula = new RegExp('^=CAP.RENDER(.*)', 'gmi');
  if (!checkFormula.test(range.formulas[0][0])) {
    setError(true, ERRORS.INCORRECT_CELL, false);
    setShowTable(false);
  } else {
    setShowTable(true);
  }

  const employeeHeaderAddress = activeSheet
    .findAll('Employee', {
      completeMatch: true,
      matchCase: false, // Case insensitive
    })
    .load('address'); // Look for the word "Employee" in the active sheet and get its cell location
  await context.sync();
  const userCol = employeeHeaderAddress.address.split('!')[1][0]; // Get the Column of the cell with the value "Employee" in the first Excel sheet
  const activeRow = range.rowIndex + 1; // Get the index of the row of the selected cell
  const activeEmployee = activeSheet
    .getRange(userCol + activeRow)
    .load('values'); // Get the name of the selected Employee
  await context.sync();

  const employeeData = range.formulas[0][0].split('(')[1].split(',');
  let data = {
    dataSheet: employeeData[0].substring(1, employeeData[0].length - 1),
    value: undefined,
  };
  
  employeeData[1] = employeeData[1].split('{')[1];
  employeeData[employeeData.length - 1] = employeeData[
    employeeData.length - 1
  ].split('}')[0];
  employeeData[1] = employeeData[1].split(';');
  data.value = employeeData[1].map((hour) => {
    return hour;
  });

  if (data.dataSheet === '') {
    setError(true, WORKSHEET_ERRORS.EMPTY, false);
    setShowTable(false);
  }
  context.workbook.worksheets.load('items');
  await context.sync();
  const sheetsName = [];
  context.workbook.worksheets.items.map((sheet) => {
    sheetsName.push(sheet.name.toLowerCase())
  });
  if (data.dataSheet !== '' && sheetsName.indexOf(data.dataSheet.toLowerCase()) === -1) {
    setError(true, WORKSHEET_ERRORS.NOT_FOUND, false);
    setShowTable(false);
  }

  if (range.values[0][0] !== CALC) {
    updateTotal(range.values[0][0]);
  }
  return { activeEmployee, data };
};
