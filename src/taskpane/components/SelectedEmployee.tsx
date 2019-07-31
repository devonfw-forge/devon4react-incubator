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
  
  const checkFormula = new RegExp('^=ADC.DYNACOLUMNS(.*)', 'gmi');
  if (!checkFormula.test(range.formulas[0][0])) {
    setError(true, ERRORS.INCORRECT_CELL, false);
    setShowTable(false);
  } else {
    setShowTable(true);
  }
  
  const employeeData = range.formulas[0][0].split('(')[1].split(',');
  const activeEmployeeCell = employeeData[0]; // Get the cell reference of the selected Employee
  let data = {
    employeeCell: activeEmployeeCell,
    dataSheet: employeeData[1].substring(1, employeeData[1].length - 1),
    value: undefined,
  };
  
  employeeData[2] = employeeData[2].split('{')[1];
  employeeData[employeeData.length - 1] = employeeData[
    employeeData.length - 1
  ].split('}')[0];
  employeeData[2] = employeeData[2].split(';');
  data.value = employeeData[2].map((value) => {
    return value;
  });
  
  if (data.dataSheet === '') {
    setError(true, WORKSHEET_ERRORS.EMPTY, false);
    setShowTable(false);
  }
  context.workbook.worksheets.load('items');
  const activeEmployee = activeSheet
    .getRange(activeEmployeeCell)
    .load('values');
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
