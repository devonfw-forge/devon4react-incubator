export const getSelectedEmployeeData = async (
  context: Excel.RequestContext,
  setError,
) => {
  const activeSheet = context.workbook.worksheets.getActiveWorksheet(); //Get the first Excel sheet
  await activeSheet.activate(); // Activate the first Excel sheet
  const range = activeSheet.context.workbook
    .getSelectedRange()
    .load(['address', 'values', 'rowIndex', 'formulas']); // Get the selected cell location, value and index of its row
  await context.sync();

  console.log(range.formulas[0][0]);
  const checkFormula = new RegExp('^=DEVON.RENDERER(.*)', 'gmi');
  setError(true, 'Select a Cell with Render formula');
  if (!checkFormula.test(range.formulas[0][0])) {
    console.log('waka');
    setError(true, 'Select a Cell with Render formula');
  }

  const selectedCellPos = range.address.split('!')[1]; // Get the selected cell Column and Row
  const selectedCat = activeSheet
    .getRange(selectedCellPos[0] + '1')
    .load('values'); // Get the header value of the selected Column
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
    fte: undefined,
  };
  employeeData[1] = employeeData[1].split('{')[1];
  employeeData[employeeData.length - 1] = employeeData[
    employeeData.length - 1
  ].split('}')[0];
  employeeData[1] = employeeData[1].split(';');
  data.fte = employeeData[1].map((hour) => {
    return hour;
  });

  return { selectedCat, activeEmployee, data };
};
