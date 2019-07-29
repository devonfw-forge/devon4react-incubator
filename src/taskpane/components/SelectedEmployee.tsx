const getSelectedEmployeeData = async (
  context: Excel.RequestContext,
  updateTotal
) => {
  const activeSheet = context.workbook.worksheets.getActiveWorksheet(); //Get the first Excel sheet
  await activeSheet.activate(); // Activate the first Excel sheet
  const range = activeSheet.context.workbook
    .getSelectedRange()
    .load(['address', 'values', 'rowIndex', 'formulas']); // Get the selected cell location, value and index of its row
  await context.sync();
  const selectedCellPos = range.address.split('!')[1]; // Get the selected cell Column and Row
  const selectedCat = activeSheet
    .getRange(selectedCellPos[0] + '1')
    .load('values'); // Get the header value of the selected Column
  const employeeHeaderAddress = activeSheet
    .findAll('Employee', {
      completeMatch: true,
      matchCase: false // Case insensitive
    })
    .load('address'); // Look for the word "Employee" in the active sheet and get its cell location
  await context.sync();
  const userCol = employeeHeaderAddress.address.split('!')[1][0]; // Get the Column of the cell with the value "Employee" in the first Excel sheet
  const activeRow = range.rowIndex + 1; // Get the index of the row of the selected cell
  const activeEmployee = activeSheet
    .getRange(userCol + activeRow)
    .load('values'); // Get the name of the selected Employee
  await context.sync();
  const data = range.formulas[0][0].split('(')[1].split(',');
  data[0] = data[0].substring(1, data[0].length - 1);
  data[1] = data[1].split('{')[1];
  data[data.length - 1] = data[data.length - 1].split('}')[0];
  data[1] = data[1].split(';');
  data[1].map(hour => {
    data.push(hour);
  });
  data.splice(1, 1);

  if (range.values[0][0] !== '#CALC!') {    
    updateTotal(range.values[0][0]);
  }

  return { selectedCat, activeEmployee, data };
};

export { getSelectedEmployeeData };