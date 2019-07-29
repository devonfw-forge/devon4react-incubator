//  Save the new hour data in the Excel file
const save = async (index: number, projects: any) => {
  try {
    await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getFirst(); // Get the Excel sheet to update
      const cellToUpdate = activeSheet.context.workbook
        .getSelectedRange()
        .load(['address', 'values', 'rowIndex', 'formulas']);
      await context.sync();
      const data = cellToUpdate.formulas[0][0].split('(')[1].split(',');
      data[0] = data[0].substring(1, data[0].length - 1);
      data[1] = data[1].split('{')[1];
      data[data.length - 1] = data[data.length - 1].split('}')[0];
      data[1] = data[1].split(';');
      data[1].map((hour) => {
        data.push(hour);
      });
      data.splice(1, 1);
      data[index + 1] = projects[index].hours;
      const formula =
        '=DEVON.RENDERER("' +
        data[0] +
        '",{' +
        data
          .slice(1, data.length)
          .map((hour: any) => {
            return hour;
          })
          .join(';') +
        '})';
      cellToUpdate.formulas = [[formula]];
    });
  } catch (error) {
    console.error(error);
  }
}

// Check the value typed by the user in Hours fields
// Called when the user start typing in Hours fields
const handleOnChange = async (e: any, index: number, state: any) => {
  const newValue = Number.parseInt(e.currentTarget.value); // Set the value to number, will be NaN if the value is composed of characters which are not numbers
  
  if (!isNaN(newValue) && e.keyCode === 13) {
    // Check if the typed value is a number or NaN
    state.projects[index].hours = newValue.toString(); // Change the hour value with the new value in the state hoursList
    save(index, state.projects); // Calls the function to save the new value in the Excel file
  }
};

export { handleOnChange };
