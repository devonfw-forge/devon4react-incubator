import { debounce } from 'lodash';
 
 // Save the new hour data in the Excel file
 const save = debounce(async (index: number, state: any) => {
    try {
      await Excel.run(async context => {
        const projectsSheet = context.workbook.worksheets.getItem(state.projectsSheet.name); // Get the Excel sheet to update
        const cellToUpdate = projectsSheet.getRange(state.hoursList[index].address.split("!")[1]).load("values"); // Get the cell to update and its current value 
        await context.sync();
        cellToUpdate.values = [[state.hoursList[index].value]]; // Update the cell with the new value
      })
    } catch (error) {
      console.error(error);
    }
  }, 200); // Wait 0.2 seconds when the function is called before to do it

  // Check the value typed by the user in Hours fields 
  // Called when the user start typing in Hours fields
  const handleHourChange = async (e: any, index: number, state: any) => {
    const newValue = Number.parseInt(e.currentTarget.textContent); // Set the value to number, will be NaN if the value is composed of characters which are not numbers 
    if (newValue !== NaN) { // Check if the typed value is a number or NaN
      state.hoursList[index].value = newValue; // Change the hour value with the new value in the state hoursList
      save(index, state); // Calls the function to save the new value in the Excel file
    }
  }

export { handleHourChange };