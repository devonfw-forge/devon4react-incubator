import { HEAD_FORMULA, ERRORS } from "./shared/constant";
import { ProjectData } from "./shared/model/interfaces/ProjectData";

//  Save the new value data in the Excel file
const save = async (index: number, projects: ProjectData[], employeeCell: string) => {
  try {
    await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getFirst(); // Get the Excel sheet to update
      const cellToUpdate = activeSheet.context.workbook
        .getSelectedRange()
        .load(['address', 'values', 'rowIndex', 'formulas']);
      await context.sync();
      const data = cellToUpdate.formulas[0][0].split('(')[1].split(',');
      data[1] = data[1].substring(1, data[1].length - 1);
      data[2] = data[2].split('{')[1];
      data[data.length - 1] = data[data.length - 1].split('}')[0];
      data[2] = data[2].split(';');
      data[2].map((value) => {
        data.push(value);
      });
      
      data.splice(2, 1);
      data[index + 2] = projects[index].value;
      const formula =
        HEAD_FORMULA +
        employeeCell +
        ',"' + data[1] +
        '",{' +
        data
          .slice(2, data.length)
          .map((value: number) => {
            return value;
          })
          .join(';') +
        '})';

      cellToUpdate.formulas = [[formula]];
    });
  } catch (error) {
    console.error(error);
  }
};

// Check the value typed by the user in value fields
// Called when the user start typing in value fields
const handleOnChange = async (e: any, index: number, state: any, setError: Function, newProjects: ProjectData[]) => {
  
  if (!isNaN(e.currentTarget.value) && e.keyCode === 13) {
    setError(false, '', true);
    state.projects[index].value = e.currentTarget.value; // Change the value value with the new value
    save(index, state.projects, state.employeeCell); // Calls the function to save the new value in the Excel file
  } else if (isNaN(e.currentTarget.value) || e.currentTarget.value === '') {
    setError(true, ERRORS.VALUE, true);
  } else if (!isNaN(e.currentTarget.value)) {
    setError(false, '', true);
    newProjects[index].value = e.currentTarget.value;
    // console.log('asdasdasd', newProjects, state.projects);
  }
};

export { handleOnChange };
