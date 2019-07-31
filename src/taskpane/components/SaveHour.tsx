import { HEAD_FORMULA, ERRORS } from './shared/constant';
import { ProjectData } from './shared/model/interfaces/ProjectData';

//  Save the new value data in the Excel file
const save = async (
  index: number,
  projects: ProjectData[],
  employeeCell: string,
) => {
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
        ',"' +
        data[1] +
        '",{' +
        projects
          .map((project: ProjectData) => {
            return project.value;
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
const handleOnChange = async (
  event: any,
  index: number,
  state: any,
  setError: Function,
  setDataLoaded: Function,
) => {
  const projects = document.getElementsByClassName('projectFTE');
  const projs = new Array();
  let error = false;
  for (let i = 0; i < projects.length; i++) {
    projs.push(projects[i]);
  }
  for (let i = 0; i < projs.length; i++) {
    const reg = new RegExp('[A-Za-z]', 'gmi');
    if (reg.test(projs[i].value) || projs[i].value === '') {
      setError(true, ERRORS.VALUE, 'red');
      setDataLoaded(true);
      error = true;
    }
  }

  if (isNaN(event.currentTarget.value) || event.currentTarget.value === '') {
    console.log('1');
    error = true;
    setError(true, ERRORS.VALUE, 'red');
    setDataLoaded(true);
  } else if (!isNaN(event.currentTarget.value) && !error) {
    console.log('2');
    setError(false, '', 'white');
    setDataLoaded(true);
  }

  if (
    !isNaN(event.currentTarget.value) &&
    event.keyCode === 13 &&
    !state.error.showError
  ) {
    for (let i = 0; i < projs.length; i++) {
      state.projects[i].value = projs[i].value;
    }
    setError(false, '', 'white');
    setDataLoaded(true);
    save(index, state.projects, state.employeeCell); // Calls the function to save the new value in the Excel file
  }
};

export { handleOnChange };
