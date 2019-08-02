import { HEAD_FORMULA, ERRORS } from './shared/constant';
import { ProjectData } from './shared/model/interfaces/ProjectData';

//  Save the new value data in the Excel file
export const save = async (
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
      console.log(data[1]); //data 1 es el sheet
      
      data[1] = data[1].substring(1, data[1].length - 1);
      data[2] = data[2].split('{')[1];
      data[data.length - 1] = data[data.length - 1].split('}')[0];
      data[2] = data[2].split(';');
      data[2].map((value) => {
        data.push(value);
        console.log('this is data ',data);
        console.log('this is value',value);
        
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
