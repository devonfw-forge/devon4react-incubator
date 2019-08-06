import * as React from 'react';
import { TOTAL, HEAD_FORMULA } from './shared/constant';
import { ProjectData } from './shared/model/interfaces/ProjectData';
import { Employee } from './shared/model/interfaces/Employee';

export const ProjectsPanel: React.FC<{
  employee: Employee;
  setError: Function;
}> = (props) => {
  const handleOnChange = async (
    event: any,
    index: number,
    employee: Employee,
    setError: Function,
  ) => {
    const projects = document.getElementsByClassName('projectFTE');
    const projs = new Array();
    const reg = new RegExp('[A-Za-z]', 'gmi');
    let errors = false;
    for (let i = 0; i < projects.length; i++) {
      projs.push(projects[i]);
    }
    for (let i = 0; i < projs.length; i++) {
      if (reg.test(projs[i].value) || projs[i].value === '') {
        // Set error incorrect value
        setError(true, 2);
        errors = true;
      }
    }

    if (isNaN(event.currentTarget.value) || event.currentTarget.value === '') {
      props.employee.worksheetData[event.currentTarget.id].error = true;
      errors = true;
      // Set error incorrect value
      setError(true, 2);
    } else if (!isNaN(event.currentTarget.value) && !errors) {
      props.employee.worksheetData[event.currentTarget.id].error = false;
      // Remove error incorrect value
      setError(false, 2);
    } else if (
      !isNaN(event.currentTarget.value) &&
      !reg.test(event.currentTarget.value)
    ) {
      props.employee.worksheetData[event.currentTarget.id].error = false;
      // Remove error incorrect value
      setError(false, 2);
    }

    if (!isNaN(event.currentTarget.value) && event.keyCode === 13 && !errors) {
      for (let i = 0; i < projs.length; i++) {
        employee.worksheetData[i].value = projs[i].value;
      }
      // Remove error incorrect value
      setError(false, 2);
      save(index, employee.worksheetData, employee.cell); // Calls the function to save the new value in the Excel file
    }
  };

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

  return (
    <table className="projectGride">
      <thead className="employeeName">
        <tr>
          <th colSpan={2}>{props.employee.name}</th>
        </tr>
      </thead>
      <tbody className="projectsContainer">
        {props.employee.worksheetData.map((definition: any, i: number) => {
          return (
            <tr className="project" key={i}>
              <td className="projectName">{definition.name}</td>
              <td>
                <input
                  id={i.toString()}
                  key={definition}
                  className={
                    definition.error ? 'projectFTE error-value' : 'projectFTE'
                  }
                  defaultValue={definition.value}
                  onKeyUp={(event) =>
                    handleOnChange(event, i, props.employee, props.setError)
                  }
                />
              </td>
            </tr>
          );
        })}
      </tbody>
      <tfoot className="total">
        <tr>
          <td>{TOTAL}</td>
          <td>{props.employee.total}</td>
        </tr>
      </tfoot>
    </table>
  );
};
