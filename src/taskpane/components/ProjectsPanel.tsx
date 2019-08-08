import * as React from 'react';
import { TOTAL } from './shared/constant';
import { Employee } from './shared/model/interfaces/Employee';

export const ProjectsPanel: React.FC<{
  employee: Employee;
  setDataEmployee: Function;
  setDataError: Function;
  save: Function;
}> = (props) => {
  const handleOnChange = async (value: string, index: number) => {
    props.setDataEmployee(value, index);
    if (!isNaN(Number(value)) && value !== '') {
      props.setDataError(false, index);
    } else {
      props.setDataError(true, index);
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
        {props.employee.worksheetData.map((definition: any, idx: number) => {
          return (
            <tr className="project" key={idx}>
              <td className="projectName">{definition.name}</td>
              <td>
                <input
                  id={idx.toString()}
                  key={definition.name}
                  className={
                    definition.error ? 'projectFTE error-value' : 'projectFTE'
                  }
                  value={definition.value}
                  onChange={(event) => handleOnChange(event.target.value, idx)}
                  onKeyPress={(event) => {
                    if (event.key === 'Enter') {
                      props.save();
                    }
                  }}
                  autoComplete="off"
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
