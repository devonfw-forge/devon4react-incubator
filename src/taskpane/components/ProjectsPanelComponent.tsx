import * as React from 'react';
import { handleOnChange } from './SaveHour';

export const ProjectsPanel: React.FC<{ state }> = (props) => {
  return (
    <table className="projectGride">
      <thead className="employeeName">
        <tr>
          <th colSpan={2}>{props.state.employeeName}</th>
        </tr>
      </thead>
      <tbody className="projectsContainer">
        {props.state.projects.map((project: any, i: number) => {
          return (
            <tr className="project" key={i}>
              <td className="projectName">{project.name}</td>
              <td>
                <input
                  id={i.toString()}
                  key={project.hours}
                  defaultValue={project.hours}
                  onKeyUp={(event) => handleOnChange(event, i, props.state)}
                />
              </td>
            </tr>
          );
        })}
      </tbody>
      <tfoot className="total">
        <tr>
          <td>Total</td>
          <td>{props.state.total}</td>
        </tr>
      </tfoot>
    </table>
  );
};
