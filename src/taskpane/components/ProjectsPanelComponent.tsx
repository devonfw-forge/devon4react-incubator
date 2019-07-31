import * as React from 'react';
import { handleOnChange } from './SaveHour';
import { TOTAL } from './shared/constant';
import { ProjectData } from './shared/model/interfaces/ProjectData';

export const ProjectsPanel: React.FC<{
  state: any;
  setError: Function;
  setDataLoaded: Function;
}> = (props) => {
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
                  key={project.value}
                  defaultValue={project.value}
                  onKeyUp={(event) =>
                    handleOnChange(
                      event,
                      i,
                      props.state,
                      props.setError,
                      props.setDataLoaded,
                    )
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
          <td>{props.state.total}</td>
        </tr>
      </tfoot>
    </table>
  );
};
