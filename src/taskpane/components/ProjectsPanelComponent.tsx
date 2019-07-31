import * as React from 'react';
import { save } from './SaveHour';
import { TOTAL } from './shared/constant';
import { ERRORS } from './shared/constant';

export const ProjectsPanel: React.FC<{
  state: any;
  setError: Function;
  setDataLoaded: Function;
}> = (props) => {
  const handleOnChange = async (
    event: any,
    index: number,
    state: any,
    setError: Function,
    setDataLoaded: Function,
  ) => {
    const projects = document.getElementsByClassName('projectFTE');
    const projs = new Array();
    const reg = new RegExp('[A-Za-z]', 'gmi');
    let error = false;
    for (let i = 0; i < projects.length; i++) {
      projs.push(projects[i]);
    }
    for (let i = 0; i < projs.length; i++) {
      if (reg.test(projs[i].value) || projs[i].value === '') {
        setError(true, ERRORS.VALUE, 'red');
        setDataLoaded(true);
        error = true;
      }
    }

    if (isNaN(event.currentTarget.value) || event.currentTarget.value === '') {
      props.state.projects[event.currentTarget.id].error = true;
      error = true;
      setError(true, ERRORS.VALUE, 'red');
      setDataLoaded(true);
    } else if (!isNaN(event.currentTarget.value) && !error) {
      props.state.projects[event.currentTarget.id].error = false;
      setError(false, '', 'white');
      setDataLoaded(true);
    } else if (
      !isNaN(event.currentTarget.value) &&
      !reg.test(event.currentTarget.value)
    ) {
      props.state.projects[event.currentTarget.id].error = false;
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
                  key={project}
                  className={
                    project.error ? 'projectFTE error-value' : 'projectFTE'
                  }
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
