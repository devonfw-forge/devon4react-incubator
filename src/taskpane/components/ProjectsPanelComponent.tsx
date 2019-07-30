import * as React from 'react';
import { handleOnChange } from './SaveHour';
import { TOTAL } from './shared/constant';
import { ProjectData } from './shared/model/interfaces/ProjectData';

export const ProjectsPanel: React.FC<{ state: any; setError: Function;}> = (props) => {
  const newProjects: ProjectData[] = [];
  props.state.projects.map((project) => {
    newProjects.push(project);
  })
  return (
    <div>
      <div className="employeeName">
        <h2>{props.state.employeeName}</h2>
      </div>
      <div className="projectsContainer">
        {props.state.projects.map((project: ProjectData, i: number) => {
          return (
            <div className="project" key={i}>
              <h3 className="projectName">{project.name}</h3>
              <input
                id={i.toString()}
                key={project.value}
                defaultValue={project.value.toString()}
                onKeyUp={(event) => handleOnChange(event, i, props.state, props.setError, newProjects)}
              />
            </div>
          );
        })}
      </div>
      <div className="total">
        <h2>{TOTAL}</h2>
        <h2>{props.state.total}</h2>
      </div>
    </div>
  );
};
