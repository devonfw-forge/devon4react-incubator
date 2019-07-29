import * as React from 'react';
import { handleHourChange } from './SaveHour';

export const ProjectsPanel: React.FC<{state}> = (props) => {
    
    return (
        <div>
            <div className='employeeName'>
                <h2>{props.state.employeeName}</h2>
            </div>
            <div className='projectsContainer'>
            {props.state.projects.map((project: any, i: number) => {
                return (
                <div className='project' key={i}>
                    <h3>{project.name}</h3>
                    <input id={i.toString()} key={project.hours} defaultValue={project.hours} onKeyUp={(event) => handleHourChange(event, i, props.state)}/>
                </div>
                )
                })}
        </div>
            <div className='total'>
                <h2>Total</h2>
                <h2>{props.state.total}</h2>
            </div>
        
        </div>
    )
}

//todo change function name, or handle name, cambiar el savehour.tsx