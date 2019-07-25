import * as React from 'react';
import { handleHourChange } from './SaveHour';

export const ProjectsPanel: React.FC<{state}> = (props) => {
    return (
        <table className='projectsContainer'>
            <tbody>
            <tr>
                <th colSpan={2}>{props.state.employeeName}</th>
            </tr>
            {props.state.projects.map((project: any, i: number) => {
                return (
                <tr key={i}>
                    <td>{project.name}</td>
                    <td id={i.toString()}>
                        <p suppressContentEditableWarning={true} contentEditable onKeyUp={(event) => handleHourChange(event, i, props.state)}>{project.hours}</p>
                    </td>
                </tr>
                )
                })}
            </tbody>
        </table>
    )
}

//todo change function name, or handle name, cambiar el savehour.tsx