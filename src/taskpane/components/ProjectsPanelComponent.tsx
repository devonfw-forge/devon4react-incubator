import * as React from 'react';
import { handleHourChange } from './SaveHour';

export const ProjectsPanel: React.FC<{state}> = (props) => {
    return (
        <table className='projectsContainer'>
            <tbody>
            <tr>
                <th colSpan={2}>{props.state.employeeName}</th>
            </tr>
            {props.state.projects.values.map((project: string[], i: number) => {
                return (
                <tr key={i}>
                    <td>{project[0]}</td>
                    <td id={project[0]}>
                    <p suppressContentEditableWarning={true} contentEditable onKeyUp={(event) => handleHourChange(event, i, props.state)}>{props.state.hoursList[i].value}</p>
                    </td>
                </tr>
                )
                })}
            </tbody>
        </table>
    )
}