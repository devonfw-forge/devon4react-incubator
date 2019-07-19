import * as React from 'react';
import { handleProjName, addProj } from './ProjectsData';

export const AddProject: React.FC<{state:any, projSheet: Excel.Worksheet, click: any}> = (props) => {
    return (
        <div>
            <input type="text" placeholder="Project Name" onChange={ (event) => handleProjName(event, props.state) }/>
            <button className='ms-welcome__action' onClick={() => addProj(props)}>Add Project</button>
        </div>
    )
}
