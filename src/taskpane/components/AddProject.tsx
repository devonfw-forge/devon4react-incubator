import * as React from 'react';
import { handleProjName, addProj } from './ProjectsData';

export default class AddProject extends React.Component<{projSheet: Excel.Worksheet, click: any}, {newProj: string}> {

    constructor(props: any, context: Excel.RequestContext) {
        super(props, context);
        handleProjName.bind(this);
        this.state = {
            newProj: null
          };
      }

  render() {
    return (
        <div>
            <input type="text" placeholder="Project Name" onChange={ (event) => handleProjName(event, this.state) }/>
            <button className='ms-welcome__action' onClick={() => addProj(this.state, this.props)}>Add Project</button>
        </div>
    );
  }

}