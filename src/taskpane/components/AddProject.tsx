import * as React from 'react';

export default class AddProject extends React.Component<{projSheet: Excel.Worksheet, click: any}, {newProj: string}> {

    constructor(props: any, context: Excel.RequestContext) {
        super(props, context);
        this.handleProjName = this.handleProjName.bind(this);
        this.state = {
            newProj: null
          };
      }

 // Check the value typed by the user in Add New Project field 
  // Called when the user start typing in Add New Project fields
  handleProjName(e: any) {
    this.setState({
      newProj: e.target.value
    }); // Set the state newProj with the name of the new project
  }

  // Save the new project in the Excel file
  addProj = async () => {
    if (this.state.newProj !== null) { // Check if the New Project Name is not empty
      try {
        await Excel.run(async context => {
          const projectsSheet = context.workbook.worksheets.getItem(this.props.projSheet.name); // Get the Excel sheet to update
          const table = projectsSheet.tables.getItemAt(0); // Get the Projects Table in the sheet to update
          const numOfCol = table.columns.getCount(); // Count numbers of Columns in the table
          await context.sync();
          const newProjectVal = []; // Empty array to set the default value of the New Project
          newProjectVal.push(this.state.newProj); // Push the name of the New Project in the array
          for (let i = 1; i < numOfCol.value; i++) {
            newProjectVal.push(0); // Set the number of hours to 0 for each Column
          }
          table.rows.load("items");
          await context.sync();
          table.rows.add(table.rows.items.length - 2, [newProjectVal]); // Add the new project at the second last position of the Projects Table
          await context.sync();
          context.workbook.worksheets.getFirst().activate() // Get and activate the first Excel sheet 
          setTimeout(() => {
            this.props.click(); // Reload the Projects Table with the new Project added once the first Excel sheet is activated
          }, 1);
        })
      } catch (error) {
        console.error(error);
      }
    }
  }

  render() {
    return (
        <div>
            <input type="text" placeholder="Project Name" onChange={ this.handleProjName }/>
            <button className='ms-welcome__action' onClick={this.addProj}>Add Project</button>
        </div>
    );
  }

}