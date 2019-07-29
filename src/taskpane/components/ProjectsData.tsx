// Get Projects' data from other sheets
const getProjectsData = async (context: Excel.RequestContext, employeeData, state: any) => {
  const projectsCol = context.workbook.worksheets
  .getItem(employeeData.data[0])
  .tables.getItemAt(0)
  .columns.load('items');
  
  await context.sync();
  const projects: string[][] = projectsCol.items[0].values.slice(1, projectsCol.items[0].values.length);
  employeeData.data
  .slice(1, employeeData.data.length)
  .map((hour: any, i: number) => {
    state.projects.push({ name: projects[i][0], hours: hour }); // Set the state projects with the projects from the sheet with their data
  });

  await context.sync();
  return state;
};

// Save Projects' data on other sheets

  // Check the value typed by the user in Add New Project field 
  // Called when the user start typing in Add New Project fields
  const handleProjName = (e: any, state: any) => {
      state.newProj = e.target.value // Set the state newProj with the name of the new project
  }

  // Save the new project in the Excel file
  const addProj = async (props: any) => {
    if (props.state.newProj !== null) { // Check if the New Project Name is not empty
      try {
        await Excel.run(async context => {
          const projectsSheet = context.workbook.worksheets.getItem(props.projSheet.name); // Get the Excel sheet to update
          const table = projectsSheet.tables.getItemAt(0); // Get the Projects Table in the sheet to update
          const numOfCol = table.columns.getCount(); // Count numbers of Columns in the table
          await context.sync();
          const newProjectVal = []; // Empty array to set the default value of the New Project
          newProjectVal.push(props.state.newProj); // Push the name of the New Project in the array
          for (let i = 1; i < numOfCol.value; i++) {
            newProjectVal.push(0); // Set the number of hours to 0 for each Column
          }
          table.rows.load("items");
          await context.sync();
          table.rows.add(table.rows.items.length, [newProjectVal]); // Add the new project at the second last position of the Projects Table
          await context.sync();
          context.workbook.worksheets.getFirst().activate() // Get and activate the first Excel sheet 
          setTimeout(() => {
            props.click(); // Reload the Projects Table with the new Project added once the first Excel sheet is activated
          }, 1);
        })
      } catch (error) {
        console.error(error);
      }
    }
  }

export { getProjectsData, handleProjName, addProj };