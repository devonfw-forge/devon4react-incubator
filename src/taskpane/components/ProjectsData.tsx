// Get Projects' data from other sheets
const getProjectsData = async (context: Excel.RequestContext, employeeData, state: any) => {
    const projCol = state.projectsSheet.findAll("Projects", {
        completeMatch: true,
        matchCase: false
    }).load("address"); // Look for the word "Projects" in the sheet with projects and get its cell location
    const totalCol = state.projectsSheet.findAll("Total", {
        completeMatch: true,
        matchCase: false
    }).load("address"); // Look for the word "Total" in the sheet with projects and get its cell location
    await context.sync();
    const firstProjPos = state.projectsSheet.getRange(projCol.address).getRowsBelow(1).load("address"); // Get the sheet and cell location of the first project in the list
    const lastProjPos = state.projectsSheet.getRange(totalCol.address).getRowsAbove(1).load("address"); // Get the sheet and cell location of the last project in the list
    await context.sync();
    const firstProjCell = firstProjPos.address.split("!")[1]; // Get the cell location of the first project in the list
    const lastProjCell = lastProjPos.address.split("!")[1]; // Get the cell location of the last project in the list
    const colToCheck = state.projectsSheet.findAll(employeeData.activeEmployee.values[0][0], {
        completeMatch: true,
        matchCase: false
    }).load("address"); // Find the location of the Employee in the sheet containing projects
    await context.sync();
    return {first: firstProjCell, last: lastProjCell, colToCheck};
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