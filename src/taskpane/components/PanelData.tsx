const setPanelData = async (context: Excel.RequestContext, projData, state: any) => {
    const employeeCol = projData.employeeCol.address.split("!")[1][0]; // Get the Column of the Employee in the sheet containing projects    
    for (let i = 0; i < state.projects.values.length; i++) { // Go through all projects one by one
        const projectCell = state.projectsSheet.findAll(state.projects.values[i][0], {
        completeMatch: true,
        matchCase: false
        }).load("address"); // Find the location of the project in the Projects sheet
        await context.sync();
        const projectRow = projectCell.address.split("!")[1][1]; // Get the Row of project
        const hourCell = state.projectsSheet.getRange(employeeCol + projectRow).load(["values", "address"]); // Get the number of hours the Employee has done in this project and its cell location
        await context.sync();
        state.hoursList.push({
        value: hourCell.values[0][0],
        address: hourCell.address
        }); // Set the state hoursList done for each project and the location of its cell in the Projects sheet
    }
    
};

export { setPanelData };