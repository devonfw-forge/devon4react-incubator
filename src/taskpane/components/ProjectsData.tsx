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

export { getProjectsData };