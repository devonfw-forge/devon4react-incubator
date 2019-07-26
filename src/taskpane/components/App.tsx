import * as React from 'react';
// import { AddProject } from './AddProjectComponent';
import { getSelectedEmployeeData } from './SelectedEmployee';
import { handleHourChange } from './SaveHour';
import { HoursList } from './shared/model/interfaces/HoursList';
import { EmployeeData } from './shared/model/interfaces/EmployeeData';
import { ProjectsPanel } from './ProjectsPanelComponent';
// import { ErrorHandler } from './ErrorHandlerComponent';
export default class App extends React.Component<
  {},
  {
    projectsSheet: Excel.Worksheet;
    projects: Excel.Range;
    total: any;
    hoursList: HoursList[];
    dataLoaded: boolean;
    employeeName: string;
    totalSet: boolean;
  }
> {
  constructor(props: any, context: Excel.RequestContext) {
    super(props, context);
    handleHourChange.bind(this);

    this.state = {
      projectsSheet: undefined,
      projects: undefined,
      total: undefined,
      hoursList: [],
      employeeName: undefined,
      dataLoaded: false,
      totalSet: false,
    };
  }

  // Called once the page is loaded and the components are ready
  componentDidMount() {
    Office.onReady(info => {
      this.clickListener();
      this.click();
    });
  }

  // Called every time the user click on a cell
  clickListener = async () => {
    await Excel.run(async context => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.onSelectionChanged.add(this.click); // Check if the selected cell has changed
      activeSheet.onChanged.add(this.click); // Check if the selected cell data has changed
      activeSheet.onCalculated.add(this.eventoHandler);
      await context.sync();
    });
  };

  updateTotal = newTotal => {
    this.setState({ total: newTotal });
  };

  eventoHandler = async event => {
    Excel.run(async context => {
      setTimeout(async () => {
        const activeSheet = context.workbook.worksheets.getActiveWorksheet(); //Get the first Excel sheet
        await activeSheet.activate(); // Activate the first Excel sheet
        const range = activeSheet.context.workbook
          .getSelectedRange()
          .load(['values']); // Get the selected cell location, value and index of its row
        await context.sync();

        this.updateTotal(range.values[0][0]);
      }, 80);

      // this.setState({ total: range.values[0][0] });
    });
  };

  // Get projects' data of the selected Employee
  click = async () => {
    try {
      return Excel.run(async context => {
        this.setState({
          projectsSheet: undefined,
          projects: undefined,
          hoursList: [],
          dataLoaded: false
        }); // Reset state to empty / false

        const employeeData: EmployeeData = {
          category: undefined,
          activeEmployee: undefined,
          data: undefined,
          total: undefined
        };
        await getSelectedEmployeeData(context, this.state.totalSet, this.updateTotal).then(
          (res: any) => {
            employeeData.category = res.selectedCat;
            employeeData.activeEmployee = res.activeEmployee;
            employeeData.data = res.data;
          }
        );

        this.setState({
          totalSet: true,
        })

        const projectsCol = context.workbook.worksheets
          .getItem(employeeData.data[0])
          .tables.getItemAt(0)
          .columns.load('items');

        await context.sync();
        const projects: string[][] = projectsCol.items[0].values.slice(
          1,
          projectsCol.items[0].values.length
        ); //todo -> get data table sin headers
        const proj: any = [];
        employeeData.data
          .slice(1, employeeData.data.length)
          .map((hour: any, i: number) => {
            proj.push({ name: projects[i][0], hours: hour });
          });

        this.setState({
          projects: proj, // Set the state projects with the projects from the sheet with their data
          employeeName: employeeData.activeEmployee.values[0][0], // Set the state name with the selected Employee
          dataLoaded: true // Set the state dataLoaded to true once the data is ready to be displayed
          // total: employeeData.total
        });
      });
    } catch (error) {
      console.error(error);
    }
  };
  render() {
    return (
      <div className="ms-welcome">
        {/* <ErrorHandler state={this.state}/> */}
        {this.state.dataLoaded && (
          <div>
            {/* <AddProject
              state={this.state}
              projSheet={this.state.projectsSheet}
              click={this.click}
            /> */}
            <ProjectsPanel state={this.state} />
          </div>
        )}
      </div>
    );
  }
}
