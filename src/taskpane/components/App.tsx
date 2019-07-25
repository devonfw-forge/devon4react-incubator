import * as React from 'react';
import { AddProject } from './AddProjectComponent';
import { getSelectedEmployeeData } from './SelectedEmployee';
import { handleHourChange } from './SaveHour';
import { HoursList } from './shared/model/interfaces/HoursList';
import { EmployeeData } from './shared/model/interfaces/EmployeeData';
import { ProjectsPanel } from './ProjectsPanelComponent';

export default class App extends React.Component<
  {},
  {
    projectsSheet: Excel.Worksheet;
    projects: Excel.Range;
    hoursList: HoursList[];
    dataLoaded: boolean;
    employeeName: string;
  }
> {
  constructor(props: any, context: Excel.RequestContext) {
    super(props, context);
    handleHourChange.bind(this);
    this.state = {
      projectsSheet: undefined,
      projects: undefined,
      hoursList: [],
      employeeName: undefined,
      dataLoaded: false,
    };
  }

  // Called once the page is loaded and the components are ready
  componentDidMount() {
    Office.onReady((info) => {
      this.clickListener();
      this.click();
    });
  }

  // Called every time the user click on a cell
  clickListener = async () => {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getFirst().onSelectionChanged.add(this.click); // Check if the selected cell has changed
      await context.sync();
    });
  };

  // Get projects' data of the selected Employee
  click = async () => {
    try {
      return Excel.run(async (context) => {
        this.setState({
          projectsSheet: undefined,
          projects: undefined,
          hoursList: [],
          dataLoaded: false,
        }); // Reset state to empty / false

        const employeeData: EmployeeData = {
          category: undefined,
          activeEmployee: undefined,
          data: undefined,
        };
        await getSelectedEmployeeData(context).then((res: any) => {
          employeeData.category = res.selectedCat;
          employeeData.activeEmployee = res.activeEmployee;
          employeeData.data = res.data;
        });

        const projectsCol = context.workbook.worksheets
          .getItem(employeeData.data[0])
          .tables.getItemAt(0)
          .columns.load('items');

        await context.sync();
        const projects: string[][] = projectsCol.items[0].values.slice(
          1,
          projectsCol.items[0].values.length,
        ); //todo -> get data table sin headers
        const proj: any = [];
        employeeData.data
          .slice(1, employeeData.data.length)
          .map((hour: any, i: number) => {
            proj.push({ name: projects[i][0], hours: hour });
          });

        this.setState({
          projects: proj, // Set the state projects with the projects from the sheet with their data
        });

        this.setState({
          employeeName: employeeData.activeEmployee.values[0][0], // Set the state name with the selected Employee
          dataLoaded: true, // Set the state dataLoaded to true once the data is ready to be displayed
        });
      });
    } catch (error) {
      console.error(error);
    }
  };
  render() {
    return (
      <div className="ms-welcome">
        {this.state.dataLoaded && (
          <div>
            <AddProject
              state={this.state}
              projSheet={this.state.projectsSheet}
              click={this.click}
            />
            <ProjectsPanel state={this.state} />
          </div>
        )}
      </div>
    );
  }
}
