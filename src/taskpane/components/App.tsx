import * as React from 'react';
import AddProject from './AddProject';
import { getSelectedEmployeeData } from './SelectedEmployee';
import { getProjectsData } from './ProjectsData';
import { setPanelData } from './PanelData';
import { handleHourChange } from './SaveHour';

interface HoursList {
  value: Number;
  address: string;
}

interface EmployeeData {
  category: any;
  activeEmployee: any;
}

interface ProjData {
  firstCell: any;
  lastCell: any;
  colToCheck: any;
}

export default class App extends React.Component<{}, {
  projectsSheet: Excel.Worksheet,
  projects: Excel.Range,
  hoursList: HoursList[],
  dataLoaded: boolean,
  employeeName: string,
}> {
  constructor(props: any, context: Excel.RequestContext) {
    super(props, context);
    handleHourChange.bind(this);
    this.state = {
      projectsSheet: null,
      projects: null,
      hoursList: [],
      employeeName: null,
      dataLoaded: false,
    };
  }

  // Called once the page is loaded and the components are ready
  componentDidMount() {
    this.clickListener();
  }

  // Called every time the user click on a cell
  clickListener = async () => {
      await Excel.run(async (context) => {
        context.workbook.worksheets.getFirst().onSelectionChanged.add(this.click); // Check if the selected cell has changed
        await context.sync()
      });
  }

  // Get projects' data of the selected Employee
  click = async () => {
    try {
      return Excel.run(async context => {
        this.setState({
          projectsSheet: null,
          projects: null,
          hoursList: [],
          dataLoaded: false
        }); // Reset state to empty / false

        const employeeData: EmployeeData = {
          category: null,
          activeEmployee: null
        };
        await getSelectedEmployeeData(context).then((res: any) => {
          employeeData.category = res.selectedCat;
          employeeData.activeEmployee = res.activeUser;
        });

        this.setState({
          projectsSheet: context.workbook.worksheets.getItem(employeeData.category.values[0][0]).load("name") // Get the Excel sheet containing projects and their data 
        });

        const projData: ProjData = {
          firstCell: null,
          lastCell: null,
          colToCheck: null
        }
        await getProjectsData(context, employeeData, this.state).then((res: any) => {
          projData.firstCell = res.first;
          projData.lastCell = res.last;
          projData.colToCheck = res.colToCheck;
        });

        this.setState({
          projects: this.state.projectsSheet.getRange(projData.firstCell + ":" + projData.lastCell).load("values") // Set the state projects with the projects from the sheet with their data
        });
        await context.sync();

        await setPanelData(context, projData, this.state);

        this.setState({
          employeeName: employeeData.activeEmployee.values[0][0], // Set the state name with the selected Employee
          dataLoaded: true // Set the state dataLoaded to true once the data is ready to be displayed
        });

      });
    } catch (error) {
      console.error(error);
    }
  }

  render() {
    return (
      <div className='ms-welcome'>
        {this.state.dataLoaded &&
        <div>
          <AddProject projSheet={this.state.projectsSheet} click={this.click} ></AddProject>
          <table className='projectsContainer'>
              <tbody>
                <tr>
                  <th colSpan={2}>{this.state.employeeName}</th>
                </tr>
                {this.state.projects.values.map((project: string[], i: number) => {
                  return (
                    <tr key={i}>
                      <td>{project[0]}</td>
                      <td id={project[0]}>
                        <p suppressContentEditableWarning={true} contentEditable onKeyUp={(event) => handleHourChange(event, i, this.state)}>{this.state.hoursList[i].value}</p>
                      </td>
                    </tr>
                    )
                  })}
              </tbody>
            </table>
          </div>
        }
      </div>
    );
  }
}
