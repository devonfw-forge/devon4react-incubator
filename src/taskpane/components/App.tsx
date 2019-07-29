import * as React from 'react';
import { AddProject } from './AddProjectComponent';
import { getSelectedEmployeeData } from './SelectedEmployee';
import { handleHourChange } from './SaveHour';
import { HoursList } from './shared/model/interfaces/HoursList';
import { EmployeeData } from './shared/model/interfaces/EmployeeData';
import { ProjectsPanel } from './ProjectsPanelComponent';
import { ErrorHandling } from './ErrorHandling';

export default class App extends React.Component<
  {},
  {
    projectsSheet: Excel.Worksheet;
    projects: any;
    hoursList: HoursList[];
    dataLoaded: boolean;
    employeeName: string;
    error: {
      showError: boolean;
      errorMessage: string;
    };
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
      error: {
        showError: true,
        errorMessage: '',
      },
    };
  }

  setError(showError: boolean, errorMessage: string) {
    this.setState({
      error: {
        showError: showError,
        errorMessage: errorMessage,
      },
    });
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
    // this.setState({
    //   errorMessage: this.state.errorMessage + 'waka',
    // });
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
          data: {
            dataSheet: undefined,
            fte: undefined,
          },
        };
        await getSelectedEmployeeData(context, this.setError.bind(this)).then(
          (res: any) => {
            employeeData.category = res.selectedCat;
            employeeData.activeEmployee = res.activeEmployee;
            employeeData.data = res.data;
          },
        );

        const projectsCol = context.workbook.worksheets
          .getItem(employeeData.data.dataSheet)
          .tables.getItemAt(0)
          .columns.load('items');

        await context.sync();
        let projectsValue = projectsCol.items[0].values.slice(
          1,
          projectsCol.items[0].values.length,
        ); //todo -> get data table sin headers

        if (projectsValue.length < employeeData.data.fte.length) {
          this.setError(
            true,
            'You specified more values than definitions for this employee',
          );
        } else if (projectsValue.length > employeeData.data.fte.length) {
          const diference = projectsValue.length - employeeData.data.fte.length;
          for (let i = 0; i < diference; i++) {
            employeeData.data.fte.push('0');
          }
        }
        if (projectsValue.length >= employeeData.data.fte.length) {
          this.setError(false, '');
        }

        const proj = projectsValue.map((project: any, idx: number) => {
          return {
            name: project[0],
            hours: employeeData.data.fte[idx],
          };
          // proj.push({ name: projects[i][0], hours: hour });
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
        <ErrorHandling error={this.state.error}>
          {this.state.dataLoaded && <ProjectsPanel state={this.state} />}
        </ErrorHandling>
      </div>
    );
  }
}
