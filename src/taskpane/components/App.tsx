import * as React from 'react';
import { ProjectsPanel } from './ProjectsPanelComponent';
import { handleOnChange } from './SaveHour';
import { getSelectedEmployeeData } from './SelectedEmployee';
import { EmployeeData } from './shared/model/interfaces/EmployeeData';
import { HoursList } from './shared/model/interfaces/HoursList';
import { ErrorHandling } from './ErrorHandling';

export default class App extends React.Component<
  {},
  {
    projectsSheet: Excel.Worksheet;
    projects: any;
    total: any;
    hoursList: HoursList[];
    dataLoaded: boolean;
    employeeName: string;
    error: {
      showError: boolean;
      errorMessage: string;
      color: string;
    };
    showTable: boolean;
  }
> {
  constructor(props: any, context: Excel.RequestContext) {
    super(props, context);
    handleOnChange.bind(this);

    this.state = {
      projectsSheet: undefined,
      projects: undefined,
      total: undefined,
      hoursList: [],
      employeeName: undefined,
      dataLoaded: false,
      error: {
        showError: true,
        errorMessage: '',
        color: 'white',
      },
      showTable: true,
    };
  }

  setError(showError: boolean, errorMessage: string, color: string) {
    this.setState({
      error: {
        showError: showError,
        errorMessage: errorMessage,
        color: color,
      },
    });
  }
  setShowTable(showTable: boolean) {
    this.setState({
      showTable: showTable,
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
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.onSelectionChanged.add(this.click); // Check if the selected cell has changed
      activeSheet.onChanged.add(this.click); // Check if the selected cell data has changed
      activeSheet.onCalculated.add(this.eventoHandler);
      await context.sync();
    });
  };

  updateTotal = (newTotal) => {
    this.setState({ total: newTotal });
  };

  eventoHandler = async () => {
    Excel.run(async (context) => {
      setTimeout(async () => {
        const activeSheet = context.workbook.worksheets.getActiveWorksheet(); //Get the first Excel sheet
        await activeSheet.activate(); // Activate the first Excel sheet
        const range = activeSheet.context.workbook
          .getSelectedRange()
          .load(['values']); // Get the selected cell location, value and index of its row
        await context.sync();
        if (range.values[0][0] !== '#CALC!') {
          this.updateTotal(range.values[0][0]);
        }
      }, 80);
    });
  };

  // Get projects' data of the selected Employee
  click = async () => {
    // this.setState({
    //   errorMessage: this.state.errorMessage + 'waka',
    // });
    try {
      return Excel.run(async (context) => {
        const employeeData: EmployeeData = {
          category: undefined,
          activeEmployee: undefined,
          data: {
            dataSheet: undefined,
            fte: undefined,
          },
          total: undefined,
        };
        await getSelectedEmployeeData(
          context,
          this.updateTotal,
          this.setError.bind(this),
          this.setShowTable.bind(this),
        ).then((res: any) => {
          employeeData.category = res.selectedCat;
          employeeData.activeEmployee = res.activeEmployee;
          employeeData.data = res.data;
        });

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
            'yellow',
          );
        } else if (projectsValue.length > employeeData.data.fte.length) {
          const diference = projectsValue.length - employeeData.data.fte.length;
          for (let i = 0; i < diference; i++) {
            employeeData.data.fte.push('0');
          }
        }
        if (projectsValue.length >= employeeData.data.fte.length) {
          this.setError(false, '', 'white');
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
          {this.state.dataLoaded && this.state.showTable && (
            <ProjectsPanel state={this.state} />
          )}
        </ErrorHandling>
      </div>
    );
  }
}
