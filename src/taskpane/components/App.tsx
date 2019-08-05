import * as React from 'react';
import { ErrorHandling } from './ErrorHandling';
import { ProjectsPanel } from './ProjectsPanelComponent';
// import { handleOnChange } from './SaveHour';
import { getSelectedEmployeeData } from './SelectedEmployee';
import {
  CALC,
  ERRORS,
  WORKSHEET_ERRORS,
  DATA_WORKSHEET,
} from './shared/constant';
import { EmployeeData } from './shared/model/interfaces/EmployeeData';
import { ProjectData } from './shared/model/interfaces/ProjectData';

export default class App extends React.Component<
  {},
  {
    employee: {
      name: string;
      cell: string;
      worksheetData: ProjectData[];
      total: number;
    };
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
    // handleOnChange.bind(this);

    this.state = {
      employee: {
        name: '',
        cell: '',
        worksheetData: [],
        total: 0,
      },
      error: {
        showError: false,
        errorMessage: '',
        color: 'white',
      },
      showTable: false,
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

  setTotal = (total: number) => {
    this.setState((prevState) => {
      let employee = Object.assign({}, prevState.employee);
      employee.total = total;
      return { employee };
    });
  };

  // Called once the page is loaded and the components are ready
  componentDidMount() {
    Office.onReady(() =>
      Excel.run(async (context) => {
        await this.addEventListeners(context);
        await this.click(context);
      }),
    );
  }

  addEventListeners = async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    // Called every time the user click on a cell
    activeSheet.onSelectionChanged.add(await this.click(context));
    // Called every time the user change a value in a cell
    activeSheet.onChanged.add(await this.click(context));
    // Called every time the ADC.DYNACOLUMNS function calculate
    activeSheet.onCalculated.add(await this.onCalculatedHandler);
    await context.sync();
  };

  // TODO: REFACTOR NEEDED!!!!
  onCalculatedHandler = async () => {
    Excel.run(async (context) => {
      setTimeout(async () => {
        const activeSheet = context.workbook.worksheets.getActiveWorksheet(); //Get the active Excel sheet
        const range = activeSheet.context.workbook
          .getSelectedRange()
          .load(['values']); // Get the selected cell location, value and index of its row
        await context.sync();
        if (range.values[0][0] !== CALC) {
          this.setTotal(range.values[0][0]);
        }
      }, 80);
    });
  };

  getDefinitions = async (context, columnData): Promise<string[]> => {
    const dataDefinitions = context.workbook.worksheets
      .getItem(DATA_WORKSHEET)
      .tables.getItemAt(0)
      .columns.getItem(columnData)
      .getDataBodyRange()
      .load('values');
    await context.sync();
    return dataDefinitions.values.filter(String).map((data) => {
      return data[0];
    });
  };

  parseFormula = (formula: string): [string, string, string[]] => {
    // return [cell, column, [values]]
    let parsedFormula = formula.split('(')[1].split(',');
    parsedFormula[2] = parsedFormula[2].substring(
      1,
      parsedFormula[2].length - 2,
    );
    parsedFormula[1] = parsedFormula[1].substring(
      1,
      parsedFormula[1].length - 1,
    );
    return [parsedFormula[0], parsedFormula[1], parsedFormula[2].split(';')];
  };

  getEmployeeData = async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet(); //Get the active Excel sheet
    const range = activeSheet.context.workbook
      .getSelectedRange()
      .load(['formulas', 'values']); // Get the selected cell location, value and index of its row
    await context.sync();

    const formula = range.formulas[0][0];
    const total = range.values[0][0];

    const checkFormula = new RegExp('^=ADC.DYNACOLUMNS(.*)', 'gmi');
    if (!checkFormula.test(formula)) {
      this.setError(true, ERRORS.INCORRECT_CELL, 'green');
      this.setShowTable(false);
    } else {
      this.setShowTable(true);
    }

    const [cell, column, dataValues] = await this.parseFormula(formula);

    if (column === '') {
      this.setError(true, WORKSHEET_ERRORS.EMPTY, 'red');
      this.setShowTable(false);
    }

    const dataDefinitions = await this.getDefinitions(context, column);

    if (dataValues.length < dataDefinitions.length) {
      this.setError(true, ERRORS.MORE_VALUES, 'yellow');
    } else if (dataValues.length > dataDefinitions.length) {
      const diference = dataValues.length - dataDefinitions.length;
      for (let i = 0; i < diference; i++) {
        dataValues.push('0');
      }
    }

    if (dataValues.length >= dataDefinitions.length) {
      this.setError(false, '', 'white');
    }

    const data = dataDefinitions.map((definition: string, idx: number) => {
      return {
        name: definition,
        value: dataValues[idx],
      };
    });

    context.workbook.worksheets.load('items');

    const employeeNameCell = activeSheet.getRange(cell).load('values');
    await context.sync();

    const sheetsName = context.workbook.worksheets.items.map((sheet) => {
      sheet.name.toLowerCase();
    });

    // TODO: review this validation
    if (column !== '' && sheetsName.indexOf(column) === -1) {
      this.setError(true, WORKSHEET_ERRORS.NOT_FOUND, 'red');
      this.setShowTable(false);
    }

    if (total !== CALC) {
      this.setTotal(range.values[0][0]);
    }
    return [employeeNameCell.values, data];
  };

  // Get projects' data of the selected Employee
  click = async (context) => {
    try {
      // const employeeData: EmployeeData = {
      //   activeEmployee: undefined,
      //   data: {
      //     employeeCell: undefined,
      //     dataSheet: undefined,
      //     value: undefined,
      //   },
      // };

      // await this.getEmployeeData(context).then((res: any) => {
      //   employeeData.activeEmployee = res.activeEmployee.values[0][0];
      //   employeeData.data = res.data;
      // });

      const [employeeName, employeeData] = await this.getEmployeeData(context);

      console.log('log Employee: ', employeeName, employeeData);

      // const projectsCol = context.workbook.worksheets
      //   .getItem(employeeData.data.dataSheet)
      //   .tables.getItemAt(0)
      //   .columns.load('items');

      // await context.sync();
      // let projectsValue = projectsCol.items[0].values.slice(
      //   1,
      //   projectsCol.items[0].values.length,
      // );

      // if (projectsValue.length < employeeData.data.value.length) {
      //   this.setError(true, ERRORS.MORE_VALUES, 'yellow');
      // } else if (projectsValue.length > employeeData.data.value.length) {
      //   const diference = projectsValue.length - employeeData.data.value.length;
      //   for (let i = 0; i < diference; i++) {
      //     employeeData.data.value.push('0');
      //   }
      // }

      // if (projectsValue.length >= employeeData.data.value.length) {
      //   this.setError(false, '', 'white');
      // }

      // const proj: ProjectData[] = projectsValue.map(
      //   (project: string[], idx: number) => {
      //     return {
      //       name: project[0],
      //       value: employeeData.data.value[idx],
      //     };
      //   },
      // );

      // this.setState((prevState) => {
      //   let employee = Object.assign({}, prevState.employee);
      //   employee.worksheetData = proj;
      //   employee.name = employeeData.activeEmployee;
      //   employee.cell = employeeData.data.employeeCell;
      //   return { employee };
      // });
    } catch (error) {
      console.error(error);
    }
  };
  render() {
    return (
      <div className="ms-welcome">
        <ErrorHandling error={this.state.error}>
          {/* {this.state.showTable && (
            <ProjectsPanel
              state={this.state}
              setError={this.setError.bind(this)}
              setDataLoaded={this.setDataLoaded.bind(this)}
            />
          )} */}
        </ErrorHandling>
      </div>
    );
  }
}
