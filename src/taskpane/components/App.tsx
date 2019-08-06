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
        try {
          await this.addEventListeners(context);
          await this.click(context);
        } catch (error) {
          console.error(error);
        }
      }),
    );
  }

  addEventListeners = async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    try {
      // Called every time the user click on a cell
      activeSheet.onSelectionChanged.add(
        async () => await this.onSelectionChangedHandler(context),
      );
      // activeSheet.onSelectionChanged.add(() => console.log('waka'));
      // Called every time the user change a value in a cell
      activeSheet.onChanged.add(async () => await this.click(context));
      // Called every time the ADC.DYNACOLUMNS function calculate
      activeSheet.onCalculated.add(
        async () => await this.onCalculatedHandler(context, activeSheet),
      );
      await context.sync();
    } catch (error) {
      console.error(error);
    }
  };

  // TODO: REFACTOR NEEDED!!!!
  onCalculatedHandler = async (context, activeSheet) => {
    setTimeout(async () => {
      const range = activeSheet.context.workbook
        .getSelectedRange()
        .load(['values']); // Get the selected cell location, value and index of its row
      await context.sync();
      if (range.values[0][0] !== CALC) {
        this.setTotal(range.values[0][0]);
      }
    }, 80);
  };

  onSelectionChangedHandler = async (context) => {
    console.log('waka');
    await this.click(context);
  };

  getDefinitions = async (context, columnData): Promise<string[]> => {
    const dataDefinitionsColumns = context.workbook.worksheets
      .getItem(DATA_WORKSHEET)
      .tables.getItemAt(0);

    let dataDefinitionsHeaders = dataDefinitionsColumns
      .getHeaderRowRange()
      .load('values');

    await context.sync();

    dataDefinitionsHeaders = dataDefinitionsHeaders.values
      .filter(String)
      .map((data) => {
        return data[0];
      });

    if (
      columnData !== '' &&
      dataDefinitionsHeaders.indexOf(columnData) === -1
    ) {
      this.setError(true, WORKSHEET_ERRORS.NOT_FOUND, 'red');
      this.setShowTable(false);

      return [''];
    } else {
      const dataDefinitionsValues = dataDefinitionsColumns.columns
        .getItem(columnData)
        .getDataBodyRange()
        .load('values');

      await context.sync();

      return dataDefinitionsValues.values.filter(String).map((data) => {
        return data[0];
      });
    }
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
    try {
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

      const worksheetProxi = context.workbook.worksheets.load('items');
      const employeeNameCell = activeSheet.getRange(cell).load('values');
      await context.sync();

      const sheetsName = worksheetProxi.items.map((sheet) =>
        sheet.name.toLowerCase(),
      );
      if (
        column !== '' &&
        sheetsName.indexOf(DATA_WORKSHEET.toLowerCase()) === -1
      ) {
        this.setError(true, WORKSHEET_ERRORS.NOT_FOUND, 'red');
        this.setShowTable(false);
      }

      if (dataValues.length < dataDefinitions.length) {
        const diference = dataDefinitions.length - dataValues.length;
        for (let i = 0; i < diference; i++) {
          dataValues.push('0');
        }
      } else if (dataValues.length > dataDefinitions.length) {
        this.setError(true, ERRORS.MORE_VALUES, 'yellow');
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

      // TODO: review this validation

      if (total !== CALC) {
        this.setTotal(range.values[0][0]);
      }
      return [employeeNameCell.values, data];
    } catch (error) {
      console.error(error);
    }
  };

  // Get projects' data of the selected Employee
  click = async (context) => {
    try {
      const [employeeName, employeeData] = await this.getEmployeeData(context);

      this.setState((prevState) => {
        let employee = Object.assign({}, prevState.employee);
        employee.name = employeeName;
        employee.worksheetData = employeeData;
        return { employee };
      });

      this.setShowTable(true);
    } catch (error) {
      console.error(error);
    }
  };
  render() {
    return (
      <div className="ms-welcome">
        <ErrorHandling error={this.state.error}>
          {this.state.showTable && (
            <ProjectsPanel
              employee={this.state.employee}
              setError={this.setError.bind(this)}
              setDataLoaded={this.setShowTable.bind(this)}
            />
          )}
        </ErrorHandling>
      </div>
    );
  }
}
