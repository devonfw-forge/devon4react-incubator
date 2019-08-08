import * as React from 'react';
import { ErrorHandling } from './ErrorHandling';
import { ProjectsPanel } from './ProjectsPanel';
import {
  CALC,
  ERRORS,
  WORKSHEET_ERRORS,
  DATA_WORKSHEET,
  COLUMN_NOT_FOUND,
  INVALID_FORMULA,
  HEAD_FORMULA,
} from './shared/constant';
import { Employee } from './shared/model/interfaces/Employee';
import { TableError } from './shared/model/interfaces/Error';

interface isState {
  employee: Employee;
  error: TableError[];
  showTable: boolean;
}

export default class App extends React.Component<{}, isState> {
  errorMessage: string;
  errorColor: string;
  constructor(props: any, context: Excel.RequestContext) {
    super(props, context);

    this.state = {
      employee: {
        name: '',
        worksheetData: [],
        cell: '',
        column: '',
        total: 0,
      },
      error: [
        // 0
        {
          showError: false,
          errorMessage: ERRORS.INCORRECT_CELL,
          color: 'green',
        },
        // 1
        {
          showError: false,
          errorMessage: ERRORS.MORE_VALUES,
          color: 'yellow',
        },
        // 2
        {
          showError: false,
          errorMessage: ERRORS.INCORRECT_VALUE,
          color: 'red',
        },
        // 3
        {
          showError: false,
          errorMessage: WORKSHEET_ERRORS.EMPTY,
          color: 'red',
        },
        // 4
        {
          showError: false,
          errorMessage: WORKSHEET_ERRORS.NOT_FOUND,
          color: 'red',
        },
        // 5
        {
          showError: false,
          errorMessage: COLUMN_NOT_FOUND,
          color: 'red',
        },
        // 6
        {
          showError: false,
          errorMessage: INVALID_FORMULA,
          color: 'red',
        },
      ],
      showTable: true,
    };
  }

  setEmployeeData(data: number, idx: number) {
    this.setState((prevState) => {
      let employee = Object.assign({}, prevState.employee);
      employee.worksheetData[idx].value = data;
      return { employee };
    });
  }

  setDataError(error: boolean, idx: number) {
    this.setState(
      (prevState) => {
        let employee = Object.assign({}, prevState.employee);
        employee.worksheetData[idx].error = error;
        return { employee };
      },
      () => {
        if (
          this.state.employee.worksheetData.find((data) => data.error === true)
        ) {
          this.setError(true, 2);
        } else {
          this.setError(false, 2);
          this.save();
        }
      },
    );
  }

  setError(showError: boolean, idx: number) {
    this.setState(
      (prevState) => {
        let error = prevState.error.map((error) => Object.assign({}, error));
        error[idx].showError = showError;
        return { error };
      },
      () => {
        if (this.state.error[6].showError) {
          // Invalid formula
          this.errorMessage = this.state.error[6].errorMessage;
          this.errorColor = this.state.error[6].color;
          this.setShowTable(false);
        } else if (this.state.error[0].showError) {
          // Incorrect cell error
          this.errorMessage = this.state.error[0].errorMessage;
          this.errorColor = this.state.error[0].color;
          this.setShowTable(false);
        } else if (this.state.error[5].showError) {
          // Columns not found
          this.errorMessage = this.state.error[5].errorMessage;
          this.errorColor = this.state.error[5].color;
          this.setShowTable(false);
        } else if (this.state.error[4].showError) {
          // Worksheet not found
          this.errorMessage = this.state.error[4].errorMessage;
          this.errorColor = this.state.error[4].color;
          this.setShowTable(false);
        } else if (this.state.error[3].showError) {
          // Worksheet empty
          this.errorMessage = this.state.error[3].errorMessage;
          this.errorColor = this.state.error[3].color;
          this.setShowTable(false);
        } else if (this.state.error[2].showError) {
          // Incorrect value error
          this.errorMessage = this.state.error[2].errorMessage;
          this.errorColor = this.state.error[2].color;
          this.setShowTable(true);
        } else if (this.state.error[1].showError) {
          // To many values error
          this.errorMessage = this.state.error[1].errorMessage;
          this.errorColor = this.state.error[1].color;
          this.setShowTable(true);
        } else {
          // No errors
          this.errorMessage = '';
          this.errorColor = 'white';
          this.setShowTable(true);
        }
      },
    );
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
          const activeSheet = context.workbook.worksheets.getActiveWorksheet();
          await this.addEventListeners(context, activeSheet);
          await this.click(context, activeSheet);
        } catch (error) {
          console.error(error);
        }
      }),
    );
  }

  addEventListeners = async (context, activeSheet) => {
    try {
      // Called every time the user click on a cell
      activeSheet.onSelectionChanged.add(
        async () => await this.click(context, activeSheet),
      );
      // Called every time the user change a value in a cell
      activeSheet.onChanged.add(
        async () => await this.click(context, activeSheet),
      );
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

  getDefinitions = async (context, columnData): Promise<string[]> => {
    try {
      const dataDefinitionsColumns = context.workbook.worksheets
        .getItem(DATA_WORKSHEET)
        .tables.getItemAt(0);

      let dataDefinitionsHeaders = dataDefinitionsColumns
        .getHeaderRowRange()
        .load('values');

      await context.sync();

      dataDefinitionsHeaders = dataDefinitionsHeaders.values[0];
      dataDefinitionsHeaders = dataDefinitionsHeaders.map((header) =>
        header.toLowerCase(),
      );

      if (columnData === '') {
        this.setError(true, 5);
        return [];
      } else if (
        columnData !== '' &&
        dataDefinitionsHeaders.indexOf(columnData) === -1
      ) {
        // Set error column not found
        this.setError(true, 5);

        return [];
      } else {
        // Remove error column not found
        this.setError(false, 5);
        const dataDefinitionsValues = dataDefinitionsColumns.columns
          .getItem(columnData)
          .getDataBodyRange()
          .load('values');

        await context.sync();

        return dataDefinitionsValues.values.filter(String).map((data) => {
          return data[0];
        });
      }
    } catch (error) {
      console.error(error);
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
        // Set error Incorrect Cell
        this.setError(true, 0);
      } else {
        this.setError(false, 0);
      }

      const [cell, column, dataValues] = await this.parseFormula(formula);

      const dataDefinitions = await this.getDefinitions(
        context,
        column.toLowerCase(),
      );

      const worksheetProxi = context.workbook.worksheets.load('items');
      const employeeNameCell = activeSheet.getRange(cell).load('values');
      await context.sync();

      const sheetsName = worksheetProxi.items.map((sheet) =>
        sheet.name.toLowerCase(),
      );
      if (sheetsName.indexOf(DATA_WORKSHEET.toLowerCase()) === -1) {
        // Set error worksheet doesn't exist
        this.setError(true, 4);
      } else {
        this.setError(false, 4);
      }

      if (dataValues.length < dataDefinitions.length) {
        const diference = dataDefinitions.length - dataValues.length;
        for (let i = 0; i < diference; i++) {
          dataValues.push('0');
        }
      } else if (dataValues.length > dataDefinitions.length) {
        // Set error Too many values in the formula
        this.setError(true, 1);
      }

      if (dataValues.length <= dataDefinitions.length) {
        this.setError(false, 1);
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
      return [employeeNameCell.values, cell, data, column];
    } catch (error) {
      console.error(error);
    }
  };

  save = async () => {
    try {
      await Excel.run(async (context) => {
        const activeSheet = context.workbook.worksheets.getFirst(); // Get the Excel sheet to update
        const cellToUpdate = activeSheet.context.workbook
          .getSelectedRange()
          .load(['address', 'values', 'rowIndex', 'formulas']);
        await context.sync();

        let formula = `${HEAD_FORMULA}${this.state.employee.cell},"${
          this.state.employee.column
        }",{`;
        this.state.employee.worksheetData.map((data, idx, arr) => {
          idx === arr.length - 1
            ? (formula = formula + data.value + '})')
            : (formula = formula + data.value + ';');
        });
        cellToUpdate.formulas = [[formula]];
      });
    } catch (error) {
      console.error(error);
    }
  };

  // Get projects' data of the selected Employee
  click = async (context, activeSheet) => {
    this.setState(
      (prevState) => {
        let error = prevState.error.map((error) => Object.assign({}, error));
        error = error.map((error) => {
          error.showError = false;
          return error;
        });

        return { error };
      },
      () => {},
    );
    try {
      const selectCell = activeSheet.context.workbook
        .getSelectedRange()
        .load(['values', 'formulas']); // Get the selected cell location, value and index of its row
      await context.sync();

      const checkFormula = new RegExp('^=ADC.DYNACOLUMNS(.*)', 'gmi');
      if (!checkFormula.test(selectCell.formulas)) {
        // Set error Incorrect Cell
        this.setError(true, 0);
      } else {
        this.setError(false, 0);

        if (selectCell.values[0][0] === '#VALUE!') {
          this.setError(true, 6);
        } else if (selectCell.formulas[0][0] === '') {
          this.setError(false, 6);
          this.setError(true, 0);
        } else {
          this.setError(false, 6);
          this.setError(false, 0);
          const [
            employeeName,
            employeeCell,
            employeeData,
            employeeColumn,
          ] = await this.getEmployeeData(context);

          this.setState((prevState) => {
            let employee = Object.assign({}, prevState.employee);
            employee.name = employeeName;
            employee.worksheetData = employeeData;
            employee.cell = employeeCell;
            employee.column = employeeColumn;
            return { employee };
          });
        }
      }
    } catch (error) {
      console.error(error);
    }
  };
  render() {
    return (
      <div className="ms-welcome">
        <ErrorHandling message={this.errorMessage} color={this.errorColor}>
          {this.state.showTable && (
            <ProjectsPanel
              employee={this.state.employee}
              setDataEmployee={this.setEmployeeData.bind(this)}
              setDataError={this.setDataError.bind(this)}
            />
          )}
        </ErrorHandling>
      </div>
    );
  }
}
