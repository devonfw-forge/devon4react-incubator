import * as React from 'react';
import { debounce } from 'lodash';

export default class App extends React.Component<{}, {newProj: string, /* newEmployee: string, */ projects: Excel.Range, cellList: any[], dataLoaded: boolean, userName: string}> {
  constructor(props, context) {
    super(props, context);
    this.handleChange = this.handleChange.bind(this);
    this.handleProjName = this.handleProjName.bind(this);
    // this.handleEmployeeName = this.handleEmployeeName.bind(this);
    this.state = {
      projects: null,
      cellList: [],
      userName: null,
      dataLoaded: false,
      newProj: null,
      // newEmployee: null
    };
  }

  componentDidMount() {
    this.clickListener();
  }

  clickListener = async () => {
      await Excel.run(async (context) => {
        context.workbook.worksheets.getFirst().onSelectionChanged.add(this.click);
        await context.sync()
      });
  }

  click = async () => {
    try {
      return Excel.run(async context => {
        this.setState({
          cellList: [],
          projects: null,
          newProj: null,
          dataLoaded: false
        });
        const activeSheet = context.workbook.worksheets.getFirst();
        await activeSheet.activate();
        const range = activeSheet.context.workbook.getSelectedRange().load(["address", "values", "rowIndex"]);
        await context.sync();
        const selectedCatCell = range.address.split("!")[1];
        const selectedCat = activeSheet.getRange(selectedCatCell[0] + "1").load("values");
        const activeRow = range.rowIndex + 1;
        const userAddress = activeSheet.findAll("Employee", {
          completeMatch: true,
          matchCase: false
        }).load("address");
        await context.sync();
        const sheet = context.workbook.worksheets.getItem(selectedCat.values[0][0]);
        const projCol = sheet.findAll("Projects", {
          completeMatch: true,
          matchCase: false
        }).load("address");
        const totalCol = sheet.findAll("Total", {
          completeMatch: true,
          matchCase: false
        }).load("address");
        await context.sync();
        const firstProjPos = sheet.getRange(projCol.address).getRowsBelow(1).load("address");
        const lastProjPos = sheet.getRange(totalCol.address).getRowsAbove(1).load("address");
        await context.sync();
        const firstProjCell = firstProjPos.address.split("!")[1];
        const lastProjCell = lastProjPos.address.split("!")[1];
        this.setState({
          projects: sheet.getRange(firstProjCell + ":" + lastProjCell)
        });
        this.state.projects.load("values");
        const userCol = userAddress.address.split("!")[1][0];
        const activeUser = activeSheet.getRange(userCol + activeRow).load("values");
        await context.sync();
        this.setState({
          userName: activeUser.values[0][0]
        });
        const colToCheck = sheet.findAll(activeUser.values[0][0], {
          completeMatch: true,
          matchCase: false
        }).load("address");
        await context.sync();
        const col = colToCheck.address.split("!")[1][0];
    
        for (let i = 0; i < this.state.projects.values.length; i++) {
          const rowToCheck = sheet.findAll(this.state.projects.values[i][0], {
            completeMatch: true,
            matchCase: false
          }).load("address");
          await context.sync();
          const row = rowToCheck.address.split("!")[1][1];
          const cell = sheet.getRange(col + row).load(["values", "address"]);
          await context.sync();
          this.state.cellList.push({
            value: cell.values[0][0],
            address: cell.address
          });
        }

        this.setState({
          dataLoaded: true
        });

      });
    } catch (error) {
      console.error(error);
    }
  }

  save = debounce(async () => {
    try {
      await Excel.run(async context => {
        let sheetToUpdate = this.state.cellList[0].address.split("!")[0];
        sheetToUpdate = sheetToUpdate.substr(1, sheetToUpdate.length - 2);
        const workSheet = context.workbook.worksheets.getItem(sheetToUpdate);
        for (let cell of this.state.cellList) {
          const cellToUpdate = workSheet.getRange(cell.address.split("!")[1]).load("values");
          await context.sync();
          cellToUpdate.values = [[cell.value]];
        }

      })
    } catch (error) {
      console.error(error);
    }
  }, 200);

  handleProjName(e: any) {
    this.setState({
      newProj: e.target.value
    });
  }

  addProj = async () => {
    if (this.state.newProj !== null) {
      try {
        await Excel.run(async context => {
          let sheetToUpdate = this.state.cellList[0].address.split("!")[0];
          sheetToUpdate = sheetToUpdate.substr(1, sheetToUpdate.length - 2);
          const workSheet = context.workbook.worksheets.getItem(sheetToUpdate);
          const table = workSheet.tables.getItemAt(0);
          const numOfCol = table.columns.getCount();
          await context.sync();
          const hoursVal = [];
          hoursVal.push(this.state.newProj);
          for (let i=1; i<numOfCol.value; i++) {
            hoursVal.push(0);
          }
          table.rows.load("items");
          await context.sync();
          table.rows.add(table.rows.items.length - 1, [hoursVal]);
          await context.sync();
          context.workbook.worksheets.getFirst().activate()
          setTimeout(() => {
            this.click();
          }, 200);
        })
      } catch (error) {
        console.error(error);
      }
    }
  }

  handleChange(e: any, index: any) {
    const newValue = Number.parseInt(e.currentTarget.textContent);
    if (newValue) {
      this.state.cellList[index].value = newValue;
      this.save();
    }
  }

  // addEmployee = async () => {
  //   if (this.state.newEmployee !== null) {
  //     try {
  //       await Excel.run(async context => {
  //         let sheetsToUpdate = context.workbook.worksheets.load("items");
  //         await context.sync();
  //         const firstSheetTable = sheetsToUpdate.getFirst().tables.getItemAt(0);
  //         const numOfCol = firstSheetTable.columns.getCount();
  //         await context.sync();
  //         const totalVal = [];
  //         totalVal.push(this.state.newEmployee);
  //         for (let i=1; i<numOfCol.value; i++) {
  //           totalVal.push(0);
  //         }
  //         // -- NOT WORKING --
  //         sheetsToUpdate.items.slice(1).forEach(async (sheet: any) => {
  //           const table = sheet.tables.getItemAt(0);
  //           const numOfRow = table.rows.getCount();
  //           await context.sync();
  //           const hoursVal = [];
  //           hoursVal.push([this.state.newEmployee]);
  //           for (let i=1; i<numOfRow.value + 1; i++) {
  //             hoursVal.push([0]);
  //           }
  //           table.columns.load("items");
  //           await context.sync();
  //           console.log(table.columns.items);
  //           console.log(hoursVal);
  //           table.columns.add(null, [hoursVal]);
  //           await context.sync();
  //         });
   //         // --
  //         firstSheetTable.rows.add(null, [totalVal]);
  //         sheetsToUpdate.getFirst().activate();
  //         await context.sync();
  //       })
  //     } catch (error) {
  //       console.error(error);
  //     }
  //   }
  // }
  
  // handleEmployeeName(e: any) {
  //   this.setState({
  //     newEmployee: e.target.value
  //   });
  // }

  render() {
    return (
      <div className='ms-welcome'>
        {/* <div>
          <input type="text" placeholder="Employee Name" onChange={ this.handleEmployeeName }/>
          <button className='ms-welcome__action' onClick={this.addEmployee}>Add Employee</button>
        </div> */}
        {/* <div>
          <button className='ms-welcome__action' onClick={this.click}>Get projects data</button>
        </div> */}
        {this.state.dataLoaded &&
        <div>
          <div>
            <input type="text" placeholder="Project Name" onChange={ this.handleProjName }/>
            <button className='ms-welcome__action' onClick={this.addProj}>Add Project</button>
          </div>
          <table className='projectsContainer'>
              <tbody>
                <tr>
                  <th colSpan={2}>{this.state.userName}</th>
                </tr>
                {this.state.projects.values.map((project: any, i: any) => {
                  return (
                    <tr key={i}>
                      <td>{project[0]}</td>
                      <td id={project[0]}>
                        <p suppressContentEditableWarning={true} contentEditable onKeyUp={(event) => this.handleChange(event, i)}>{this.state.cellList[i].value}</p>
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
