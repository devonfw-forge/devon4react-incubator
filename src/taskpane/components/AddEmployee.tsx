// import * as React from 'react';

// export default class AddEmployee extends React.Component<{}, {newEmployee: string}> {

    // constructor(props, context) {
        // super(props, context);
        // this.handleEmployeeName = this.handleEmployeeName.bind(this);
        // this.state = {
          // newEmployee: null
        // };
    //   }

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

//   render() {
    // return (
        {/* <div>
          <input type="text" placeholder="Employee Name" onChange={ this.handleEmployeeName }/>
          <button className='ms-welcome__action' onClick={this.addEmployee}>Add Employee</button>
        </div> */}
        {/* <div>
          <button className='ms-welcome__action' onClick={this.click}>Get projects data</button>
        </div> */}
    // );
//   }

// }