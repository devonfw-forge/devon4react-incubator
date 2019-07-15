import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
// import { HeroListItem } from './HeroList';
import Progress from './Progress';

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

// export interface AppState {
//   listItems: HeroListItem[];
// }

export default class App extends React.Component<AppProps, {/* listItems: HeroListItem[], */ projects: Excel.Range, cell: string[], dataLoaded: boolean, userName: string}> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      projects: null,
      cell: null,
      userName: null,
      dataLoaded: false
      // listItems: []
    };
  }

  componentDidMount() {
    // this.setState({
    //   listItems: [
    //     {
    //       icon: 'Ribbon',
    //       primaryText: 'Achieve more with Office integration'
    //     },
    //     {
    //       icon: 'Unlock',
    //       primaryText: 'Unlock features and functionality'
    //     },
    //     {
    //       icon: 'Design',
    //       primaryText: 'Create and visualize like a pro'
    //     }
    //   ]
    // });
  }

  click = async () => {
    try {
      await Excel.run(async context => {
        this.setState({
          cell: null,
          projects: null,
          dataLoaded: false
        });
        const range = context.workbook.getSelectedRange();
        const activeSheet = context.workbook.worksheets.getFirst();
        range.load(["address", "values", "rowIndex"]);
        await context.sync();
        const selectedCatCell = range.address.split("!")[1];
        const selectedCat = activeSheet.getRange(selectedCatCell[0] + "1");
        selectedCat.load("values");
        await context.sync();
        const sheet = context.workbook.worksheets.getItem(selectedCat.values[0][0]);
        const activeRow = range.rowIndex + 1;
        const userAddress = activeSheet.findAll("Employee", {
          completeMatch: true,
          matchCase: false
        });
        userAddress.load("address");
        await context.sync();
        const projCol = sheet.findAll("Projects", {
          completeMatch: true,
          matchCase: false
        });
        projCol.load("address");
        const totalCol = sheet.findAll("Total", {
          completeMatch: true,
          matchCase: false
        });
        totalCol.load("address");
        await context.sync();
        const firstProjCellNum = +projCol.address.split("!")[1][1] + 1;
        const firstProjCell = projCol.address.split("!")[1][0] + firstProjCellNum;
        const lastProjCellNum = +totalCol.address.split("!")[1][1] - 1;
        const lastProjCell = totalCol.address.split("!")[1][0] + lastProjCellNum;
        this.setState({
          projects: sheet.getRange(firstProjCell + ":" + lastProjCell)
        });
        this.state.projects.load("values");
        await context.sync();
        const userCol = userAddress.address.split("!")[1][0];
        const activeUser = activeSheet.getRange(userCol + activeRow);
        activeUser.load("values");
        await context.sync();
        this.setState({
          userName: activeUser.values[0][0]
        });
        const colToCheck = sheet.findAll(activeUser.values[0][0], {
          completeMatch: true,
          matchCase: false
        });
        colToCheck.load("address");
        await context.sync();
        const col = colToCheck.address.split("!")[1][0];
        let cellList: string[] = [];
    
        for (let i = 0; i < this.state.projects.values.length; i++) {
          const rowToCheck = sheet.findAll(this.state.projects.values[i][0], {
            completeMatch: true,
            matchCase: false
          });
          rowToCheck.load("address");
          await context.sync();
          const row = rowToCheck.address.split("!")[1][1];
          const cell = sheet.getRange(col + row);
          cell.load("values");
          await context.sync();
          cellList.push(cell.values[0][0]);
        }

        this.setState({
          cell: cellList,
          dataLoaded: true
        });

      });
    } catch (error) {
      console.error(error);
    }
  }

  render() {
    const {
      title,
      isOfficeInitialized,
    } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo='assets/logo-filled.png'
          message='Please sideload your addin to see app body.'
        />
      );
    }

    return (
      <div className='ms-welcome'>
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
        {this.state.dataLoaded &&
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
                      <p contentEditable>{this.state.cell[i]}</p>
                    </td>
                  </tr>
                  )
              })}
            </tbody>
          </table>
        }
      </div>
    );
  }
}
