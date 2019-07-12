import React, { Component } from 'react';
import './App.css';
import readXlsxFile from 'read-excel-file'

export class ExcelRender extends Component<{}, { rows: any; dataLoaded: boolean; headers: any; fileOpened: any; projects: any[]; }> {
  constructor(props: any) {
    super(props);
    this.state = {
      dataLoaded: false,
      rows: null,
      headers: null,
      fileOpened: null,
      projects: [],
    };
  }

  fileHandler = (event: any) => {
    this.setState({
      dataLoaded: false
    });
    let fileObj = event.target.files[0];
    //just pass the fileObj as parameter
    readXlsxFile(fileObj).then((resp: any, i: any) => {
      const data: any = [];
        resp.forEach((row: any, i: any) => {
          data[i] = [];
          row.forEach((employeeData: any) => {
            data[i].push({ value: employeeData });
          });
        });
        this.setState({
          dataLoaded: true,
          rows: resp,
          headers: resp[0],
          fileOpened: fileObj
        });
    });
  };
  
  getEmployeeProjects = (employeeData: any, index: any) => {
    readXlsxFile(this.state.fileOpened, {sheet: this.state.headers[index + 1]}).then((data: any) => {
      const employeeIndex = data[0].indexOf(employeeData[0]);
      const projects: any[] = [];
      data.slice(1, data.length - 1).forEach((row: any) => {
        projects.push({name: row[0], hour: row[employeeIndex]});
      });
      this.setState({
        projects: projects,
      });
    });
  };

  save = () => {

  }

  render() {
    return (
      <div>
        <div>
          <input
            type='file'
            onChange={this.fileHandler.bind(this)}
            style={{ padding: '10px' }}
          />
        </div>
        <div id='dataContainer'>
        {this.state.dataLoaded &&
        <table className='employeesData'>
          <tbody>
          <tr>
              {this.state.headers.map((title: any, i: any) => {
                return <th key={i}>{title}</th>;
              })
            }
          </tr>
            {this.state.rows.slice(1, this.state.rows.length).map((row: any, i: any) => {
              return (
                <tr key={i}>
                  {row.slice(0, 1).map((employee: string, j: any) => {
                    return <th key={j}>{employee}</th>
                  })}
                  {row.slice(1, row.length).map((value: string, j: any) => {
                    return <td key={j}><button className='cell' onClick={() => this.getEmployeeProjects(row, j)}>{value}</button></td>
                  })}
                </tr>
              )
            })
          }
          </tbody>
        </table>
        }
        <div className='projectsContainer'>
        {this.state.projects.length > 0 &&
        <table className='projectsData'>
          <tbody>
            <tr>
              <th>Projects</th>
              <th>Hours</th>
            </tr>
            {this.state.projects.map((project: any, i: any) => {
              return (
                <tr key={i}>
                  <th>{project.name}</th>
                  <td contentEditable>{project.hour}</td>
                </tr>
              )
            })}
          </tbody>
        </table>
        }
        {this.state.projects.length > 0 &&
        <div><button onClick={this.save}>Save</button></div>
      }
      </div>
      </div>

      </div>
    );
  }
}

const App: React.FC = () => {
  return (
    <div className='App'>
      <ExcelRender />
    </div>
  );
};

export default App;
