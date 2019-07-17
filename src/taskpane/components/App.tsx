/*import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import Header from './Header';
import HeroList, { HeroListItem } from './HeroList';
import Progress from './Progress';

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: 'Ribbon',
          primaryText: 'Achieve more with Office integration'
        },
        {
          icon: 'Unlock',
          primaryText: 'Unlock features and functionality'
        },
        {
          icon: 'Design',
          primaryText: 'Create and visualize like a pro'
        }
      ]
    });
  }

  click = async () => {
    try {
      await Excel.run(async context => {
        
        //Insert your Excel code here
         
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
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
        <Header logo='assets/logo-filled.png' title={this.props.title} message='Welcome' />
        <HeroList message='Discover what Office Add-ins can do for you today!' items={this.state.listItems}>
          <p className='ms-font-l'>Modify the source files, then click <b>Run</b>.</p>
          <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
        </HeroList>
      </div>
    );
  }
}*/
import * as React from 'react';
import { useState } from 'react';
import './App.css';
import { Addin } from '../../shared/model/interfaces/Addin';
import { empleado } from './data';
import { AppContainer } from 'react-hot-loader';
import { Grid } from '../../components/Grid';
import * as ReactDOM from 'react-dom';

export const App: React.FunctionComponent<any> = () => {
  let employee: Addin = empleado;
  if (localStorage.getItem('empleado')) {
    employee = JSON.parse(localStorage.getItem('empleado') || '');
  }
  const [emp, setEmp] = useState(employee);

  const handleClick = () => {
    emp.fields.push({ fieldName: '', fieldValue: '' });
    localStorage.setItem('empleado', JSON.stringify(emp));
    setEmp(emp);
    //ReactDOM.render(<App />, document.getElementById('root'));
  };

  const render = (Component) => {
    ReactDOM.render(
        <AppContainer>
            <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </AppContainer>,
        document.getElementById('container')
    );
  };

  const clearStorage = () => {
    localStorage.clear();
    render(App);
  };

  return (
    <div className="App">
      <header className="App-header">
        <div>
          <h1>Hello {emp.empleado.name}!</h1>

          <Grid employee={emp} setEmp={setEmp} />

          <button onClick={handleClick}>+</button>
          <button onClick={clearStorage}>clear</button>
        </div>
      </header>
    </div>
  );
};
