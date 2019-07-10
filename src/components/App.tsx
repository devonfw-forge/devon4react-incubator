import * as React from 'react';

export class App extends React.Component<IProps, IState> {
  // IProps para las PROPIEDADES y IState para el ESTADO

  constructor(props: IProps) {
    super(props);
    this.state = {
      tasks: [],
    };
  }

  render() {
    return (
      <div>
        <h1>{this.props.title}</h1>
      </div>
    );
  }
}

interface IProps {
  title: string;
}

interface IState {
  tasks: [];
}
