import * as React from 'react';
import { render } from 'react-dom';

export const ErrorHandler: React.FC<{state}> = (props) => {
    console.log(props);

    return(
        <div>Hello</div>
       
    );
}