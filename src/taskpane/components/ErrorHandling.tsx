import * as React from 'react';

export const ErrorHandling: React.FC<{ errorMessage }> = (props) => {
  return <div>waka: {props.errorMessage}</div>;
};
