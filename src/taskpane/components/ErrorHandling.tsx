import * as React from 'react';

const withErrorHandling = (WrappedComponent) => ({ error, children }) => {
  const colorMessage = 'color-message-' + error.color;
  console.log(error);
  return (
    <WrappedComponent>
      <div className={'error-message ' + colorMessage}>
        {error.showError ? error.errorMessage : ''}
      </div>
      {children}
    </WrappedComponent>
  );
};

export const ErrorHandling = withErrorHandling(({ children }) => (
  <div>{children}</div>
));
