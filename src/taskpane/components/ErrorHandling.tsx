import * as React from 'react';

const withErrorHandling = (WrappedComponent) => ({ error, children }) => {
  const colorMessage = 'color-message-red';
  return (
    <WrappedComponent>
      <div className={'error-message ' + colorMessage}>
        {error.displayError ? error.errorMessage : ''}
      </div>
      {children}
    </WrappedComponent>
  );
};

export const ErrorHandling = withErrorHandling(({ children }) => (
  <div>{children}</div>
));
