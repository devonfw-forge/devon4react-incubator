import * as React from 'react';

const withErrorHandling = (WrappedComponent) => ({ error, children }) => {
  console.log('Show Errors: ', error.showError);
  return (
    <WrappedComponent>
      {error.showError && (
        <div className="error-message">{error.errorMessage}</div>
      )}
      {children}
    </WrappedComponent>
  );
};

export const ErrorHandling = withErrorHandling(({ children }) => (
  <div>{children}</div>
));
