import * as React from 'react';

const withErrorHandling = (WrappedComponent) => ({ error, children }) => {
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
