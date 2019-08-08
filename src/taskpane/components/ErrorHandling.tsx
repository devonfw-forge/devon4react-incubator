import * as React from 'react';

export const ErrorHandling: React.FC<{
  message: string;
  color: string;
  children;
}> = ({ message, color, children }) => {
  const colorMessage = 'color-message-' + color;
  return (
    <div>
      <div className={'error-message ' + colorMessage}>{message}</div>
      {children}
    </div>
  );
};

