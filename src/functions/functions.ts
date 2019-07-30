/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
function increment(
  incrementBy: number,
  invocation: CustomFunctions.StreamingInvocation<number>,
): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
function logMessage(message: string): string {
  console.log(message);

  return message;
}

/**
 * Render the project fte for this employee correlated to the project defined in worksheet.
 * @customfunction
 * @param employee Employee cell reference
 * @param worksheet Worksheet name where the project are defined
 * @param fte Fte for this employee
 * @returns Total FTE for this employee
 */
function render(employee: any, worksheet: string, fte: number[][]): number {
  const parsedFte = fte.map((fte) => {
    return fte[0];
  });
  return parsedFte.reduce((acumulated, value) => acumulated + value);
}
