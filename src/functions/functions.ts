/**
 * Render the project fte for this employee correlated to the project defined in worksheet.
 * @customfunction
 * @param employee Employee cell reference
 * @param worksheet Worksheet name where the project are defined
 * @param fte Fte for this employee
 * @returns Total FTE for this employee
 */
function dynaColumns(
  employee: any,
  worksheet: string,
  fte: number[][],
): number {
  const parsedFte = fte.map((fte) => {
    return fte[0];
  });
  return parsedFte.reduce((acumulated, value) => acumulated + value);
}
