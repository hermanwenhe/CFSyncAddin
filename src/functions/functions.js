/* global clearInterval, console, setInterval, Excel */

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
export function add(first, second) {
  return first + second;
}


/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 * @returns The value of the cell at the input address.
 **/
export async function getRange(address, invocation) {
  // Retrieve the context object. 
  const context = invocation.getRequestContext();
  
  // Use the context object to access the cell at the input address. 
  const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
  range.load("values");
  await context.sync();
  
  // Return the value of the cell at the input address.
  return range.values[0][0];
}
