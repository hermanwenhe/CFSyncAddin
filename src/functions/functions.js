/* global clearInterval, console, setInterval, Excel */

/**
 * Add two numbers
 * @customfunction
 * @supportSync
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
export function add(first, second) {
  return first + second;
}

/**
 * @customfunction
 * @supportSync
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
export async function getRange(address) {
  // Retrieve the context object. 
  const context = new Excel.RequestContext()
  context.setInvocation(invocation);
  
  // Use the context object to access the cell at the input address. 
  const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
  range.load("values");
  await context.sync();
  
  // Return the value of the cell at the input address.
  return range.values[0][0];
}

/**
 * @customfunction
 * @supportSync
 * @param {string} address The address of the cell from which to retrieve the value.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns The value of the cell at the input address.
 **/
export async function getRange1(address, invocation) {
  // New context to execute API calls synchronously.
  const context = invocation.getRequestContext();
  context.setInvocation(invocation);
  
  try {
    // Use the context object to access the cell at the input address. 
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    range.load("values");
    await context.sync();
  } catch (e) {
    // Handle the error if the address is invalid or the cell is not found.
    return `Error: ${e.message}`;
  }
  
  // Return the value of the cell at the input address.
  return range.values[0][0];
}



/**
 * @customfunction
 * @supportSync
 * @param {string} address The address of the cell from which to retrieve the value.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns The value of the cell at the input address.
 **/
export async function getRange2(address, invocation) {
  // New context to execute API calls synchronously.
  const context = new Excel.RequestContext();
  context.setInvocation(invocation);

  try {
    // Use the context object to access the cell at the input address. 
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    range.load("values");
    await context.sync();

    range.values[0][0] = "Hello World";
    await context.sync();
  } catch (e) {
    // Handle the error if the address is invalid or the cell is not found.
    return `Error: ${e.message}`;
  }
  
  // Return the value of the cell at the input address.
  return range.values[0][0];
}



/**
 * @customfunction
 * @supportSync
 * @param {string} address The address of the cell from which to retrieve the value.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns The value of the cell at the input address.
 **/
export async function getRangeExcelContext(address, invocation) {
  const context = new Excel.RequestContext(undefined, invocation);
  
  const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
  range.load("values");
  await context.sync();
  
  return range.values[0][0];
}

/**
 * @customfunction
 * @supportSync
 * @param {string} address The address of the cell from which to retrieve the value.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns The value of the cell at the input address.
 **/
export async function getRangeExcelContextSet(address, invocation) {
  const context = new Excel.RequestContext();
  context.setInvocation(invocation);
  
  const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
  range.load("values");
  await context.sync();
  
  return range.values[0][0];
}
