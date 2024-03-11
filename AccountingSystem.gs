/**
 * @author: Ahron Schwartz-Messing
 * @email: 
 * 
 * If you have any questions, feel free to reach out.
 * 
 * 
 * Aggregate all financial form submissions into one master sheet.
 * Reformat each row so that it fits the following template:
 *  ["TimeStamp", "Amount", "From Sheet", "First Name", "Last Name", ... all other content]
 * Sort the aggregation sheet by TimeStamp in ascending order.
 * 
 * This script should be attached to the Financial Aggregation Spreadsheet. 
 * This has two ramifications:
 *  1 - This allows us to create a custom menu called "Scripts" in the Ui of the 
 *       Financial Aggregation Spreadsheet from which one can run this code on demand
 *  2 - This allows us to access the Financial Aggregation Spreadsheet from this code by calling
 *        SpreadsheetApp.getActiveSpreadsheet(). If this code is ever disconnected form the sheet 
 *        then this call will throw an error.
 * 
 * Ways that this code can go terribly wrong:
 * 1 - Don't mess with the title row! Make sure that it follows the correct format or else it will 
 *      not be read properly!
 * 
 * TODO:
 * [] - check if the columns exist 
 *      if they don't
 *        prompt the user to identify the correct column
 *        take the user identified column and ask them if they're sure
 *        change the header to match
 * [] - check if there is more than one column that matches
 *       if yes
 *         prompt the user to select which column to use
 *         rename the other one to something that will not conflict
 */

/*
 * The ID of a folder in Google Drive can be found at the end of the URL
 */
const financialFormFolderID = "";
const financialAggregationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const expenseReimbursementSheetURL = "";

const financialFormFolder = DriveApp.getFolderById(financialFormFolderID);
const financialAggregationSheet = financialAggregationSpreadsheet.getSheetByName("Aggregation");
const expenseReimbursementSheet = SpreadsheetApp.openByUrl(expenseReimbursementSheetURL).getSheetByName("Shenk Expense Reimbursement Form");


/*
 * Find the column numbers for the the Payment-amount, first-name, and last-name fields of a file by analyzing the title row names
 * 
 * Warning: if there are multiple rows with the same name, this code will take the last one
 **/
function findColumnIndices(titleRow) {
  for (var i = 0; i < titleRow.length; i++) {
    var heading = titleRow[i];
    if (heading.endsWith(": Products")) {
      var paymentIndex = i;
    } else if (heading === "First Name") {
      var firstNameIndex = i;
    } else if (heading === "Last Name") {
      var lastNameIndex = i;
    }
  }
  return [paymentIndex, firstNameIndex, lastNameIndex]
}

/*
 * Extract the payment amount from a cell containing the full payment information
 **/
function getPaymentAmount(cellContents) {
  const paymentAmount = cellContents.match(/Total: (\d+.\d{2}) USD/);
  return paymentAmount === null ? "0.00" : paymentAmount[1];
}

/*
 * Append data to aggregation sheet.
 * 
 * instead of calling appendRow on each row,calculate the range that will be added and write to it directly
 * this decreased runtime by 10x!
 **/
function appendDataToAggregationSheet(values) {
  if (values.length === 0) return;

  // construct the Range in the aggregate sheet to write the reformatted data to
  const firstRow = financialAggregationSheet.getLastRow() + 1;
  const firstColumn = 1; //note that column indexing starts at 1 not 0
  const numRows = values.length;
  const numColumns = values[0].length;

  const range = financialAggregationSheet.getRange(firstRow, firstColumn, numRows, numColumns);

  range.setValues(values);
}

/*
 * Reformat the info on a data sheet and write it to the aggreagation sheet.
 * 
 * 1 - remove title row
 * 2 - rearrange the columns to fit the format of the aggregation sheet
 * 3 - write the rearranged data to the sheet
 **/
function processIncomeSheet(file) {
  // file info
  const fileName = file.getName();
  const sheet = SpreadsheetApp.openByUrl(file.getUrl()).getActiveSheet();

  // extract values
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) return;

  // process the title row to get the indices of the important fields
  const titleRow = values.shift();
  const [paymentIndex, firstNameIndex, lastNameIndex] = findColumnIndices(titleRow);

  // move arounds the columns of each row to match the format of the aggregation sheet
  function reformatRow(row) {
    row.splice(1, 0, getPaymentAmount(row[paymentIndex]));
    row.splice(2, 0, fileName);
    row.splice(3, 0, row[firstNameIndex + 2]);
    row.splice(firstNameIndex + 3, 1);
    row.splice(4, 0, row[lastNameIndex + 2]);
    row.splice(lastNameIndex + 3, 1);
  };
  
  for (var row of values) {
    reformatRow(row);
  }

  // write the reformatted data to the aggregate sheet
  appendDataToAggregationSheet(values);
}

/*
 * The Expense Reimbursement form needs to be processed separately because it is formatted differently than the other sheets
 **/
function processExpenseReimbursements() {
  const values = expenseReimbursementSheet.getDataRange().getDisplayValues();

  // get rid of title row
  values.shift();

  // reformat the rows
  for (row of values) {
    row.splice(1, 0, "-" + row[5]);
    row.splice(6, 1);
    row.splice(2, 0, "Expense Reimbursements");
  }

  // write the reformatted data to the aggregate sheet
  appendDataToAggregationSheet(values);
}

/*
 * Process all of the sheets containing incoming money
 * Process the sheet containing outgoing money
 * Sort the aggregation sheet by timestamp
 */
function reformatAndAggregate() {
  // prep the aggregation sheet
  financialAggregationSheet.clearContents();
  financialAggregationSheet.appendRow(["TimeStamp", "Amount", "From Sheet", "First Name", "Last Name"]);

  // process all of the sheets containing incoming money
  const files = financialFormFolder.getFiles();
  while (files.hasNext()) {
    processIncomeSheet(files.next());
  }

  // process the sheet containing outgoing money
  processExpenseReimbursements();

  financialAggregationSheet.sort(1);
  SpreadsheetApp.flush();
}


/*
 * The purpose of this funciton is to create a menu called "Scripts" with an option called "Refresh" that would run this script.
 * This function is not called directly, rather it is called automatically when the document is opened.
 * Currently, this only works if you open the sheet and the only Google account that you are logged into is yeshivacommunityshul@gmail.com.
 * Alas.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts')
      .addItem('Refresh', 'reformatAndAggregate')
      .addToUi();
}
