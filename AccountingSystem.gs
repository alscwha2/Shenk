/**
 * @author: Ahron Schwartz-Messing
 * @email: aron.messin@gmail.com
 * 
 * If you have any questions, feel free to reach out.
 * 
 * 
 * Aggregate all financial form submissions into one master sheet.
 * Reformat each row so that it fits the following template:
 *  ["TimeStamp", "Amount", "From Sheet", "First Name", "Last Name", ... all other content]
 * Sort the aggregation sheet by TimeStamp in ascending order.
 */

/*
 * The ID of a folder in Google Drive can be found at the end of the URL
 * Hiding these URLs in Github for security reasons.
 */
const financialFormFolderID = "";
const financialAggregationSheetURL = "";
const expenseReimbursementSheetURL = "";

const financialFormFolder = DriveApp.getFolderById(financialFormFolderID);
const financialAggregationSheet = SpreadsheetApp.openByUrl(financialAggregationSheetURL).getSheetByName("Aggregation");
const expenseReimbursementSheet = SpreadsheetApp.openByUrl(expenseReimbursementSheetURL).getSheetByName("Shenk Expense Reimbursement Form");


/*
 * Process all of the sheets containing incoming money
 * Process the sheet containing outgoing money
 * Sort the aggregation sheet by timestamp
 */
function reformatAndAggregate() {
  // ############################################ Function Definitions ##############################################################
  function processIncomeSheet(file) {
    // file info
    const fileName = file.getName();
    const sheet = SpreadsheetApp.openByUrl(file.getUrl()).getActiveSheet();
    const values = sheet.getDataRange().getDisplayValues();

    // index info
    var paymentIndex;
    var firstNameIndex;
    var lastNameIndex;
    
    const titleRow = values.shift();

    function findColumnIndices(titleRow) {
      for (var i = 0; i < titleRow.length; i++) {
        var heading = titleRow[i];
        if (heading.endsWith(": Products")) {
          paymentIndex = i;
        } else if (heading === "First Name") {
          firstNameIndex = i;
        } else if (heading === "Last Name") {
          lastNameIndex = i;
        }
      }
    };
    findColumnIndices(titleRow);

    // application logic
    function reformatRow(row) {
      function getPaymentAmount(cellContents) {
        var paymentAmount = cellContents.match(/Total: (\d+.\d{2}) USD/);
        return paymentAmount === null ? "0.00" : paymentAmount[1];
      }
      row.splice(1, 0, getPaymentAmount(row[paymentIndex]));
      row.splice(2, 0, fileName);
      row.splice(3, 0, row[firstNameIndex + 2]);
      row.splice(firstNameIndex + 3, 1);
      row.splice(4, 0, row[lastNameIndex + 2]);
      row.splice(lastNameIndex + 3, 1);
    };
    
    for (var row of values) {
      reformatRow(row);
      financialAggregationSheet.appendRow(row);
    }
  };

  function processExpenseReimbursements() {
    const values = expenseReimbursementSheet.getDataRange().getDisplayValues();
    values.shift();
    for (row of values) {
      row.splice(1, 0, "-" + row[5]);
      row.splice(6, 1);
      row.splice(2, 0, "Expense Reimbursements");
      financialAggregationSheet.appendRow(row);
    }
  };

  // ############################################ Application Logic ##############################################################

  // prep the aggregation sheet
  financialAggregationSheet.clearContents();
  financialAggregationSheet.appendRow(["TimeStamp", "Amount", "From Sheet", "First Name", "Last Name"]);

  // process all of the sheets containing incoming money
  const files = financialFormFolder.getFiles();
  while (files.hasNext()) {
    processIncomeSheet(files.next())
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
};
