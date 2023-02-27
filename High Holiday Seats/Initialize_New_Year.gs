/*
 * This script prepares the Yom Kippur signup to be used for a new year.
 * 
 * It does the following tasks:
 * 1) Save registration from previous year to our records spreadsheet.
 * 2) Clear all registration information from last year.
 * 3) Update the name and title of the Google Form to match the current year.
 * 4) Mark all seats as empty in the Google Properties Database. This is used when someone types in a 
 *      seat number on the Google Form to check whether that seat is available.
 * Note that you have to manually change the year on the Google Site.
 * 
 * 
 * If you have any questions about this system, contact Ahron Schwartz-Messing at aron.messin@gmail.com.
 */

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const PROPERTIES = SCRIPT_PROPERTIES.getProperties();


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Shenk")
    .addItem("New Year", "main")
    .addToUi();
}

function main() {
  const ui = SpreadsheetApp.getUi();
  const YEAR = ui.prompt("Year: ").getResponseText();

  const FORM = FormApp.openById(PROPERTIES.FORM_ID);
  const OLD_FORM_TITLE = FORM.getTitle();
  const NEW_FORM_TITLE = `Shenk Yom Kippur Seating ${YEAR} - Women`;


  const RESPONSE_SHEET = SpreadsheetApp.getActiveSheet()
  
  
  refreshResponses(RESPONSE_SHEET, OLD_FORM_TITLE);
  updateTitles(FORM, PROPERTIES.FORM_ID, NEW_FORM_TITLE);
  refreshSeatList(FORM);
}

/*******************************************************************************************************************
 * **************************************** FUNCTION DEFINITIONS ***************************************************
 * *****************************************************************************************************************
 */

/**
 * Save last year's signups as an old record;
 * Clear the signup sheet so that it's empty for the new year.
 */
function refreshResponses(responseSheet, oldTitle) {
  const firstResponseRow = 2;
  const numResponseRows = responseSheet.getLastRow() - 1;
  if(numResponseRows > 0) {
    // save old responses
    const OLD_SIGNUPS_SPREADSHEET = SpreadsheetApp.openById(PROPERTIES.OLD_SIGNUPS_ID);
    const backup_sheet = responseSheet.copyTo(OLD_SIGNUPS_SPREADSHEET);

    // in case someone runs the script twice with the same name, this will prevent a breaking exception
    try{
      backup_sheet.setName(oldTitle);
    } catch(e) { 
      console.log(e);
    }

    //clear response sheet entries
    responseSheet.deleteRows(firstResponseRow, numResponseRows);
  }
}

/**
 * Change the title of the Google Form to the new year.
 */
function updateTitles(form, formId, formTitle) {
  form.setTitle(formTitle);
  DriveApp.getFileById(formId).setName(formTitle);
}

/**
 * Update the database to reflect that all seats are now available.
 */
function refreshSeatList(form) {
const NUM_SEATS = 173
const ALL_SEATS = [...Array(NUM_SEATS + 1).keys()]
ALL_SEATS.shift();
const SEATS_REGEX_STRING = `\\b(${ALL_SEATS.join("|")})\\b`;

  PropertiesService.getScriptProperties().setProperty("SEATS", JSON.stringify(ALL_SEATS));
  form.getItems()[3].asTextItem()
                        .setValidation(FormApp.createTextValidation()
                                .requireTextMatchesPattern(SEATS_REGEX_STRING)
                                .setHelpText("That seat is not available. Make sure to only type the seat number.")
                                .build());
}
