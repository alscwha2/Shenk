/**
 * @author: Aaron Schwartz-Messing
 * @email: aron.messin@gmail.com
 * 
 * Please feel free to reach out if you have any questions about this script.
 * 
 * 
 * This script processes the form submission for the YK signup. It does the following tasks:
 * 1) Checks that seat is available
 * 2) If seat is available:
 *      1) Mark seat as unavailable
 *      2) Update the data validation on the Google Form so that no one else can sign up for that seat
 *      3) Send confirmation email
 * 3) If seat is unavailable:
 *      1) Remove the form submission from the signup sheet
 *      2) Send rejection email
 * 4) If an error occurs, send an error email.
 */


function processSubmission(e) {
  const ONE_MINUTE = 60 * 1000;
  const lock = LockService.getDocumentLock();
  if(lock.tryLock(ONE_MINUTE)) {
    try{
      const seatNumber = parseInt(e.namedValues.Seat[0]);
      const emptySeats = JSON.parse(PROPERTIES.SEATS);

      if(emptySeats.includes(seatNumber)) {
        emptySeats.splice(emptySeats.indexOf(seatNumber), 1);
        SCRIPT_PROPERTIES.setProperty("SEATS", JSON.stringify(emptySeats));
        updateEmptySeatDataValidation(emptySeats);
        sendConfimationEmail(e);
      } else {
        removeSubmissionFromSheet(e);
        sendRejectionEmail(e);
      }
    } catch(error) {
      removeSubmissionFromSheet(e);
      sendErrorEmail(error);
      throw error
    } finally{
      lock.releaseLock();
    }
  } else {
      removeSubmissionFromSheet(e);
      Logger.log("DOCUMENT LOCK TIMEOUT");
      sendErrorEmail(e);
  }
}

function removeSubmissionFromSheet(e) {
  const range = e.range;
  const RESPONSE_SHEET = SpreadsheetApp.getActiveSheet();
  RESPONSE_SHEET.deleteRow(range.getRowIndex())
}

function updateEmptySeatDataValidation(emptySeats) {
  const seatArrayToRegex = (array) => `\\b(${array.join("|")})\\b`;

  const FORM = FormApp.openById(PROPERTIES.FORM_ID);
  FORM.getItems()[3].asTextItem()
      .setValidation(FormApp.createTextValidation()
                                .requireTextMatchesPattern(seatArrayToRegex(emptySeats))
                                .setHelpText("Seat unavailable. Be sure to only type the number of an empty seat with no other characters.")
                                .build());

}

function sendConfimationEmail(e) {
  const emailSubject = "Reservation for Yom Kippur Seating - Women's Seat: " + e.namedValues.Seat + " Confirmed";

  const emailBody = `Hello ${e.namedValues["First Name"]},

  Thank you for reserving women's seat ${e.namedValues.Seat}. Your reservation has been confirmed.


  If you no longer wish to keep this seat selection, please inform us as soon as possible, so that someone else can reserve your seat.

  If you are not a member, please be sure to pay for your Yom Kippur seats before coming to shul. You can submit your payment here: https://form.jotform.com/222583799450063. (student rate: $75, professional rate: $100)

  We look forward to greeting you on Yom Kippur,
  The Shenk Seating Team`;
  
  GmailApp.sendEmail(e.namedValues['Email Address'], emailSubject, emailBody);
}

function sendRejectionEmail(e) {
  const emailSubject = "Reservation for Yom Kippur Seating - Women's Seat: " + e.namedValues.Seat + " NOT Confirmed";

  const emailBody = `Hello ${e.namedValues["First Name"]},

  Unfortunately the seat you selected has already been taken. Please select another seat from https://sites.google.com/site/shenkshulseating

  Thanks,
  The Shenk Seating Team`;

  GmailApp.sendEmail(e.namedValues['Email Address'], emailSubject, emailBody);
}

function sendErrorEmail(e) {
  const emailSubject = "Error occurred during seat reservation";

  const emailBody =  `If you have received this email something went wrong. Please let us know at shenkshulseating@gmail.com

  Thanks,
  The Shenk Seating Team`;

  GmailApp.sendEmail(e.namedValues['Email Address'], emailSubject, emailBody);
  GmailApp.sendEmail("shenkshulseating@gmail.com", emailSubject, emailBody);
}
