/*
 * @author: Aaron Schwartz-Messing
 * @email: aron.messin@gmail.com
 * 
 * If you have any questions about how this system works, feel free to reach out.
 */

/* 
 * Edit this to change the maximum number of days that a name stays on the list
 */
const maxDays = 30

/*
 * Edit this to change the text of the email
 */
function getEmailText(row) {
   return `Dear ${row[YOURNAME]},\n\nNames are automatically removed from the tehillim list after ${maxDays} days. If ${row[CHOLEH]} still needs a refuah, please fill out the form again. The form can be found at shenkshul.com/tehillim.`
}

/*
 * Edit this to change the subject of the email
 */
function getEmailSubject(row) {
  return `Does ${row[CHOLEH]} still need a refuah?`
}

/*
 * Edit this to change the URL of the tehillim form submissions
 * URL hidden on GitHub for secuiry purposes.
 */
const sheet = SpreadsheetApp.openByUrl("").getSheetByName("List")

const TIMESTAMP = 0
const YOURNAME = 1
const EMAIL = 2
const CHOLEH = 3
const GENDER = 4

/*
 * For every tehillim entry that was submitted longer than maxDays ago:
 *    1) send them an email
 *    2) delete the entry
 */
function cleanTehillimList() {
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4)
  const rows = range.getDisplayValues()

  // Iterate backwards so that you don't mess up the numbering when deleting
  for(var i = rows.length - 1; i >= 0; i--) {
    row = rows[i]

    // if it's been too long, send the email and delete the row
    if (enoughTimeElapsed(row[TIMESTAMP])) {
      //MailApp.sendEmail("aron.messin@gmail.com", getEmailSubject(row), getEmailText(row))
      // sheet.deleteRow(i + 2)
      SpreadsheetApp.flush()
    }
  }
}

/*
 * Test whether enough time has elapsed to safely remove name from list
 */
function enoughTimeElapsed(submissionDate) {
  submissionDate = new Date(reformatTimestamp(submissionDate))
  const maxDaysInMillis = maxDays * 24 * 60 * 60 * 1000
  return (Date.now() - submissionDate) > maxDaysInMillis
}

/*
 * Convert the google-forms timestamp to ISO format
 */
function reformatTimestamp(stamp) {
  function reformatDate(date) {
    date = date.split("/")

    // add leading zeroes
    if(date[0].length < 2) {
      date[0] = "0" + date[0]
    }
    if(date[1].length < 2) {
      date[1] = "0" + date[1]
    }

    //put pieces in the correct order and join
    date = [date[2], date[0], date[1]]
    date = date.join("-")
    return date
  }

  function reformatTime() {
    // add a leading zero if the hour was a single digit
    if(time.length == 7) {
      time = "0" + time
    }
    return time

  }
  stamp = stamp.split(" ")
  var date = stamp[0]
  var time = stamp[1]

  date = reformatDate(date)
  time = reformatTime(time)

  return date + "T" + time
}
