// add menu to Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
      .addItem('Send Emails', 'SendEmails')
      .addToUi();
}

// Send a confirmation email to the sender with their unique Google Drive link 
function SendEmails() {

  var sheet = SpreadsheetApp.getActiveSheet();

  //start at row r
  var r = 2

  //Column indices
  var nameIdx = 1
  var emailIdx = 2
  var statusIdx = 3

  var rangen = sheet.getRange(r,nameIdx)
  var emailRange = sheet.getRange(r, emailIdx)
  var ranges = sheet.getRange(r,statusIdx)
  var status = 'PROCESSED'

  while (emailRange.isBlank() == false) {
    rangen = sheet.getRange(r,nameIdx)
    emailRange = sheet.getRange(r, emailIdx)
    ranges = sheet.getRange(r,statusIdx)

    if(ranges.getValue() != status && emailRange.isBlank() == false) {

      var emailAddress = emailRange.getValue()
      var author = rangen.getValue()
      var message = 'Hello '+author+'!\n\nPlease update TODO. Please send your e-mail to: innovatus@uap.asia.\n\nBest Regards,\nInnovatus Team'
      var subject = 'Innovatus Information Technology Journal: Call for Papers';
      MailApp.sendEmail(emailAddress, subject, message);
      ranges.setValue("PROCESSED")
    }

    //move row to the next one.
    r = r+1
  }
}

