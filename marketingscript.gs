// add menu to Sheet
// run this first before anything else.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
      .addItem('Get Quota','getQuota')
      .addItem('Send Emails', 'SendEmails')
      .addToUi();
}

function getQuota()
{
  var emailQuotaRemaining = Math.floor(MailApp.getRemainingDailyQuota() * 0.2);
  console.log(emailQuotaRemaining);
}

// Send a confirmation email to the sender with their unique Google Drive link 
function SendEmails() {
  //Need to find out how much we can send.
  //This is 150 emails at 10% (UAP)
  //This is 10 emails for Gmail
  var emailQuotaRemaining = Math.floor(MailApp.getRemainingDailyQuota() * 0.2);
  //Call for Paper drive
  var cfp_file = DriveApp.getFileById('1uFCVFHH_nNi8o4Qy5IAeQkDRwUWanL5YIBSj_Z-le-0');

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

  var sent = 0
  //Put the email quota cap.
  //emailQuotaRemaining = 2; //debug
  while (emailRange.isBlank() == false && sent < emailQuotaRemaining) {
    rangen = sheet.getRange(r,nameIdx)
    emailRange = sheet.getRange(r, emailIdx)
    ranges = sheet.getRange(r,statusIdx)

    if(ranges.getValue() != status && emailRange.isBlank() == false) {

      var emailAddress = emailRange.getValue()
      var author = rangen.getValue()
      var message = 'Greetings sir/ma\'am:\n'+
      '\n'+
      'Warm greetings!\n'+
      '\n'+
      'The University of Asia and the Pacific (UA&P) – Information Science and Technology Department continues to undertake high level, interdisciplinary research in the area of Information Technology for the common good of society and to communicate the results of such research through various media and to varied audiences. As such we are pleased to announce the Special Issue of Innovatus entitled: \'Special Issue on Digital Transformation in Business Information Systems\'\n'+
      'In this special issue, we would like to focus on technology and methods used to transform businesses, particularly in these areas, during the COVID-19 pandemic and how these will continue to apply moving forward. Please refer to the attached Call for Paper PDF file for additional information.\n'+
      '\n' +
      'Important Dates\n'+
      '\n' + 
      'Full Paper Submission Deadline: December 1, 2021\n' +
      'Notification of Acceptance: Approximately 2 weeks after submission confirmation \n' +
      '\n' +
      'Suitable research/capstone topics include but are not limited to: \n'+
      '● Custom Systems Implementation supporting Remote Business Operations\n' +
      '● Information Technology Entrepreneurship\n' +
      '● Business Analytics and Data Science and its Applications to Improve Businesses\n' +
      '● Business Information Systems\n' +
      '● Mobile Applications and Internet of Things (IoT)\n' +
      '● Educational Technology and Technology Integrations\n' +
      '● Information Systems in Education, Healthcare, Manufacturing, and Transportation\n' + 
      '● E-commerce\n'+
      '\n'+
      'Keywords: Business Information Systems, Digital Transformation, Custom Systems, Remote Work, Information Technology Entrepreneurship, Business Analytics, Data Science, Internet of Things, E-commerce, Online Learning\n' +
      '\n' +
      'Papers submitted under this category will require a peer review phase with evaluators consulting the editor-in-chief regarding revisions and its acceptance for publication. The submission guidelines and the submission portal can be found here. Article Processing Charge (APC) will be waived for now. For inquiries, please email Mr. Giuseppe Ng at innovatus@uap.asia.\n'+
      '\n' +
      'If you do not wish to receive these notifications, please let us know at innovatus@uap.asia' +
      '\n' +
      'Best Regards,\n'+
      'Innovatus Team'
      //https://developers.google.com/apps-script/reference/mail/mail-app#sendEmail%28String,String,String,Object%29
      //TODO: The HTML version should likely be load an html file from Google Drive instead.
      var subject = 'Innovatus Information Technology Journal: Call for Papers for Special Issue 2021';
      MailApp.sendEmail(emailAddress, subject, message,{
        htmlBody: '<style>body {font-family: Garamond;} a{font-family: \'Courier New\';}</style>Greetings sir/ma\'am:<br />'+
      '<br />'+
      'Warm greetings!<br />'+
      '<br />'+
      '<p>The <b>University of Asia and the Pacific (UA&P) – Information Science and Technology Department</b> continues to undertake high level, interdisciplinary research in the area of Information Technology for the common good of society and to communicate the results of such research through various media and to varied audiences. As such we are pleased to announce the Special Issue of Innovatus entitled: <b><i>\'Special Issue on Digital Transformation in Business Information Systems\'</i></b></p>'+
      '<p>In this special issue, we would like to focus on technology and methods used to transform businesses, particularly in these areas, during the COVID-19 pandemic and how these will continue to apply moving forward. Please refer to the attached Call for Paper PDF file for additional information.</p>' +
      '<b>IMPORTANT DATES</b><br />'+
      '<br />' + 
      '<b>Full Paper Submission Deadline:</b> December 1, 2021<br />' +
      '<b>Notification of Acceptance:</b> Approximately 2 weeks after submission confirmation <br />' +
      '<b>Official Website: </b><a href=\'https://innovatus.uap.asia\'>https://innovatus.uap.asia</a><br />'+
      '<i>Indexed at Google Scholar</i>'+
      '<br />'+
      '<p>Suitable research/capstone topics include but are not limited to: <br />'+
      '<ul> <li>Custom Systems Implementation supporting Remote Business Operations</li>' +
      '<li>Information Technology Entrepreneurship</li>' +
      '<li>Business Analytics and Data Science and its Applications to Improve Businesses</li>' +
      '<li>Business Information Systems</li>' +
      '<li>Mobile Applications and Internet of Things (IoT)</li>' +
      '<li>Educational Technology and Technology Integrations</li>' +
      '<li>Information Systems in Education, Healthcare, Manufacturing, and Transportation</li>' + 
      '<li>E-commerce</li></ul>'+
      '<br />'+
      'Keywords: Business Information Systems, Digital Transformation, Custom Systems, Remote Work, Information Technology Entrepreneurship, Business Analytics, Data Science, Internet of Things, E-commerce, Online Learning</p>' +
      '<p>Papers submitted under this category will require a peer review phase with evaluators consulting the editor-in-chief regarding revisions and its acceptance for publication. The submission guidelines and the submission portal can be found here. Article Processing Charge (APC) will be waived for now. For inquiries, please email Mr. Giuseppe Ng at <a href=\'mailto:innovatus@uap.asia\'>innovatus@uap.asia</a>.</p>'+
      '<br />' +
      'If you do not wish to receive these notifications, please let us know at <a href=\'mailto:innovatus@uap.asia\'>innovatus@uap.asia</a>' +
      '<br />' +
      'Best Regards,<br />'+
      '<b>Giuseppe Ng</b><br />' +
      '<i>Editor-in-chief</i><br />' +
      '<i>Innovatus</i>',
        attachments: [cfp_file.getAs(MimeType.PDF)]
      });
      ranges.setValue("PROCESSED")
      sent = sent + 1
    }

    //move row to the next one.
    r = r+1
  }
}

