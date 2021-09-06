// add menu to Sheet
// run this first before anything else.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
      .addItem('Send Emails', 'SendEmails')
      .addToUi();
}

// Send a confirmation email to the sender with their unique Google Drive link 
function SendEmails() {
  //Need to find out how much we can send.
  //Reduce by 100 just to give a buffer.
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota() - 100;
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
      '\n'+
      'The digital transformation in the business industry has been on the rise for many years, and has only been accelerated by the COVID-19 pandemic. Many industries have been forced to shift to a remote work setup to adapt to the “new normal”. While the digital transformations for some businesses may have come about due to the circumstances, the changes implemented may still be improved to remain relevant in the foreseeable future.\n'+
      '\n' +
      'Mobile and Internet of Things (IoT) applications have helped with many different types of business improvements. During the pandemic, they have helped increase the efficiency and effectivity of remote setups, particularly in the field of medicine and education. Some of the challenges of these applications are data accuracy and the integration of new technologies towards monitoring, real-time processing, and process self-optimizations. As the shift towards the new normal continues, innovations in the mobile space can help businesses cope and thrive with the challenges of the pandemic.\n'+
      '\n' +
      'The education industry has taken a huge shift due to the pandemic. Schools and universities have been forced to adapt to the online setup, while some institutionshave slowly been shifting to a flexible blended learning format. As we continue to adjust to the new normal, the education setup and methods of teaching and learning will undoubtedly continue to evolve.\n'+
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
      'In this special issue, we would like to focus on technology and methods used to transform businesses, particularly in these areas, during the COVID-19 pandemic and how these will continue to apply moving forward.\n'+
      '\n' +
      'Important Dates\n'+
      '\n' + 
      'Full Paper Submission Deadline: December 1, 2021\n' +
      'Notification of Acceptance: Approximately 2 weeks after submission confirmation \n' +
      '\n'+
      'Papers submitted under this category will require a peer review phase with evaluators consulting the editor-in-chief regarding revisions and its acceptance for publication. The submission guidelines and the submission portal can be found here. Article Processing Charge (APC) will be waived for now. For inquiries, please email Mr. Giuseppe Ng at innovatus@uap.asia.\n'+
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
      '<p>The <b>University of Asia and the Pacific (UA&P) – Information Science and Technology Department</b> continues to undertake high level, interdisciplinary research in the area of Information Technology for the common good of society and to communicate the results of such research through various media and to varied audiences. As such we are pleased to announce the Special Issue of Innovatus entitled: <b><i>\'Special Issue on Digital Transformation in Business Information Systems\'</i></b></p><br />'+
      '<br />'+
      '<center><b>CALL FOR PAPERS</b></center>' +
      '<br />'+
      '<p>The digital transformation in the business industry has been on the rise for many years, and has only been accelerated by the COVID-19 pandemic. Many industries have been forced to shift to a remote work setup to adapt to the “new normal”. While the digital transformations for some businesses may have come about due to the circumstances, the changes implemented may still be improved to remain relevant in the foreseeable future.</p>'+
      '<p>Mobile and Internet of Things (IoT) applications have helped with many different types of business improvements. During the pandemic, they have helped increase the efficiency and effectivity of remote setups, particularly in the field of medicine and education. Some of the challenges of these applications are data accuracy and the integration of new technologies towards monitoring, real-time processing, and process self-optimizations. As the shift towards the new normal continues, innovations in the mobile space can help businesses cope and thrive with the challenges of the pandemic.</p>'+
      '<p>The education industry has taken a huge shift due to the pandemic. Schools and universities have been forced to adapt to the online setup, while some institutions have slowly been shifting to a flexible blended learning format. As we continue to adjust to the new normal, the education setup and methods of teaching and learning will undoubtedly continue to evolve.</p>'+
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
      '<p>In this special issue, we would like to focus on technology and methods used to transform businesses, particularly in these areas, during the COVID-19 pandemic and how these will continue to apply moving forward.</p>'+
      '<br />' +
      '<b>IMPORTANT DATES</b><br />'+
      '<br />' + 
      '<b>Full Paper Submission Deadline:</b> December 1, 2021<br />' +
      '<b>Notification of Acceptance:</b> Approximately 2 weeks after submission confirmation <br />' +
      '<b>Official Website: </b><a href=\'https://innovatus.uap.asia\'>https://innovatus.uap.asia</a><br />'+
      '<br />'+
      '<p>Papers submitted under this category will require a peer review phase with evaluators consulting the editor-in-chief regarding revisions and its acceptance for publication. The submission guidelines and the submission portal can be found here. Article Processing Charge (APC) will be waived for now. For inquiries, please email Mr. Giuseppe Ng at <a href=\'mailto:innovatus@uap.asia\'>innovatus@uap.asia</a>.</p>'+
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

