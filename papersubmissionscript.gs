// add menu to Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
      .addItem('Extract Emails', 'extractEmails')
      .addItem('Generate Tickets', 'GenerateTickets')
      .addItem('Generate Folders', 'GenerateFolders')
      .addItem('Send Emails', 'SendEmails')
      .addToUi();
}

// GLOBALS
//Array of file extension which you would like to extract to Drive
var fileTypesToExtract = ['pdf'];
//Name of the folder in google drive in which files will be put
var folderName = 'Paper Submissions';
//Name of the label which will be applied after processing the mail message
var labelName = 'Extracted';

// Paper submissions will be extracted from Gmail. The name and email of the sender will be pasted on the Google Sheet
function extractEmails() {

  try { 

    //Gmail search 
    var query = 'subject:(Paper-Submissions) has:attachment -label:Extracted in:anywhere';
    var threads = GmailApp.search(query);
    var messages = GmailApp.getMessagesForThreads (threads);
    var label = getGmailLabel_(labelName);
    var parentFolder;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    var emailArray = [];
    /*var titleArray = [];

    for (var i=0; i<threads.length; i++) {
      var title = messages[i].getAttachments().getName();
      titleArray.push(title.getName())
    }*/
  
    // get array of email addresses
    messages.forEach(function(message) {
      message.forEach(function(d) {
        emailArray.push(d.getFrom());
      });
    });

    // de-duplicate the array
    /*var uniqueEmailArray = emailArray.filter(function(item, pos) {
      return emailArray.indexOf(item) == pos;
    });*/
  
    var cleanedEmailArray = /*uniqueEmailArray*/emailArray.map(function(el) {
    var name = "";
    var email = "";
  
    var matches = el.match(/\s*"?([^"]*)"?\s+<(.+)>/);
  
    if (matches) {
      name = matches[1]; 
      email = matches[2];
    }
    else {
      name = "N/k";
      email = el;
    }

    return [name,email];
    
  })

    var r = 2
    var count = 0

    var rangen
    var rangee
  
    while (count == 0) {

      rangen = sheet.getRange(r,2)
      rangee = sheet.getRange(r,3)
      ranget = sheet.getRange(r,4)

      if(rangen.isBlank() && rangee.isBlank() && ranget.isBlank()) {
        sheet.getRange(r,2,cleanedEmailArray.length,2).setValues(cleanedEmailArray);
        //sheet.getRange(r,4,titleArray.length,1).setValues(titleArray);
        count = 1
      }

      r = r+1 

    }

    if(threads.length > 0){
      parentFolder = getFolder_(folderName);
    }

    var root = DriveApp.getRootFolder();
    var r2 = 2

    for(var i in threads){
      var mesgs = threads[i].getMessages();
	    for(var j in mesgs){
        //get attachments
        //var attachments = null;
        var attachments = mesgs[j].getAttachments();

        var l = attachments.length;
        var inc = "INCORRECT FILE TYPE"
        /*var attachment = attachments[0]; //get first attachment
        var isDefinedType = checkIfDefinedType_(attachment);
        var fileid = "";
        if(isDefinedType) {
          var attachmentBlob = attachment.copyBlob();
          var file = DriveApp.createFile(attachmentBlob);
          fileid = file.getId(); 
          //parentFolder = getFolder_(file.getName())
          parentFolder.addFile(file);
          root.removeFile(file);
        }*/
        var atts = []
        for(k in attachments) {
          var attachment = attachments[k];
          var isDefinedType = checkIfDefinedType_(attachment);
          var fileid = "";
          if(isDefinedType) {
            var attachmentBlob = attachment.copyBlob();
            var file = DriveApp.createFile(attachmentBlob);
            fileid = file.getId(); 
            //parentFolder = getFolder_(file.getName())
            parentFolder.addFile(file);
            root.removeFile(file);
            atts.push(fileid);
            //Logger.log("1:"+atts)
          }
        }

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getActiveSheet()

        var rangen2 = sheet.getRange(r2,2)
        var rangee2 = sheet.getRange(r2,3)
        var ranget2 = sheet.getRange(r2,4)
        var count = 0

        while((rangen2.isBlank() == false || rangee2.isBlank() == false) && count == 0) {

          rangen2 = sheet.getRange(r2,2)
          rangee2 = sheet.getRange(r2,3)
          ranget2 = sheet.getRange(r2,4)

          if (ranget2.isBlank()) {
            if (isDefinedType == false) {
              ranget2.setValue(inc)     
              count = 1
            }
            else {
              if (l > 1) 
              {
                var str = "";
                for(k in atts) {
                  str += atts[k]+","
                }
                ranget2.setBackground("yellow");
                //ranget2.setValue("WARNING: Multiple attachments found: " + l);
                //Logger.log(l);
                //Logger.log("2:"+atts)
                ranget2.setValue(str); 
              }
              else
              {
                ranget2.setValue(fileid);                
              }
              count = 1
            }
          }
          r2 = r2+1 
        }
	    } 

	    threads[i].addLabel(label);

    }

    } catch (e) {
      var ui = SpreadsheetApp.getUi();
      ui.alert('No new email submissions.');
    }

}

// This function will get the parent folder in Google drive
function getFolder_(folderName){
  var folder;
  var fi = DriveApp.getFoldersByName(folderName);
  if(fi.hasNext()){
    folder = fi.next();
  }
  else{
    folder = DriveApp.createFolder(folderName);
  }
  return folder;
}

// getDate n days back
// n must be integer
function getDateNDaysBack_(n){
  n = parseInt(n);
  var date = new Date();
  date.setDate(date.getDate() - n);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function getGmailLabel_(name){
  var label = GmailApp.getUserLabelByName(name);
  if(!label){
	label = GmailApp.createLabel(name);
  }
  return label;
}

// this function will check for filextension type.
// and return boolean
function checkIfDefinedType_(attachment){
  var fileName = attachment.getName();
  var temp = fileName.split('.');
  var fileExtension = temp[temp.length-1].toLowerCase();
  if(fileTypesToExtract.indexOf(fileExtension) !== -1) return true;
  else return false;
}

// Generates tickets for each email address extracted
function GenerateTickets() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()

  var r = 2

  var rangen = sheet.getRange(r,2)
  var rangee = sheet.getRange(r,3)
  
  while (rangen.isBlank() == false || rangee.isBlank() == false) {

    rangen = sheet.getRange(r,2)
    rangee = sheet.getRange(r,3)
    var ranget = sheet.getRange(r-1,1)

    if(ranget.isBlank()) {
      ranget.setValue(r-2)
    }

    r = r+1

  }

}

// Generates a unique folder in Google Drive for each ticket and pastes the link in Google Sheets
function GenerateFolders() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()

  var r = 2

  var ranget = sheet.getRange(r,1)
  var rangen = sheet.getRange(r,2)
  var rangel = sheet.getRange(r,5)

  while (rangen.isBlank() == false) {

    ranget = sheet.getRange(r,1)
    rangen = sheet.getRange(r,2)
    rangel = sheet.getRange(r,5)
    ranget2 = sheet.getRange(r,4)

    var temp1 = ranget.getValue()
    var temp2 = rangen.getValue()

    var inc = "INCORRECT FILE TYPE"
  
    if (rangel.isBlank() && rangen.isBlank() == false && ranget2.getValue() != "INCORRECT FILE TYPE") {

      var folder = DriveApp.createFolder(temp1+" - "+temp2);
      folder.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

      var url = folder.getUrl()
      var id = folder.getId()
      var getid = DriveApp.getFolderById(id)

      rangel.setValue(url)

      //Paper Submissions folder
      //var inputFolder = DriveApp.getFolderById('1kSUU2Mrq2E9y5d2G3piAA5DjxmpCmbff')
      //var files = inputFolder.getFiles();
      var id = ranget2.getValue()

      if (ranget2.isBlank() == false) {

        if (id.includes(",")) {

          var array = [];
          array = id.split(",");
          var i = 0

          while (i < (array.length-1)) { 

            var inputFolder = DriveApp.getFolderById('1kSUU2Mrq2E9y5d2G3piAA5DjxmpCmbff')
            var files = inputFolder.getFiles();

            var id2 = array[i];
            //Logger.log(id2)

            while (files.hasNext()) {
              var file = files.next()
              var fileid = file.getId()

              if (fileid == id2) {
                file.moveTo(getid);
              }

              //Logger.log("array[i]"+array[i])
              //Logger.log("id2: "+id2)

            }

            i = i+1
            //Logger.log("i: "+i)
          }

        }

        else {

          var inputFolder = DriveApp.getFolderById('1kSUU2Mrq2E9y5d2G3piAA5DjxmpCmbff')
          var files = inputFolder.getFiles();

          while (files.hasNext()) {
            var file = files.next()
            var fileid = file.getId()

            if (fileid == id) {
              file.moveTo(getid);
            }
          }

        }
    
      }

    }

    r = r+1

  }

}

// Send a confirmation email to the sender with their unique Google Drive link 
function SendEmails() {

  var sheet = SpreadsheetApp.getActiveSheet();

  var r = 2

  var emailRange = sheet.getRange(r, 3)
  var linkRange = sheet.getRange(r, 5)
  var ranges = sheet.getRange(r,6)
  var status = 'PROCESSED'
  var ranget = sheet.getRange(r,1)
  var rangen = sheet.getRange(r,2)

  while (emailRange.isBlank() == false || linkRange.isBlank() == false) {

    emailRange = sheet.getRange(r, 3)
    linkRange = sheet.getRange(r, 5)
    ranges = sheet.getRange(r,6)
    ranget = sheet.getRange(r,1)
    rangen = sheet.getRange(r,2)

    if(ranges.getValue() != status && (emailRange.isBlank() == false || linkRange.isBlank() == false)) {

      var emailAddress = emailRange.getValue()
      var link = linkRange.getValue()
      var author = rangen.getValue()
      var tn = ranget.getValue()
      var message = 'Text\nLink: '+link;
      var message = 'Hello '+author+'!\n\nWe have received your paper for peer review. Your ticket number is: '+tn+'. Your paper and the peer review feedback will be viewable on the link below:\n'+link+'\n\nFor inquiries, please send your e-mail to: innovatus@uap.asia.\n\nBest Regards,\nInnovatus Team'
      var subject = 'Submission Confirmation';
      MailApp.sendEmail(emailAddress, subject, message, {cc: 'innovatus@uap.asia'});

      ranges.setValue("PROCESSED")

    }

    r = r+1

  }

}

