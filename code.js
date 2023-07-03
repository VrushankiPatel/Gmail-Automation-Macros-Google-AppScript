function scheduleEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Emails');
  var sheet2 = ss.getSheetByName('7_Days_Prior_Template');
  var sheet3 = ss.getSheetByName('2_Days_After_Template');
  var sheet4 = ss.getSheetByName('7_Days_After_Template');
  
  var dataRange = sheet1.getDataRange();
  var data = dataRange.getValues();
  
  // Delete all existing triggers 
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendMails') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Loop through all rows of data in Sheet1
  for (var i = 1; i < data.length; i++) {
    var eventName = data[i][0];
    var contactName = data[i][1];
    var recipient = data[i][2];
    var subject = data[i][11 + i]
    var eventDate = new Date(data[i][3]);
    var date7DaysBefore = new Date(eventDate.getTime() - 7 * 24 * 60 * 60 * 1000);   // Schedule date for 7 days prior mail
    var date2DaysAfter = new Date(eventDate.getTime() + 2 * 24 * 60 * 60 * 1000);    // Schedule date for 2 days after mail
    var date7DaysAfter = new Date(eventDate.getTime() + 7 * 24 * 60 * 60 * 1000);    // Schedule date for 2 days after mail
    
    // Create the email drafts using the templates in Sheet2, Sheet3, and Sheet4
    var drafts = [{template: sheet2.getRange('A1').getValue(), date: date7DaysBefore}, {template: sheet3.getRange('A1').getValue(), date: date2DaysAfter},{template: sheet4.getRange('A1').getValue(), date: date7DaysAfter}];
    
    var draftIds = [];
    
    // Loop through the email drafts and create a new draft message for each
    for (var j = 0; j < drafts.length; j++) {
      var draft = drafts[j];
      var emailBody = draft.template
        .replace('[EVENT_NAME]', eventName)
        .replace('[CONTACT_NAME]', contactName)
        .replace('[EVENT_DATE]', eventDate.toDateString());
      
      var newDraft = GmailApp.createDraft(recipient, subject, emailBody);
      draftIds.push(newDraft.getId());
      
      // Schedule the email draft to be sent at the specified date and time
      ScriptApp.newTrigger('sendEmail')
        .timeBased()
        .at(draft.date)
        .create();
    }
    
    // Update Sheet1 with the draft IDs and scheduled send dates
    sheet1.getRange(i+1, 5).setValue(draftIds[0]);
    sheet1.getRange(i+1, 6).setValue(date7DaysBefore);
    sheet1.getRange(i+1, 7).setValue(draftIds[1]);
    sheet1.getRange(i+1, 8).setValue(date2DaysAfter);
    sheet1.getRange(i+1, 9).setValue(draftIds[2]);
    sheet1.getRange(i+1, 10).setValue(date7DaysAfter);
  }
}

// Function to send the scheduled email drafts
function sendEmail() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var draftId = triggers[i].getUniqueId();
    var draft = GmailApp.getDraft(draftId);
    if (draft) {
      draft.send();
      draft.getMessage().moveToTrash();
    }
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// 2023-02-23
// 2023-02-25
// 2023-02-28
// 2023-03-03


// function sendEventEmails() {
//   var sheet = SpreadsheetApp.getActiveSheet();
//   var data = sheet.getDataRange().getValues();
//   var today = new Date();
  
//   // Date format: "YYYY-MM-DD"
//   // Loop through each row in the sheet
//   for (var i = 1; i < data.length; i++) {
//     var eventName = data[i][0];
//     var contactName = data[i][1];
//     var email = data[i][2];
//     var eventDate = new Date(data[i][3]);
//     var processed = data[i][4];

//     // Check if the email has already been processed
//     if (processed == "Y") {
//       continue;
//     }
    
//     // Check if the event is 7 days prior
//     if (Math.ceil((eventDate - today) / (1000 * 60 * 60 * 24)) == 7) {
//       var templateSheet = SpreadsheetApp.getActive().getSheetByName("7 Days Prior Template");
//       var template = templateSheet.getRange("A1").getValue();
//       // SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template1").getRange(1, 1).getValue();
      
//       template = template.replace("{eventName}", eventName);
//       template = template.replace("{contactName}", contactName);
//       template = template.replace("{eventDate}", eventDate.toLocaleDateString());
      
//       GmailApp.createDraft(email, "Event Reminder: " + eventName, template);

//       sheet.getRange(i + 1, 6).setValue(eventDate - today);
//       sheet.getRange(i + 1, 5).setValue("Y");
//     }
//     // sheet.getRange(i + 1, 5).setValue("N");
//     // Check if the event is 2 days after
//     if (Math.ceil((today - eventDate) / (1000 * 60 * 60 * 24)) == 2) {
//       var templateSheet = SpreadsheetApp.getActive().getSheetByName("2 Days After Template");
//       var template = templateSheet.getRange("A1").getValue();
      
//       template = template.replace("{eventName}", eventName);
//       template = template.replace("{contactName}", contactName);
//       template = template.replace("{eventDate}", eventDate.toLocaleDateString());
      
//       GmailApp.createDraft(email, "Event Follow-up: " + eventName, template);

//       sheet.getRange(i + 1, 7).setValue(today - eventDate);
//       sheet.getRange(i + 1, 5).setValue(+1);
//     }
    
//     sheet.getRange(i + 1, 5).setValue(+1);
//     // Check if the event is 7 days after
//     if (Math.ceil((today - eventDate) / (1000 * 60 * 60 * 24)) == 7) {
//       var templateSheet = SpreadsheetApp.getActive().getSheetByName("7 Days After Template");
//       var template = templateSheet.getRange("A1").getValue();
      
//       template = template.replace("{eventName}", eventName);
//       template = template.replace("{contactName}", contactName);
//       template = template.replace("{eventDate}", eventDate.toLocaleDateString());
      
//       GmailApp.createDraft(email, "Event Recap: " + eventName, template);

//       sheet.getRange(i + 1, 8).setValue(today - eventDate);
    
//   }
// }

// function createTrigger() {
//   var trigger = ScriptApp.newTrigger("sendEventEmails")
//   .timeBased()
//   .onWeekDay(ScriptApp.WeekDay.FRIDAY)
//   .atHour(9)
//   .create();
// }
// }