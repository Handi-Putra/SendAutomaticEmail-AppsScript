function sendAutomatedEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Testing Send Automatic Email');

  var recipientRange = sheet.getRange("G2:G4"); // email to
  var ccRecipientRange = sheet.getRange("H2:H4"); // email cc
  var dateRange = sheet.getRange("B2:B4"); // Date
  var subjectRange = sheet.getRange("J2:J4"); // Subject
  var bodyRange = sheet.getRange("I2:I4"); // body

  // Get the values from the specified ranges
  var recipientValues = recipientRange.getValues(); // Value email To
  var ccRecipientValues = ccRecipientRange.getValues(); // Value email CC
  var dateValues = dateRange.getValues(); // Value Send Date email
  var subjectValues = subjectRange.getValues(); // Value Subject email
  var bodyValues = bodyRange.getValues(); // Value Body email

  for (var i = 0; i < dateValues.length; i++) {
    var date = dateValues[i][0];
    var recipient = recipientValues[i][0];
    var ccRecipient = ccRecipientValues[i][0];
    var subject = subjectValues[i][0];
    var bodyEmail = bodyValues[i][0];
    
    if (date && subject && recipient) {
      var formattedDate = formatDate(date);
      var subjectEmail = subject;
      var body = "Dengan Hormat Pak " + recipient + ",\n\n" + bodyEmail +
                 "\n";

      // Schedule the email to be sent
      var today = new Date();
      if (isSameDate(today, new Date(date))) {
        GmailApp.sendEmail(recipient, subjectEmail, body, {cc: ccRecipient});
      }
    }
  }
}

function formatDate(date) {
  var options = { year: 'numeric', month: 'long', day: 'numeric' };
  return new Date(date).toLocaleDateString('id-ID', options);
}

function isSameDate(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

function createDailyTrigger() {
  // Create a daily trigger to run the sendAutomatedEmail function
  ScriptApp.newTrigger('sendAutomatedEmail')
    .timeBased()
    .everyDays(1)
    .atHour(9) // Adjust the hour as needed
    .create();
}
