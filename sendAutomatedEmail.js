function sendAutomatedEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Restart JBoss Pega Teller');

  var recipient = "INPUT TO MAIL HERE"; // Replace with the recipient's email address
  var ccRecipient = "INPUT CC MAIL HERE"; // Replace with the actual CC email address

  var dateRange = sheet.getRange("B39:B41"); // Date
  var subjectRange = sheet.getRange("J39:J41"); // Subject

  // Get the values from the specified ranges
  var dateValues = dateRange.getValues();
  var subjectValues = subjectRange.getValues();

  for (var i = 0; i < dateValues.length; i++) {
    var date = dateValues[i][0];
    var subject = subjectValues[i][0];
    
    if (date && subject) {
      var formattedDate = formatDate(date);
      var subjectEmail = subject;
      var body = "Dengan Hormat Pak " + recipient + ",\n\n" +
                 "hanya Sekedar testing saja\n";

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
