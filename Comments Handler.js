/* 
* This function handles comments submitted through
* forms and alerts boss.
* @Trigger: Form-submit.
* Author: Abraham Vargas
*/
function commentsHandler() {
  var ss = SpreadsheetApp.openById("1xoED6K6FSYiII-Mg1qiyzt-e-5vpto6ZxT-k5pyCNKc");
  var sheet = ss.getSheets()[0]; // ”0″ is the first sheet
  var lastRow = sheet.getLastRow(); // Number of rows to process
  var comCol = sheet.getLastColumn()-1; // Comments column
  var dataRange = sheet.getRange(lastRow, comCol); // last row, Comments column
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
 
  // Write email if comment exists.
  if(String(data).length != 0){
    var emailAddress = "BOSSUSER@rowan.edu";
    var message = "A comment was submitted in the equipment checkout database."; // pulls the topic from the sixth column for the email message Abe edit string
    var subject = "Comment Submitted"; // makes the subject what is in quotes
    
    MailApp.sendEmail(emailAddress, subject, message); // send email to Boss.
  }
  descendOrder(); // Sorts sheet on form submit
}
