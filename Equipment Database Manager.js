/**
* Equipment Database Manager
* This script was created to manage the rental of equipment in the database. It will alert staff via email when an
* item has not been returned by the client before the due date.
* Script will also send an email to the clients to return their items immediately.
* @Trigger: Daily between 9am-12pm.
* Author: Abraham Vargas.
*/
 function exceededDeadlineEmail() {
   var thisDate = new Date();
   var greens = ['#d9ead3', '#00ff00', '#b6d7a8', '#93c47d', '#6aa84f', '#38761d', '#274e13'] // every shade of green
   
   var ss = SpreadsheetApp.openById("1xoED6K6FSYiII-Mg1qiyzt-e-5vpto6ZxT-k5pyCNKc");
   var sheet = ss.getSheets()[0]; // ”0″ is the first sheet
   var startRow = 1; // First row of data to process-actual row#
   var numRows = sheet.getLastRow(); // Number of rows to process
   var numColumns = sheet.getLastColumn();
   var dataRange = sheet.getRange(startRow, 1, numRows, numColumns) // row2, col1, rowlast, colLast
   // Fetch values for each row in the Range.
   var data = dataRange.getValues();
 
   // Check every row with data.
   var i = 1;
   for (i; i <= numRows-1; i++) {
     var row = data[i];
     var dueDate = new Date(row[6]);
     dueDate.setDate(dueDate.getDate()+1);
   
     // If item was not returned, check deadline. Alert value is an option to cease emails to client while we investigate their special circumstance.
     if(String(row[1]).toLowerCase() != "yes" && String(row[1]).toLowerCase() != "alert"){
       
       // If deadline exceeded, write email.
       if(dueDate.getTime() < thisDate.getTime()){
         var emailAddress = "LABUSER@gmail.com";
         var message = "The following person has exceeded their equipment rental deadline: " + row[2] + " " + row[3] +
           ". These items were due by " + dueDate.toLocaleDateString() + ": " + markedLate(); // pulls the topic from the sixth column for the email message Abe edit string
         var subject = "Rental Deadline Exceeded"; // makes the subject what is in quotes
       
         MailApp.sendEmail(emailAddress, subject, message); // send email to lab assistant.
         
         emailAddress = "BOSSUSER@rowan.edu";
         MailApp.sendEmail(emailAddress, subject, message); // send email to Boss.
         
         
       
         emailAddress = row[5];
         message = "Hello. Our records state that you have not yet returned your rented equipment from the Westby Print Lab. " + '\n' +
         "Please return the rented equipment immediately to Westby room 200." + '\n' + '\n' + "The following has yet to be returned: " +  markedLate() + '\n' +
           '\n' + "If the equipment is not returned ASAP, the Art Department will place a hold on your Rowan University account.";
         
         MailApp.sendEmail(emailAddress, subject, message); // send email to client.
       
       } // end if
      } // end if
     
       
       // If last item was returned, paint row green.
       else if(String(row[1]).toLowerCase() == "yes" && !green(2)){
         for (var n = 1; n < numColumns; n++){
           var greenCells = dataRange.getCell(i+1, n);
           greenCells.setBackground('#d9ead3'); // paint cell light green 3
         } // end for
     } // end else     
   } // end for
   
   /**
   * Returns true if value at given cell is green, false otherwise.
   * @param y = column
   */
   function green (y){
     var cell = dataRange.getCell(i+1, y); // get current cell
     for(var n = 0; n < greens.length; n++){ // check if cell matches every green.
       if(cell.getBackground().equals(greens[n]))
         return true
     }
     return false;
   } // end green
   
   /* 
   * Marks all items that have not yet been returned, late (red).
   * Returns data value of items that have not been returned so that it may be displayed in email.
   */
   function markedLate(){
     var redCell;
     var equipment = [];
     var equipmentDue;
     var equipName = data[0];
     // Begins check from first rentable equipment.
     for(var y = 8; y < numColumns; y++){
       if(equipName[y] == "Comments:") // skip comments column
         y++;
       
       if(String(row[y]).length != 0 && !green(y+1)){  // If cell is not empty and not green
         redCell = dataRange.getCell(i+1,y+1);         
         redCell.setBackground('#ff0000'); // paint cell red
           
         if(y < 15 || y > 17) // Return the title name of items.
           equipment[equipment.length] = equipName[y];
         else // Return the item name of batteries, misc, and lens.
           equipment[equipment.length] = row[y];
          
       } // end if
     } // end for
     
     equipmentDue = equipment.join(", ");
     return equipmentDue; // returns a list of what still needs to be returned.
   } // end markedLate
 
 } // end deadlineEmail


/**
* This is a list of content within each row.

  0 - timestamp
  1 - returned
  2 - fname
  3 - lname
  4 - phone#
  5 - email
  6 - duedate
  7 - labassistant
  8+- equipment
  
  Colors used:
  green: #00ff00
  dark green 1: #6aa84f
  dark green 2: #38761d
  dark green 3: #274e13
  light green 1: #93c47d
  light green 2: #b6d7a8
  light green 3: #d9ead3
* red: #ff0000
**/