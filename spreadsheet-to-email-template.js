//MODULE DESCRIPTION
//copy pasta template for automated email sending from spreadsheets, specifically for invoices
//script assumes column structure: email, item description, supplier, item-cost, invoice link
//however it's pretty clear how to mod if something else..

//USAGE NOTES
//replace <content> with your info and repeat for all "TODO" sections

// Examples:
//    var firstRow = <first-row-as-integer>; -> var firstRow = 1;
//    var emailString = <email-address-as-string>; -> var emailString = "someperson@email.com";


function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var mainInvoiceSpreadsheet = <main-spreadsheet-url-as-string>; //TODO add main spreadsheet url
  var emailAddress = <email-address-as-string>; //TODO main to email address
  var ccEmailAddrses = <cc-email-address-as-string>; //TODO cc email address if desired

  //TODO Fill specify ranges (see EOL TODOs for instructions)
  // specify row range
  var startRow = <first-row-as-integer>;  // TODO First row of data to process
  var endRow = <last-row-as-integer>; // TODO last row as integer
  // specify column range
  var startColumn = <starting-column-as-int-1-indexed>; //TODO leftmost column
  var endColumn = <ending-column-as-int-1-indexed>; //TODO rightmost column

  // Fetches values for each row in the Range.
  var numRows = (endRow - startRow + 1);   // calculates the number of rows to process
  var numColumns = (endColumn - startColumn + 1); // " " for columns
  
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns);
  var data = dataRange.getValues(); // data gets stored here
  
  for (i in data) {
    var row = data[i];
    var date = new Date(row[0]);
 
    row[0] = (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear();
   
    var subject = "This is Email Number: " + (parseInt(i)+1); //TODO replace with your subject
  
    var message = "Purchase made on: " + row[0] 
    + "\n" + "Purchase made from: " + row[2] 
    + "\n" + "Purchase description: " + row[1]
    + "\n" + "Purchase amount: $"  + row[3]
    + "\n\n" + "link to invoice: " + row[4]
    + "\n\n" + "Main spreadsheet link follows: " + mainInvoiceSpreadsheet;
    
    MailApp.sendEmail(emailAddress, subject, message, {cc: ccEmailAddress});

    Utilities.sleep(2000); //wait before next iteration to prevent email race conditions..
  }
}
