// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function emailFunction() {
  
    var sheet = SpreadsheetApp.getActiveSheet();
  
    var startRow = 2; // First row of data to process
    var numRows = 20; // Number of rows to process
    // Fetch the range of cells - we limit the number of rows to save runtime
    var dataRange = sheet.getRange(startRow, 1, numRows, 7);
    // Fetch values for each row in the Range.
    var data = dataRange.getValues();

    for (var i = 0; i < data.length; ++i) {

        var htmlBody =  HtmlService.createTemplateFromFile('pitch_email');
  

        // Send an email with two attachments: a file from Google Drive (as a PDF) 
        var file = DriveApp.getFileById('googledocsidIsHere123456');

        var row = data[i];
        var emailAddress = row[5]; 
        var recipientName = row[4].split(' ').slice(0, 1);
        var position = row[1];
        var company = row[2];
        var emailSent = row[0]; // Third column;
        var subject = position + " Serge";
        
      // here goes the placeholders that go to email body
      htmlBody.recipientname_applied = recipientName;
      htmlBody.companyname_applied = company;
      htmlBody.position_applied = position;
    
    
      // before I used .getContent() straight away, but now i replaces - content will be created from evaluations
        var email_html = htmlBody.evaluate().getContent();
  
      
        if (emailSent !== EMAIL_SENT ) { // Prevents sending duplicates
            var subject = position + " - Serge";

            MailApp.sendEmail({
                to: emailAddress,
                subject: position + ' Serge',
                htmlBody: email_html,   // before we use htmlBody - but to add things one more step was done
                attachments: [file.getAs(MimeType.PDF)]
            });

            sheet.getRange(startRow + i, 1).setValue(EMAIL_SENT);
            // Make sure the cell is updated right away in case the script is interrupted
            SpreadsheetApp.flush();
          
        }
    }
}

// then in html file we implement variables <?= recipientname_applied ?> and ets. 
