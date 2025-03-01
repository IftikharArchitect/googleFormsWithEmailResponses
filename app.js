/**
 * Sends a email message to the client who completed the google form.
 * To learn more about sending an Email via App Script and Sheets: 
 * https://max-brawer.medium.com/learn-to-magically-send-emails-from-your-google-form-responses-8bbdfd3a4d02
 *
 * @param {Event Object} e - Event Object to call with the function
 */
function myFunction() {
  
}

function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  
  // Column index where you want to place the Approval Number
  // (Assuming first form question is in column A, second in B, etc.
  // Adjust index as needed based on your sheet structure.)
  const approvalColIndex = 12; 
  
  // Generate a short ID. 
  // Option 1: A simple sequential number plus year 
  // Option 2: A random 6-character code

  // Get the current year dynamically
  const currentYear = new Date().getFullYear();
  
  // Example: "CAN-2024-" plus row number for uniqueness:
  const uniqueID = "CAN-" + currentYear + "-" + (row - 1); // Subtracting 1 if there's a header row
  
  // Place the ID in the designated cell
  sheet.getRange(row, approvalColIndex).setValue(uniqueID);

  sendEmailOnFormSubmit(e);
}


/**
 * Sends an email with form submission data formatted as a vertical table
 * This function should be triggered by the Form Submit event
 */
function sendEmailOnFormSubmit(e) {
  // If triggered manually without event object, exit
  if (!e) {
    Logger.log("No event object provided. This function should be triggered by a form submission.");
    return;
  }

  try {
    const jamaatColIndex = 11;
    const sheet = e.range.getSheet();
    const row = e.range.getRow();

    // Get the first row (headers)
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Get all values from the current row
    var rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    const jamaatValue = sheet.getRange(row, jamaatColIndex).getValue();
    
    // Get timestamp of submission (using current time if not available)
    const timestamp = new Date();
    const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "MMM dd, yyyy 'at' hh:mm a");
    
    // Use sheet name as form title if actual form title not available
    const formTitle = sheet.getName() || "Form Submission";
    
    // Build the email body with a vertical table
    let emailBody = `<h3>Poster Review - ${jamaatValue} - New Submission: ${formTitle}</h3>`;
    emailBody += `<p>Submitted on: ${formattedDate}</p>`;
    emailBody += "<table style='border-collapse: collapse; width: 100%; border: 1px solid #dddddd;'>";
    
    // Add each form field as a row in the table
    for (let i = 0; i < headerRow.length; i++) {
      // Skip empty headers or values
      if (!headerRow[i]) continue;
      
      const question = headerRow[i];
      const answer = rowValues[i] || "";
      
      // Format the answer based on type
      let formattedAnswer = answer;
      
      // If answer is an array (like for checkboxes), join with commas
      if (Array.isArray(answer)) {
        formattedAnswer = answer.join(", ");
      }
      
      // Add row to table with alternating background colors
      const bgColor = i % 2 === 0 ? "#f2f2f2" : "#ffffff";
      emailBody += `
        <tr style='background-color: ${bgColor};'>
          <td style='padding: 8px; border: 1px solid #dddddd; font-weight: bold; width: 30%;'>${question}</td>
          <td style='padding: 8px; border: 1px solid #dddddd;'>${formattedAnswer}</td>
        </tr>
      `;
    }
    
    emailBody += "</table>";
    emailBody += "<p>This is an automated message. Please do not reply to this email.</p>";
    
    // Send the email
    MailApp.sendEmail({
      to: "iftikhar.ahmed@ahmadiyya.ca, gs@ahmadiyya.ca",
      subject: `Poster Review - Jama'at ${jamaatValue} - New Form Submission: ${formTitle}`,
      htmlBody: emailBody
    });
    
    Logger.log("Email sent successfully");
    
  } catch (error) {
    Logger.log("Error sending email: " + error.toString());
    
    // Optional: Send error notification to admin
    MailApp.sendEmail({
      to: "iftikhar.ahmed@ahmadiyya.ca, gs@ahmadiyya.ca",
      subject: "Error in Form Submission Email Script",
      body: "There was an error processing a form submission: " + error.toString()
    });
  }
}
