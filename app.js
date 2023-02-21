/**
 * Sends a email message to the client who completed the google form.
 * To learn more about sending an Email via App Script and Sheets: 
 * https://max-brawer.medium.com/learn-to-magically-send-emails-from-your-google-form-responses-8bbdfd3a4d02
 *
 * @param {Event Object} e - Event Object to call with the function
 */

function onFormSubmit(e) {
    var values = e.namedValues; // get all the client form responses in a Key Pair list 
    var htmlBodyObj = '<h3> Assalamo alaikum, <br>Jazak\'Allah for submitting request for I`tikaf, kindly see below:</h3> <ul>';
   
    //iterate through the key/value pairs and save them into a sudo html object for later
    for (Key in values) {
      var label = Key;
      var data = values[Key];
      htmlBodyObj += '<li>' + label + ": " + data + '</li>';
    };
    htmlBodyObj += '</ul>';
  
    // Send the email
    MailApp.sendEmail({
      to: '<myemail@domain.org>,'+values['Email Address'],
      replyTo: '<myemail@domain.org>',
      subject: '<Objective?> Form responses',
      htmlBody: htmlBodyObj
    });
  }