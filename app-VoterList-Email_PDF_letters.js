//template google document for voters (fill with their name)


function sendDrafts() {
  var draftMessages = GmailApp.getDrafts();
  for (var i = 0; i < draftMessages.length; i++) {
    

      var gDraftMessage = draftMessages[i];
      var gMessageToField = gDraftMessage.getMessage().getTo();
      console.log(i + ": sending draft for: " + gMessageToField);
      try {
        gDraftMessage.send();
      } catch (e) {
        console.error('sendDrafts() yielded an error: ' + e);
      }
      console.log(i + ": done sending draft for: " + gMessageToField);
      // Add delay between emails (e.g. 1000ms = 1 second)
      Utilities.sleep(2000);

    
    
  }
}
var rootPDFFolderID = '13wnJONfgo0yomOe-TZjdd9QeF0SVV8I5';
var docId ="1VvM5gnO81ADTDU2anGhYL8nw3mdXaTOuzVvAW8HwqDQ";// "1v3tFkDcjAgtRIi9cnJHdJ771g9cbN_0sUML3hzDh96E";

var name_row_num = 2;
var code_row_num = 1;
var address_row_num = 7;
var email_row_num = 8;
var simpleEmailBody = "Assalamo alaikum wa Rahmatullah, \r\n"+
  
  
  "I pray you are in the best of health and spirit, Ameen!\r\n\r\n"+
  
  "Please see the attached letter of Approval as a Voter\r\n\r\nInsh'Allah the elections will take place on Saturday Apr 12th at 5:00pm in Baitul Islam Mosque\r\n\r\n" +
  
  
  "Jazakumullah \r\nWassalam \r\nKhaksar,\r\n\r\n\r\nIftikhar Ahmed\r\n(Humble servant of Khilafat)\r\nServing as General Secretary\r\nVaughan North Halqa, Vaughan Jama`at 2022-2025\r\nM: 416-450-4224\r\nE: gs.vaughannorth@ahmadiyya.ca\r\n\r\n\r\nDisclaimer: The information shared in this email is confidential and may be legally privileged. It is intended solely for the addressee. Access to this email or its attachments by anyone else is unauthorized. If you are not the intended recipient, any disclosure, copying, distribution or any action taken or omitted to be taken in reliance on it, is prohibited and may be unlawful. You may contact the sender if you believe you have received this email in error and delete this message from your system. In addition, please note that email delivery is not guaranteed to be secure or error-free. Messages could be intercepted, corrupted, lost, arrive late or may contain viruses and the sender is not liable for these risks.";


var htmlBodyTxt = "Assalamo alaikum wa Rahmatullah, <br><br>"+
  
  
  "I pray you are in the best of health and spirit, Ameen!<br><br>"+
  
  "Please see the attached letter of Approval as a Voter<br><br>Insh'Allah the elections will take place on:<br><b>Date:Saturday Apr 12th<br>Time:5:00pm<br>Place:Baitul Islam Mosque</b><br><br>" +
  
  
  "Jazakumullah <br>Wassalam <br>Khaksar,<br><br><br>Iftikhar Ahmed<br>(Humble servant of Khilafat)<br>Serving as General Secretary<br>Vaughan North Halqa, Vaughan Jama`at 2022-2025<br>M: 416-450-4224<br>E: gs.vaughannorth@ahmadiyya.ca<br><br><br>Disclaimer: The information shared in this email is confidential and may be legally privileged. It is intended solely for the addressee. Access to this email or its attachments by anyone else is unauthorized. If you are not the intended recipient, any disclosure, copying, distribution or any action taken or omitted to be taken in reliance on it, is prohibited and may be unlawful. You may contact the sender if you believe you have received this email in error and delete this message from your system. In addition, please note that email delivery is not guaranteed to be secure or error-free. Messages could be intercepted, corrupted, lost, arrive late or may contain viruses and the sender is not liable for these risks.";

function testDoc() {
  var name = "Mohammed, Anwar(14040)";
  var email = "gs.vaughannorth@ahmadiyya.ca";

  for (var i = 0; i < 10; i++) {
    createDraft(name+""+i,email+""+i);
  }
  
  
}

function checkDocumentAccess() {
  try {
    var file = DriveApp.getFileById(docId);
    var fileDoc = DocumentApp.getFileById(docId);
    fileDoc.getBody();
    var access = file.getAccess(Session.getActiveUser());
    Logger.log("Current user has " + access + " access");
  } catch(e) {
    Logger.log("Error: " + e.toString());
  }
}

//creat draft documents / pdfs in emails and gdrive 
function testDraft() {
  //createDraft("Iftikhar Ahmed (4365)", "info@iftikhar.org");
  

  
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var counter = 0;
  var rootPDFFolder = DriveApp.getFolderById(rootPDFFolderID);
  for (var i=0 ; i<sheets.length ; i++) {
    //out.push(  "=counta('"+sheets[i].getName()+"'!B:B)"  )

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = sheets[i].getName();
    if (sheetName == 'Cleaned-Voter-list') {
      var sheet = ss.getSheetByName(sheetName);
      var values = sheet.getDataRange().getValues();
      var lastColumnIndex = values.length;
      var initialIndex = 0;
      console.log("sheet name:" + sheetName);
      
      //var sheetFolder = rootPDFFolder.createFolder(sheetName);
      var sheetFolder = rootPDFFolder;

      for (var j = lastColumnIndex - 1; j >= 1; j--) {
        // name (code) - address 
        // email is email
        var name = values[j][name_row_num] + " Sahib (" + values[j][code_row_num] + ") - " + values[j][address_row_num] ;        
        var email = values[j][email_row_num]; ;
        
          var logMessage = "create PDF draft for:" + name +", with email: "+ email;
          counter++;
          //name = counter + "."+name;
          console.log("counter: " + counter + "," +logMessage);
          out.push(logMessage);
          createDraft(name,email);
          createPDFDraftInGDrive(name,sheetFolder.getId(),counter);
        
      }
    }
    
  }
  return out;
}

function createDraft(name, email) {
  appendToFile(name);
  var emailBody = DriveApp.getFileById(docId);
  
  var emailBodyBinary = emailBody.getAs(MimeType.PDF);
  if (email != "#N/A") {
    var gDraft = GmailApp.createDraft(email, 'Approval for Jama`at Office Bearers Elections Apr 12th, 2025 at 5:00pm',simpleEmailBody , {
        attachments: [emailBodyBinary],
        name: 'GS Vaughan North',
        cc: "gs.vaughannorth@ahmadiyya.ca",//cc: "gs.vaughannorth@ahmadiyya.ca, sadr.vaughannorth@ahmadiyya.ca",
        htmlBody: htmlBodyTxt
    });
  }
  
  revertToFile(name);
}

function createPDFDraftInGDrive(name, folderID, counter) {
  appendToFile(name);
  var emailBody = DriveApp.getFileById(docId);
  
  var emailBodyBinary = emailBody.getAs(MimeType.PDF);

  // Set the name of the PDF file to include the 'name' variable
  var pdfFileName = counter +"."+name + ".pdf"; // Customize the filename as needed
  emailBodyBinary.setName(pdfFileName);

  var savingFolder = DriveApp.getFolderById(folderID);
  savingFolder.createFile(emailBodyBinary);

  revertToFile(name);
}

function appendToFile(name) {
  var docBody = DocumentApp.openById(docId);
  var body = docBody.getBody();
  var foundElement = body.findText("--.*?--");
  if (foundElement) {
    var text = foundElement.getElement().asText();
    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();
    text.deleteText(start, end);
    text.insertText(start, "--" + name + "--");
  }
  //Logger.log(docBody.getText());
  docBody.saveAndClose();
}

function revertToFile(name) {
  var docBody = DocumentApp.openById(docId);
  var body = docBody.getBody();
  var foundElement = body.findText("--.*?--");
  if (foundElement) {
    var text = foundElement.getElement().asText();
    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();
    text.deleteText(start, end);
    text.insertText(start, "--fullName--");
  }
  //Logger.log(docBody.getText());
  docBody.saveAndClose();
}

function myFunction() {
  console.log("'"+sheetNames()[0]+"'");
  return (sheetNames());
}

function showEmails() {
  return getEmails();
}

function showNames() {
  return getNames();
}

function FindEmails(input) {
  var regex = /(?:[a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])/gm
  var result = input.match(regex);
  if (result && result.length > 1) {
    var longResult = result[0];
    for (var i =1 ; i < result.length; i++) {
      longResult += "," + result[i];
    }
    result = longResult;
  }
  return result;
}

function sheetNames() {
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) {
    //out.push(  "=counta('"+sheets[i].getName()+"'!B:B)"  )

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = sheets[i].getName();
    if (sheetName != 'Street Leaders') {
      var sheet = ss.getSheetByName(sheetName);
      var values = sheet.getDataRange().getValues();
      var lastColumnIndex = values.length;
      var initialIndex = 0;
      console.log("sheet name:" + sheetName);
      for (var j = lastColumnIndex - 1; j >= 0; j--) {
        var codeValid = values[j][1];
        if (codeValid != "#VALUE!") {
          var memCode = values[j][1] +"," + sheetName +"," + values[j][2]+"," + values[j][3]+"," + values[j][4] +"," + values[j][5]; ;
          console.log(memCode);
          out.push(memCode);
        }
      }
    }
    
  }
  
  return out  
}

function displaySheetNames() {
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) {
    //out.push(  "=counta('"+sheets[i].getName()+"'!B:B)"  )

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = sheets[i].getName();
    if (sheetName != 'Street Leaders') {
      var sheet = ss.getSheetByName(sheetName);
      var values = sheet.getDataRange().getValues();
      var lastColumnIndex = values.length;
      var initialIndex = 0;
      console.log("sheet name:" + sheetName);
      out.push(sheetName);
    }
    
  }
  
  return out  
}

function getEmails() {
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) {
    //out.push(  "=counta('"+sheets[i].getName()+"'!B:B)"  )

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = sheets[i].getName();
    if (sheetName != 'Street Leaders') {
      var sheet = ss.getSheetByName(sheetName);
      var values = sheet.getDataRange().getValues();
      var lastColumnIndex = values.length;
      var initialIndex = 0;
      console.log("sheet name:" + sheetName);
      for (var j = lastColumnIndex - 1; j >= 0; j--) {
        var codeValid = values[j][1];
        if (codeValid != "#VALUE!") {
          var memCode = values[j][5]; ;
          console.log(memCode);
          out.push(memCode);
        }
      }
    }
    
  }
  
  return out  
}

function getPhoneNumbers() {
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) {
    //out.push(  "=counta('"+sheets[i].getName()+"'!B:B)"  )

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = sheets[i].getName();
    if (sheetName != 'Street Leaders') {
      var sheet = ss.getSheetByName(sheetName);
      var values = sheet.getDataRange().getValues();
      var lastColumnIndex = values.length;
      var initialIndex = 0;
      console.log("sheet name:" + sheetName);
      for (var j = lastColumnIndex - 1; j >= 0; j--) {
        var codeValid = values[j][1];
        if (codeValid != "#VALUE!") {
          var memCode = values[j][7]; ;
          console.log(memCode);
          out.push(memCode);
        }
      }
    }
    
  }
  
  return out  
}

function getCellPhoneNumbers() {
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) {
    //out.push(  "=counta('"+sheets[i].getName()+"'!B:B)"  )

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = sheets[i].getName();
    if (sheetName != 'Street Leaders') {
      var sheet = ss.getSheetByName(sheetName);
      var values = sheet.getDataRange().getValues();
      var lastColumnIndex = values.length;
      var initialIndex = 0;
      console.log("sheet name:" + sheetName);
      for (var j = lastColumnIndex - 1; j >= 0; j--) {
        var codeValid = values[j][1];
        if (codeValid != "#VALUE!") {
          var memCode = values[j][7];
          var outNum = "";
          var numbers = memCode.toString().split(" ");
          for (var k = 0; k < numbers.length; k++) {
            var tempNum = numbers[i];
            if ( tempNum != undefined && (tempNum.startsWith("647") || tempNum.startsWith("416")) ) {
              outNum += tempNum +",";
            }
          }
          outNum = values[j][6] + " - " + outNum;
          console.log(outNum);
          out.push(outNum);
        }
      }
    }
    
  }
  
  return out  
}

function getNames() {
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) {
    //out.push(  "=counta('"+sheets[i].getName()+"'!B:B)"  )

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = sheets[i].getName();
    if (sheetName != 'Street Leaders') {
      var sheet = ss.getSheetByName(sheetName);
      var values = sheet.getDataRange().getValues();
      var lastColumnIndex = values.length;
      var initialIndex = 0;
      console.log("sheet name:" + sheetName);
      for (var j = lastColumnIndex - 1; j >= 0; j--) {
        var codeValid = values[j][1];
        if (codeValid != "#VALUE!") {
          var rawString = values[j][6]; ;        
          
          
          console.log(rawString);
          out.push(rawString);
        }
      }
    }
    
  }
  
  return out  
}


function showNameAndAddresses() {
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) {
    //out.push(  "=counta('"+sheets[i].getName()+"'!B:B)"  )

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = sheets[i].getName();
    if (sheetName != 'Street Leaders') {
      var sheet = ss.getSheetByName(sheetName);
      var values = sheet.getDataRange().getValues();
      var lastColumnIndex = values.length;
      var initialIndex = 0;
      console.log("sheet name:" + sheetName);
       //out.push("sheet name:" + sheetName);
      for (var j = lastColumnIndex - 1; j >= 0; j--) {
        var codeValid = values[j][1];
        if (codeValid != "#VALUE!") {
          var rawString = values[j][3] + '\r\n' + values[j][4] ;        
          
          
          console.log(rawString);
          out.push(rawString);
        }
      }
    }
    
  }
  
  return out  
}

function showAddresses() {
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) {
    //out.push(  "=counta('"+sheets[i].getName()+"'!B:B)"  )

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = sheets[i].getName();
    if (sheetName != 'Street Leaders') {
      var sheet = ss.getSheetByName(sheetName);
      var values = sheet.getDataRange().getValues();
      var lastColumnIndex = values.length;
      var initialIndex = 0;
      console.log("sheet name:" + sheetName);
       //out.push("sheet name:" + sheetName);
      for (var j = lastColumnIndex - 1; j >= 0; j--) {
        var codeValid = values[j][1];
        if (codeValid != "#VALUE!") {
          var rawString = values[j][4] ;        
          
          
          console.log(rawString);
          out.push(rawString);
        }
      }
    }
    
  }
  
  return out  
}




function isEastOrWest(street1, street2, town, postalCodePrefix) {
// Helper function to get coordinates from Nominatim API
function getCoordinates(street, town, postalCodePrefix) {
  let queries = [];
  // Build a list of queries to try
  queries.push(`${street}, ${town}, ${postalCodePrefix}, Canada`);
  queries.push(`${street}, ${town}, Canada`);
  queries.push(`${street}, Canada`);

  for (let i = 0; i < queries.length; i++) {
    const query = queries[i];
    const url = `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(query)}`;
    console.log(`Trying query: ${query}`);
    try {
      const response = UrlFetchApp.fetch(url, {
        headers: {
          "User-Agent": "YourAppName/1.0 (your.email@example.com)",
        },
        muteHttpExceptions: true,
      });
      const data = JSON.parse(response.getContentText());

      if (data && data.length > 0) {
        // Return the longitude as a number
        console.log(`Found coordinates for ${query}`);
        return parseFloat(data[0].lon);
      }
    } catch (error) {
      console.error(`Error fetching coordinates for ${query}:`, error);
    }
  }
  // If all queries fail
  throw new Error(`No coordinates found for ${street}, ${town}, ${postalCodePrefix}, Canada`);
}

try {
  // Update town from 'Maple' to 'Vaughan' for better recognition
  const correctedTown = town === 'Maple' ? 'Vaughan' : town;

  // Get longitudes for both streets
  const lon1 = getCoordinates(street1, correctedTown, postalCodePrefix);
  const lon2 = getCoordinates(street2, correctedTown, postalCodePrefix);

  console.log(`Longitude of ${street1}: ${lon1}`);
  console.log(`Longitude of ${street2}: ${lon2}`);

  if (lon1 > lon2) {
    //return `${street1} is east of ${street2}`;
    return 'east';
  } else if (lon1 < lon2) {
    //return `${street1} is west of ${street2}`;
    return 'west';
  } else {
    //return `${street1} and ${street2} are at the same longitude`;
    return 'west';
  }
} catch (error) {
  console.error(`An error occurred: ${error.message}`);
  return `An error occurred: ${error.message}`;
}
}



// Example usage:
function getAnsarHalqa() {

  const streetNames = ["America Ave",
  "AMERICA AVE",
  "Cape Verde Way",
  "Chart Ave",
  "Convoy Cres",
  "Discovery Trail",
  "Equator Cres",
  "Equator Crescent",
  "Ferdinand Ave",
  "Genoa Dr",
  "Genoa Rd",
  "Gully Lane",
  "John Deisman Blvd",
  "Journal Ave",
  "Madeira Ave",
  "Naples Ave",
  "Native Trail",
  "Ocean Ave",
  "Stern Gate",
  "Sail Cres",
  "Santa Maria Trail",
  "Treasure Rd",
  "Treasure Rd (Basement)",
  "Treasure Road",
  "Windy Way"];
  const referenceStreet = "Discovery Trail";
  const town= "Maple";
  const postalCodePrefix= "L6A";

  for (const currentStreet of streetNames) {
    const result = isEastOrWest(
      currentStreet,
      referenceStreet,
      town,
      postalCodePrefix
    );
    
    console.log(result);
  }
}



function getAnsarHalqaEW(currentStreet) {

  const referenceStreet = "Discovery Trail";
  const town= "Maple";
  const postalCodePrefix= "L6A";

  
    const result = isEastOrWest(
      currentStreet,
      referenceStreet,
      town,
      postalCodePrefix
    );
    
    console.log(result);
    return result;
  
}
