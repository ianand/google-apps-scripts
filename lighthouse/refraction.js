// REFRACTION
// A Google Apps Script that imports Lighthouse (lighthouseapp.com) tickets into
// a spreadsheet.

// Configuration: Set variables (1) through (3). If you don't change these
// values you'll get the test data I used to write the script.

// (1) Lighthouse Token. You'll need to generate this on your lighthouse account page.
var LIGHTHOUSE_TOKEN = "017f0e53e5f8665ee2b8cd94730b8c60110ce5e2";

// (2) Project Id. e.g. the "55411" in http://foo.lighthouseapp.com/projects/55411-foobar/overview
var PROJECT_ID = "55411"; 

// (3) Subdomain. e.g. the "foo" in http://foo.lighthouseapp.com
var SUBDOMAIN = "refraction"; 

function importLighthouseTickets() {

  var query = Browser.inputBox("Lighthouse search query (e.g. 'test state:open')");
  var data = getTicketData("test");
  // Logger.log("XML Result " + data);
  var tickets = parseTicketData(data);
  writeToSpreadsheet(tickets);
        
  // Ticket object
  function createTicket(xml) {
    var that = {};
    
    function getProperty(property) {
      var value = "";
      try {
        value = xml.getElement(property).getText();
      } catch(err) {
        Logger.log("Error when looking for "+property+": " + err );
      }
      return value;
    }
    
    that.title = getProperty("title");
    that.bugNumber = getProperty("number");
    that.url = getProperty("url");
    that.state = getProperty("state");
    that.assignedUserName = getProperty("assigned-user-name");
        
    return that;
  }
  
  // Download ticket XML data from lighthouse api
  function getTicketData(query) {
    
    var parameters = "q=" + encodeURIComponent(query) +
        "&_token=" + encodeURIComponent(LIGHTHOUSE_TOKEN);
    
    return UrlFetchApp.fetch("http://" + SUBDOMAIN + ".lighthouseapp.com/projects/" + 
                                 PROJECT_ID + "/tickets.xml?" + parameters).getContentText();
    
  }
  
  // Parse XML data into ticket objects
  function parseTicketData(data) {
    var doc = Xml.parse(data, true);
    var tickets = doc.getElement().getElements("ticket");
    var length = tickets.length;
    var result = [];
    for(var i = 0; i < length; i++) {
      result.push(createTicket(tickets[i]));
    }
    return result;
  }
  
  // Dump a list of tickets into the spreadsheet at the currently active cell.
  function writeToSpreadsheet(tickets) {
    // Get the position of the actively selected cell. 
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange();
    var row0 = range.getRow();
    var column0 = range.getColumn();
    
    // Write the column labels
    sheet.getRange(row0, column0+0).setValue("BugNumber");
    sheet.getRange(row0, column0+1).setValue("Title");
    sheet.getRange(row0, column0+2).setValue("URL");
    sheet.getRange(row0, column0+3).setValue("State");
    sheet.getRange(row0, column0+4).setValue("Assigned User");    
    
    // Write out the ticket data to the spreadsheet
    var length = tickets.length;
    for(var i = 0; i < length; i++) {
      sheet.getRange(row0+1+i, column0+0).setValue(tickets[i].bugNumber);
      sheet.getRange(row0+1+i, column0+1).setValue(tickets[i].title);
      sheet.getRange(row0+1+i, column0+2).setValue(tickets[i].url);
      sheet.getRange(row0+1+i, column0+3).setValue(tickets[i].state);      
      sheet.getRange(row0+1+i, column0+4).setValue(tickets[i].assignedUserName);       
    }
    
  } 
}
​