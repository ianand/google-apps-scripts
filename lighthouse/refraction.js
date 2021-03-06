// REFRACTION
// A Google Apps Script that imports Lighthouse (lighthouseapp.com) tickets into
// a spreadsheet. The script accepts a lighthouse query and dumps the ticket
// information at the currently selected cell.

// Configuration: Set variables (1) through (3). If you don't change these
// values you'll get the test data I used to write the script.

// (1) Lighthouse Token. You'll need to generate this on your lighthouse account page.
var LIGHTHOUSE_TOKEN = "017f0e53e5f8665ee2b8cd94730b8c60110ce5e2";

// (2) Project Id. e.g. the "55411" in http://foo.lighthouseapp.com/projects/55411-foobar/overview
var PROJECT_ID = "55411"; 

// (3) Subdomain. e.g. the "foo" in http://foo.lighthouseapp.com
var SUBDOMAIN = "refraction"; 

// Ticket object. This is a Factory object. Tickets are created by
// this object by calling Tickets.createTicketFromXml(). Do not
// use new().
var Ticket = {
  
  // "OrderedHash" of ticket field names. 
  // We need a list of ticket fields and that list to have a consistent
  // order. Could have just used an array but wanted to have better performance.
  ticketFieldsHash: {},
  ticketFieldsArray: [],
  saveTicketFieldName: function (name) {
    if(!this.ticketFieldsHash[name]) {
      this.ticketFieldsHash[name] = true;
      this.ticketFieldsArray.push(name);
    }
  },  
  
  // Parses a <ticket> XML object and returns a javascript Object representation
  createTicketFromXml: function (xml) {
    // Create a new ticket "instance".
    var ticket = {};

    // Helper method to assign fields to the new instance from XML
    var that = this; // helper this pointer
    function extractXmlField(xmlElement) {
      try {
        // use lowercase version so we can easily match the name
        // even if the user used a different case
        var fieldName = xmlElement.getName().getLocalName().toLowerCase();
        var fieldValue = xmlElement.getText();
        
        // Save the field to our new instance.
        ticket[fieldName] = fieldValue;
        
        // Remember the field name for later.
        that.saveTicketFieldName(fieldName);
      } catch(err) {
        Logger.log("Error when extracting property: " + err );
      }
    }
    
    var fields = xml.getElements();
    if(fields != null && fields.length > 0) {
      var length = fields.length;
      for(var i = 0; i < length; i++) {
       extractXmlField(fields[i]);
      }
    }
    return ticket;    
  }
};

// Queries lighthouse API and dumps the list of tickets at the actively
// selected cell.
function importLighthouseTickets() {
    
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
      result.push(Ticket.createTicketFromXml(tickets[i]));
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
    
    // Default to printing all the ticket properties.
    var ticketFields = Ticket.ticketFieldsArray;
    var ticketFieldsLength = ticketFields.length;
    
    // If the active cell isn't empty than treat that row
    // as the list of properties to print.
    var activeCellContent = sheet.getRange(row0, column0).getValue();
    if(activeCellContent != null && activeCellContent != "") {
      ticketFields = [];
      var i = 0;
      while(activeCellContent != null && activeCellContent != "" ) {

        // Save this as a property name to print
        ticketFields.push(activeCellContent.toLowerCase());
        
        // Read the property name in the next column
        i++;
        activeCellContent = sheet.getRange(row0, column0+i).getValue();

      }
      
      // ticketFields has changed so update the length
      ticketFieldsLength = ticketFields.length;
    } else {

      // If the active cell is empty than print the default labels for the ticket 
      // properties.
      for(var i = 0; i < ticketFieldsLength; i++) {
        sheet.getRange(row0, column0+i).setValue(Ticket.ticketFieldsArray[i]);
      }
    }
    

    //
    // Helper functions for Spreadsheet operations. Ideally
    // I'd make these properties of the sheet object itself
    // but Google Apps Script doesn't allow you to add new 
    // properties to system defined objects.
    //
    function isRowEmpty(row) {
      var rowValues = (sheet.getRange(row, column0, row, column0 + ticketFieldsLength).getValues())[0];
      return "" == rowValues.join("");
    }
    var ticketNumberIndex = ticketFields.indexOf("number");
    function ticketNumberForRow(row)  {
      if(ticketNumberIndex < 0) {
        return -1;
      }
      return sheet.getRange(row, column0 + ticketNumberIndex).getValue();
    }
    function writeTicketToRow(ticket, row) {
      for(var j = 0; j < ticketFieldsLength; j++) {
        var fieldName = ticketFields[j];
        var fieldValue = ticket[fieldName];
        if(fieldValue != undefined) {
          sheet.getRange(row, column0+j).setValue(fieldValue);
        }
      }  
    }
    
    //
    // Helper functions for the tickets array.
    //
    function indexOfTicketNumber(ticketNumber) {
      var length = tickets.length;
      for(var i = 0; i < length; i ++) {
        if(tickets[i].number == ticketNumber) {
          return i;
        }
      }
      return -1;
    }
    
    //
    // Output the tickets to the spreadsheet.
    //

    // Start output of ticket data on the next line
    var currentRow = row0+1;
    
    // Write out the ticket data to the spreadsheet
    while(tickets.length > 0) {
      
      // If the current row is empty, then just output the next ticket
      // in the tickets array.
      if(isRowEmpty(currentRow)) {
        var ticket = tickets.shift();
        writeTicketToRow(ticket, currentRow);
      } else {
        
        // If there's ticket info already here, write out the data that
        // corresponds to this row. The corresponding ticket is found
        // by ticket number.
        var ticketNumber = ticketNumberForRow(currentRow);
        var index = indexOfTicketNumber(ticketNumber);
        if(index >= 0) {
          var ticket = tickets.splice(index, 1)[0];
          writeTicketToRow(ticket, currentRow);
        }
      }
      
      // Try output on the next row of the sheet
      currentRow++;
    }    
  } 
  
  var query = Browser.inputBox("Lighthouse search query (e.g. 'test state:open' )");
  var data = getTicketData(query);
  var tickets = parseTicketData(data);
  writeToSpreadsheet(tickets);  
}

// TODO:
//
// - Query user for token, project id, subdomain only once (by saving info in spreadsheet for later use).
// - Parse out ticket data for "meta data" in tickets to indicate time spent and estimated completion.


​