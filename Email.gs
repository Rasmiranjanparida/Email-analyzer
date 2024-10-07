function analyzeEmailsFromPastYear() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear(); // Clear the sheet for fresh data

  var endDate = new Date();
  var startDate = new Date();
  startDate.setFullYear(endDate.getFullYear() - 1); // Set start date to one year ago

  var formattedStartDate = formatDate(startDate);
  var formattedEndDate = formatDate(endDate);

  // Get all threads within the past year
  var threads = GmailApp.search('after:' + formattedStartDate + ' before:' + formattedEndDate);
  var emails = [];

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var from = messages[j].getFrom();
      emails.push(from);
    }
  }

  var domainCounts = {};

  emails.forEach(function(email) {
    var domain = email.match(/@([\w.-]+)/);
    if (domain && domain[1]) {
      domain = domain[1];
      if (domainCounts[domain]) {
        domainCounts[domain]++;
      } else {
        domainCounts[domain] = 1;
      }
    }
  });

  var data = [["Domain", "Count"]];
  for (var domain in domainCounts) {
    data.push([domain, domainCounts[domain]]);
  }

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

// Helper function to format date to yyyy/MM/dd
function formatDate(date) {
  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);
  return year + '/' + month + '/' + day;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Analysis')
      .addItem('Analyze Past Year Emails', 'analyzeEmailsFromPastYear')
      .addToUi();
}
