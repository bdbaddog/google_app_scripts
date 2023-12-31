function clearRowEnd() {
    const ssProp = PropertiesService.getDocumentProperties()
    try {
        const new_count_str = Utilities.formatString("%d",0)
        ssProp.setProperty('ROW_END', new_count_str)
    } catch (err) {
        console.log(' Failed writing ROW_END : %s', err.messages)
    }
  }
  
  function checkRowEnd() {
    const ssProp = PropertiesService.getDocumentProperties()
    var row_end = 0
    try {
      row_end = Number(ssProp.getProperty('ROW_END'))
      console.log("Row count: %s", row_end)
    } catch (err) {
      console.log(' Failed reading ROW_END : %s', err.messages)
    }
    return row_end
  }
  
  function getValuesFromSpreadsheet(unique_emails) {
    const row_end = checkRowEnd()
    console.log('Retrieving %d rows', row_end)
    var aUnique = SpreadsheetApp.getActiveSheet().getRange(2, 1, row_end+2, 2).getValues()
    console.log(aUnique)
  
    aUnique.forEach(function(i) {
      // console.log('From spreadsheet: %s %s',i[0], i[1])
      if (i[0] != '') {
        unique_emails.set(i[0], i[1])
      }
    })
    return unique_emails
  }
  
  function updateSpreadsheet(unique_emails) {
    const emailArray = Array.from(unique_emails)
  
    // Sort by name
    emailArray.sort(function first_element(element1, element2){
      var first = element1[1]
      if (first == '') {
        first = element1[0]
      }
      var second = element2[1]
      if (second == '') {
        second = element2[0]
      }
  
      if (first > second) {
        // console.log("%s > %s",first, second)
        return 1
      }
      if (first < second) {
        // console.log("%s < %s",first, second)
        return -1
      }
      return 0;
    })
  
    const ssProp = PropertiesService.getDocumentProperties()
    var row_end = 2
    try {
      row_end = Number(ssProp.getProperty('ROW_END'))
      console.log("Row count: %s", row_end)
    } catch (err) {
      console.log(' Failed reading ROW_END : %s', err.messages)
    }
    // add data to sheet
    SpreadsheetApp.getActiveSheet().getRange(2, 1, emailArray.length, 2)
      .setValues(emailArray);
    try {
      const new_count_str = Utilities.formatString("%d",emailArray.length)
      ssProp.setProperty('ROW_END', new_count_str)
    } catch (err) {
      console.log(' Failed writing ROW_END : %s', err.messages)
    }
  }
  
  function getEmails(filter) { 
    // http://stackoverflow.com/a/12029701/1536038  
    // get all messages      
    var eMails = GmailApp.getMessagesForThreads(
      GmailApp.search(filter))
        .reduce(function(a, b) {return a.concat(b);})
        .map(function(eMails) {
      return eMails.getFrom() 
    });
    return eMails;
  }
  
  function uniquifyEmails (eMails) {
    // create 2D-array
    var aUnique = new Map(); 
     
    var parts = new Array();
    var counter = 0;
    var name;
    var email_address;
    for(let k of eMails.keys()) {
      // this contains something like '"Lori B." <lbb@verizon.net>'
      var this_email = k;
      if (this_email.includes(" <")) {
        parts = this_email.split(" <")
        name = String(parts[0]).replace('"','');
        name = name.replace('"','');
        email_address = String(parts[1]).replace(">","");
      } else {
        name = "";
        email_address = this_email;
      }
      aUnique.set(email_address, name)
      counter++;
      if (counter % 10 == 0)  {
        console.log("Messages Processed: %d",counter);
      }
    }
    return aUnique;
  }
  
  function GetAllUniqueEmailsFromInbox() {
    var start = new Date(2019, 8, 1);
    var end = new Date();
    var months = (end.getMonth() - start.getMonth()) + (12 * (end.getFullYear() - start.getFullYear()));
    console.log("Months "+ months );
  
    var unprocessed_senders = new Map();
  
    var start_month = 40
    var end_month = 46
    for (i=start_month; i< end_month; i++) {
      var newStart = new Date(start.getFullYear(), start.getMonth()+i,1);
      var newEnd = new Date(start.getFullYear(), start.getMonth()+i+1,1);
  
      var filter="after:"+ newStart.toLocaleDateString('en-US', {year: 'numeric', month: '2-digit', day: '2-digit'})+
      " before:"+newEnd.toLocaleDateString('en-US', {year: 'numeric', month: '2-digit', day: '2-digit'});
  
      console.log(filter);
      var newEmails = getEmails(filter);
      console.log("Found "+newEmails.length+" senders");
      console.log(newEmails);
      for (var e in newEmails) {
        unprocessed_senders.set(newEmails[e],true)
      }
    }
  
    var unique_emails = uniquifyEmails(unprocessed_senders);
    unique_emails=getValuesFromSpreadsheet(unique_emails)
  
    updateSpreadsheet(unique_emails);
  
  
  }