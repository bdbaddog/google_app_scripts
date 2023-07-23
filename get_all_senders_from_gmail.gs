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

function getValuesFromSpreadsheet(aUnique) {
  const row_end = checkRowEnd()
  aUnique = SpreadsheetApp.getActiveSheet().getRange(2, 1, row_end+2, 2).getValues()
  console.log(aUnique)
  return aUnique
}

function addToSpreadsheet(aUnique) {
  const ssProp = PropertiesService.getDocumentProperties()
  var row_end = 2
  try {
    row_end = Number(ssProp.getProperty('ROW_END'))
    console.log("Row count: %s", row_end)
  } catch (err) {
    console.log(' Failed reading ROW_END : %s', err.messages)
  }
  // add data to sheet
  SpreadsheetApp.getActiveSheet().getRange(2, 1, aUnique.length, 2)
    .setValues(aUnique);
  try {
    const new_count_str = Utilities.formatString("%d",aUnique.length)
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
  // sort and filter for unique entries  
  var aEmails = eMails.sort().filter(function(el,j,a)
    {if(j==a.indexOf(el))return 1;return 0});  

  // create 2D-array
  var aUnique = new Array(); 
   
  var parts = new Array();
  var counter = 0;
  var name;
  var email_address;
  for(var k in aEmails) {
    // this contains something like '"Lori B." <lbb@verizon.net>'
    var this_email = aEmails[k];
    if (this_email.includes(" <")) {
      parts = this_email.split(" <")
      name = String(parts[0]).replace('"','');
      name = name.replace('"','');
      email_address = String(parts[1]).replace(">","");
    } else {
      name = "";
      email_address = this_email;
    }
    aUnique.push([name, email_address]);
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

  var aUnique = new Map();
  // aUnique=getValuesFromSpreadsheet()

  var start_month = 0
  var end_month = 2
  for (i=0; i< end_month; i++) {
    var newStart = new Date(start.getFullYear(), start.getMonth()+i,1);
    var newEnd = new Date(start.getFullYear(), start.getMonth()+i+1,1);

    var filter="after:"+ newStart.toLocaleDateString('en-US', {year: 'numeric', month: '2-digit', day: '2-digit'})+
    " before:"+newEnd.toLocaleDateString('en-US', {year: 'numeric', month: '2-digit', day: '2-digit'});

    console.log(filter);
    var newEmails = getEmails(filter);
    console.log("Found "+newEmails.length+" senders");
    console.log(newEmails);
    // aUnique = aUnique.set[(newEmails);
  }

  aUnique = uniquifyEmails(aUnique);
  addToSpreadsheet(aUnique);


}