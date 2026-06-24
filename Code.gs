// To run this script:
// 1. Open 'https://docs.google.com/spreadsheets/d/1fkQOOJG_5CQg3BMHa4NyBYjrFmIgiB-56i1ziUXhT00/edit?gid=451675374#gid=451675374' in another window; visible concurrent with this one for best results
// 2. NO LONGER NECESSARY: Select the "Auto descriptions" tab, so that the getActiveSheet() works.
// 3. Press the Run button in the Apps Script editor
// 4. Wait for the Alert() dialog to show up in the spreadsheet browser window.
// 5. Copy/paste load that new document in yet another window
// 6. Press "OK"
// 7. Wait a few seconds, then the content should show up in that new document.
function generateGroupedDoc() {
  //var sheetdoc = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1fkQOOJG_5CQg3BMHa4NyBYjrFmIgiB-56i1ziUXhT00");
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //var sheetbios = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //var sheet = sheetdoc.getSheetByName("Auto descriptions");
  
  var biossheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bios");
  var biosdata = biossheet.getDataRange().getValues();
  var biosrows = biosdata.slice(1);
  var biosmap = {};
  biosrows.forEach(function(row) {
    // Presenter name
    var bioname = row[0];
    var biodraft = row[1];
    if(!bioname) return;
    if(!biodraft){
      biodraft = row[2];
    }
    // if(!biosmap[bioname]) biosmap[bioname] = [];
    biosmap[bioname] = biodraft;
  });

  var workshopdescriptionssheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auto descriptions");
  var data = workshopdescriptionssheet.getDataRange().getValues();

  // Skip the header row (assumes row 1 is headers)
  var rows = data.slice(1);

  // 2. Group the data using an object map
  var presenterWorkshops = {};
  var workshopDetail = {};
  var workshopDescription = {};
  var workshopPrimary = {};
  
  rows.forEach(function(row) {
    // The spreadsheet-derived primary presenter, use for grouping workshop descriptions in the program
    var presentername = row[0];
    // The original list of presenters to be billed with it in the sheet
    var presenterlist = row[1];
    var workshoptitle = row[2];
    var workshoplocation = row[3];
    var workshopdate = row[4];
    var workshoptime = row[5];
    var workshopdescription = row[6]; 
    var workshopdescriptiondraft = row[7]; 

    // ignore empty
    if(!workshoptitle) return;

    // Create entry in the array as needed for each presenter's name with workshops
    if (!presenterWorkshops[presentername]) {
      presenterWorkshops[presentername] = [];
    }

    // Use the cleaner description, if available
    if(!workshopdescription) workshopdescription = workshopdescriptiondraft;

    // Add the title to its respective group
    //presenterWorkshops[presentername].push("<b>" + workshoptitle + "</b>: " + workshopdate + " " + workshoptime + " " + workshoplocation + "\n" + workshopdescription);
    presenterWorkshops[presentername].push(workshoptitle);
    workshopPrimary[workshoptitle] = presentername;
    //workshopDetail[workshoptitle] = ": " + workshopdate + " " + workshoptime + " " + workshoplocation + "\n" + workshopdescription + "\nPresented by: " + presenterlist;
    workshopDetail[workshoptitle] = ": " + workshopdate + " " + workshoptime + " " + workshoplocation;
    var alsopresentedby = "";
    if(presenterlist.match(/,/)){
      alsopresentedby = "\nAlso presented by: " + presenterlist;
      alsopresentedby = alsopresentedby.replace(/,[^,]*$/,"");
    }
    workshopDescription[workshoptitle] = "\n" + workshopdescription + alsopresentedby;

    // TODO: Parse out presenterlist, initialize each presenter as needed, insert "* Workshop foo, loc, date, time: see _primary_presenter_" for each secondary presenter
    var presenters = presenterlist.split(/, */);
    presenters.forEach(function(otherpresenter){
      if (presentername === otherpresenter){
         return;
      }
      if (!presenterWorkshops[otherpresenter]) presenterWorkshops[otherpresenter] = [];
      presenterWorkshops[otherpresenter].push(workshoptitle);
    });
  });

  // 3. Create a brand new Google Doc - or update existing?
  // toLocaleDateString or toISOString
  var doc = DocumentApp.create('Starwood 2026 Presenters and Workshops - Draft ' + new Date().toISOString());
  var body = doc.getBody();

  

  // 4. Append the grouped data into the Doc
  for (var presenter in presenterWorkshops) {
    // Add the grouping value as a Heading 1 section
    var heading = body.appendParagraph(presenter);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    // Insert the bio paragraph
    body.appendParagraph(biosmap[presenter]);

    // Loop through and list all items belonging to this group
    presenterWorkshops[presenter].forEach(function(itemTitle) {
      var listitem = body.appendListItem('').setGlyphType(DocumentApp.GlyphType.BULLET);
      listitem.appendText(itemTitle).setBold(true);
      listitem.appendText(workshopDetail[itemTitle]).setBold(false);
      if(workshopPrimary[itemTitle] === presenter){
        if(workshopDescription[itemTitle]) listitem.appendText(workshopDescription[itemTitle]).setBold(false);
      }else{
        listitem.appendText("\nSee description under " + workshopPrimary[itemTitle]).setBold(false);
      }
    });
    
    // Add an empty space between sections
    body.appendParagraph("");
  }

  // 4.5 Now create the schedule
  var schedssheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auto schedule");
  var schedsdata = schedssheet.getDataRange().getValues();
  var schedsrows = schedsdata.slice(1);
  var schedsmap = {};
  var schedday = {};
  var scheddaymap = {};
  var scheddetails = {};
  schedsrows.forEach(function(row) {
    // Presenter name
    var schedinitial = row[0];
    var schedwho = row[1];
    var schedwhat = row[2];
    var schedwhere = row[3];
    var schedwhenday = row[4];
    var schedwhentime = row[5];
    if(!schedwhat) return;
    if(!schedday) return;
    if(schedwhenday === 'When-day') return;

    // Init the array of times per day, when needed
    if(!schedday[schedwhenday]){
      schedday[schedwhenday] = [];
      scheddaymap[schedwhenday] = {};
      scheddetails[schedwhenday] = {};
    }
    // Init the array of events per time
    if(!scheddaymap[schedwhenday][schedwhentime]){
      scheddaymap[schedwhenday][schedwhentime] = [];
      schedday[schedwhenday].push(schedwhentime);
      scheddetails[schedwhenday][schedwhentime] = [];
    }
    scheddaymap[schedwhenday][schedwhentime].push(schedwhat);

    // save details
    scheddetails[schedwhenday][schedwhentime][schedwhat] = [];
    scheddetails[schedwhenday][schedwhentime][schedwhat]["who"] = schedwho;
    scheddetails[schedwhenday][schedwhentime][schedwhat]["initial"] = schedinitial;
    scheddetails[schedwhenday][schedwhentime][schedwhat]["where"] = schedwhere;
  });

  for (var daylist in schedday){
    var dayitem = body.appendParagraph(daylist);
    dayitem.setHeading(DocumentApp.ParagraphHeading.HEADING1);

    schedday[daylist].forEach(function(daytime) {
      //var listitem = body.appendListItem('').setGlyphType(DocumentApp.GlyphType.BULLET);
      //listitem.appendText(daytime);
      var timeitem = body.appendParagraph(daytime + "\n").setBold(true);
      scheddaymap[daylist][daytime].forEach(function(eventitem) {
        timeitem.appendText(eventitem + " - " + scheddetails[daylist][daytime][eventitem]["initial"] + " - " + scheddetails[daylist][daytime][eventitem]["where"] + "\n").setBold(false);
      });
    });
    //schedday.forEach(function(eventitem){
     // var listitem = body.appendListItem('').setGlyphType(DocumentApp.GlyphType.BULLET);
     // listitem.appendText(eventitem);
    //});
  }

  // 5. Notify the user with a link to the new file
  var url = doc.getUrl();
  SpreadsheetApp.getUi().alert('Success! Your grouped document has been created. Open it here:\n\n' + url);
}
