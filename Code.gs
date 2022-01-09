function sendEmail() {
  var mailApp = GmailApp;

  var spreadSheetApp = SpreadsheetApp;
  var spreadSheet = spreadSheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getActiveSheet();
  Logger.log(sheet.getRange(sheet.getLastRow(),1,1,sheet.getLastColumn()).getValues());

  // to_string()
  sheet.getRange(sheet.getLastRow(),12).setFormula('=TO_TEXT(E'+ sheet.getLastRow() +')'); // contact
  sheet.getRange(sheet.getLastRow(),13).setFormula('=TO_TEXT(G'+ sheet.getLastRow() +')'); // date
  sheet.getRange(sheet.getLastRow(),14).setFormula('=TO_TEXT(H'+ sheet.getLastRow() +')'); // t1
  sheet.getRange(sheet.getLastRow(),15).setFormula('=TO_TEXT(I'+ sheet.getLastRow() +')'); // t2

  var matching_sample = "\tName : {m_name}\n\tID : {m_id}\n\tEmail : {m_email}\n\tContact : {m_contact}\n\tStart Time : Between {m_start_t1} and {m_start_t2}\n\tComments : {m_comments}\n\n";

  var matching = "", matching_fail = "";

  var lastRowValues = sheet.getRange(sheet.getLastRow(),1,1,sheet.getLastColumn()).getValues();
  var name = lastRowValues[0][2], 
      id = lastRowValues[0][3], 
      email = lastRowValues[0][1], 
      contact = lastRowValues[0][4], 
      starting_location = lastRowValues[0][5], 
      ending_location = lastRowValues[0][9], 
      start_date = lastRowValues[0][12], 
      start_t1 = lastRowValues[0][13], 
      start_t2 = lastRowValues[0][14], 
      comments = lastRowValues[0][10];

  for(var itr = 2; itr <= sheet.getLastRow(); itr++) {
    let rowValues = sheet.getRange(itr,1,1,sheet.getLastColumn()).getValues();
    let m_name = rowValues[0][2], 
        m_id = rowValues[0][3], 
        m_email = rowValues[0][1],
        m_contact = rowValues[0][4], 
        m_starting_location = rowValues[0][5], 
        m_ending_location = rowValues[0][9], 
        m_start_date = rowValues[0][12], 
        m_start_t1 = rowValues[0][13], 
        m_start_t2 = rowValues[0][14], 
        m_comments = rowValues[0][10];
    
    if(m_email == email)
      continue;

    if(start_date != m_start_date)
      continue;

    if(starting_location != m_starting_location)
      continue;
    
    if(ending_location != m_ending_location)
      continue;

    matching_fail += matching_sample.replace('{m_name}',m_name).replace('{m_start_t1}',m_start_t1).replace('{m_start_t2}',m_start_t2).replace('{m_contact}',m_contact).replace('{m_id}',m_id).replace('{m_email}',m_email).replace('{m_comments}',m_comments);

    let a = Date.parse(start_date + ' ' + start_t1), b = Date.parse(start_date + ' ' + start_t2);
    let c = Date.parse(m_start_date + ' ' + m_start_t1), d = Date.parse(m_start_date + ' ' + m_start_t2);

    let op1 = ((a <= c) && (c <= b)), op2 = ((a <= d) && (d <= b)), op3 = ((c <= a) && (a <= d)), op4 = ((c <= b) && (b <= d));
    if(op1 || op2 || op3 || op4)
      matching += matching_sample.replace('{m_name}',m_name).replace('{m_start_t1}',m_start_t1).replace('{m_start_t2}',m_start_t2).replace('{m_contact}',m_contact).replace('{m_id}',m_id).replace('{m_email}',m_email).replace('{m_comments}',m_comments);
  }

  // date formatting for subject
  var d = start_date.split('/');
  const monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];

  mailApp.sendEmail(
    sheet.getRange(sheet.getLastRow(),1,1,sheet.getLastColumn()).getValues()[0][1], 
    "CabPool | " + d[1] + ' ' + monthNames[d[0]-1] + ' ' + d[2] + " | " + starting_location + " -> " + ending_location, 
    ("Hello {name},\n\nYour CabPool details :\n\tStart Date [MM/DD/YYYY] : {start_date}\n\tStart Time : Between {start_t1} and {start_t2}\n\tStarting Location : {starting_location}\n\tEnding Location : {ending_location}\n\tComments : {comments}\n\nCabPool matches :\n" + (matching == "" ? "\tNone\n\nShowing all cabpool(s) for {start_date} which are from {starting_location} to {ending_location} :\n".replace('{start_date}',start_date).replace('{starting_location}',starting_location).replace('{ending_location}',ending_location) + (matching_fail == "" ? "\tNone. Don\'t worry! Next person making an overlapping request will be notified about yours.\n" : matching_fail ) : matching) + "\nRegards,\nflux").replace('{name}',name).replace('{contact}',contact).replace('{start_date}',start_date).replace('{start_t1}',start_t1).replace('{start_t2}',start_t2).replace('{starting_location}',starting_location).replace('{ending_location}',ending_location).replace('{comments}',comments) );
}
