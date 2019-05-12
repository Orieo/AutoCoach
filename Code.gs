var spreadsheet;
var doneSheet;
var inactiveSheet;
var emailSheet;
var log;
var doneUsers;
var inactiveUsers;
var range;
var doneRange;
var inactiveRange;

var slot1Obj;
var slot2Obj;
var slot3Obj;
var slot4Obj;
var slot5Obj;
var slot6Obj;
var slot7Obj;
var slot8Obj;

var poke1Obj;
var poke2Obj;
var poke3Obj;
var poke4Obj;
var poke5Obj;
var poke6Obj;
var poke7Obj;
var poke8Obj;
var poke9Obj;
var poke10Obj;
var poke11Obj;
var poke12Obj;
var poke13Obj;
var poke14Obj;
var poke15Obj;
var poke16Obj;
var poke17Obj;
var poke18Obj;
var poke19Obj;
var poke20Obj;
var poke21Obj;
var poke22Obj;
var poke23Obj;
var poke24Obj;
var poke25Obj;
var poke26Obj;
var poke27Obj;
var poke28Obj;
var poke29Obj;
var poke30Obj;
var poke31Obj;
var poke32Obj;
var poke33Obj;
var poke34Obj;
var poke35Obj;

var team1Obj;
var team2Obj;
var team3Obj;
var team4Obj;
var team5Obj;
var team6Obj;
var team7Obj;
var team8Obj;
var team9Obj;
var team10Obj;
var team11Obj;
var team12Obj;
var team13Obj;
var team14Obj;
var team15Obj;
var team16Obj;
var team17Obj;
var team18Obj;
var team19Obj;
var team20Obj;
var team21Obj;
var team22Obj;
var team23Obj;
var team24Obj;
var team25Obj;
var team26Obj;
var team27Obj;
var team28Obj;
var team29Obj;
var team30Obj;

var config;
var subType;

var dateObj;
var dateSplit;
var date;
var curDate;
var engDate;


/**
* Initializes spreadsheet and property variables
*/
function initGlobalVariables() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  doneSheet = spreadsheet.getSheetByName("Completed Clients");
  inactiveSheet = spreadsheet.getSheetByName("Inactive Clients");
  emailSheet = spreadsheet.getSheetByName("Email Tracker");
  range = parseInt(emailSheet.getSheetValues(1, 19, 1, 1), "10") + 1;
  doneRange = parseInt(doneSheet.getSheetValues(1, 7, 1, 1), "10") + 1;
  inactiveRange = parseInt(inactiveSheet.getSheetValues(1, 7, 1, 1), "10") + 1;
  emailValues = emailSheet.getSheetValues(2, 1, range, 17);
  doneUsers = doneSheet.getSheetValues(2, 1, doneRange, 3);
  inactiveUsers = inactiveSheet.getSheetValues(2, 1, inactiveRange, 3);

  
  dateObj = new Date();
  dateSplit = dateObj.toString().split(" ");
  date = dateSplit[1] + " " + dateSplit[2] + " " + dateSplit[3] + " " + dateSplit[4];
  curDate = dateSplit[1] + "/" + dateSplit[2];
  engDate = (dateObj.getMonth()+ 1)+"/"+dateObj.getDate()+"/"+dateObj.getYear()+" "+dateObj.getHours()+":"+dateObj.getMinutes()+":"+dateObj.getSeconds();

  subType = authorize();
  
  slot1Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("slot1"));
  slot2Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("slot2"));
  slot3Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("slot3"));
  slot4Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("slot4"));
  slot5Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("slot5"));
  slot6Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("slot6"));
  slot7Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("slot7"));
  slot8Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("slot8"));
  
  poke1Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke1"));
  poke2Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke2"));
  poke3Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke3"));
  poke4Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke4"));
  poke5Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke5"));
  poke6Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke6"));
  poke7Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke7"));
  poke8Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke8"));
  poke9Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke9"));
  poke10Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke10"));
  poke11Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke11"));
  poke12Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke12"));
  poke13Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke13"));
  poke14Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke14"));
  poke15Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke15"));
  
  poke16Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke16"));
  poke17Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke17"));
  poke18Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke18"));
  poke19Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke19"));
  poke20Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke20"));
  poke21Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke21"));
  poke22Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke22"));
  poke23Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke23"));
  poke24Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke24"));
  poke25Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke25"));
  poke26Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke26"));
  
  poke27Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke27"));
  poke28Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke28"));
  poke29Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke29"));
  poke30Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke30"));
  poke31Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke31"));
  poke32Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke32"));
  poke33Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke33"));
  poke34Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke34"));
  poke35Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("poke35"));
  
  team1Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team1"));
  team2Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team2"));
  team3Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team3"));
  team4Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team4"));
  team5Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team5"));
  team6Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team6"));
  team7Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team7"));
  team8Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team8"));
  team9Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team9"));
  team10Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team10"));
  team11Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team11"));
  team12Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team12"));
  team13Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team13"));
  team14Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team14"));
  team15Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team15"));
  
  team16Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team16"));
  team17Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team17"));
  team18Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team18"));
  team19Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team19"));
  team20Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team20"));
  team21Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team21"));
  team22Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team22"));
  team23Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team23"));
  team24Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team24"));
  team25Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team25"));
  team26Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team26"));
  
  team27Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team27"));
  team28Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team28"));
  team29Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team29"));
  team30Obj = JSON.parse(PropertiesService.getDocumentProperties().getProperty("team30"));
    
  config = JSON.parse(PropertiesService.getDocumentProperties().getProperty("config"));
  
  if(config.logId == ''){
    log = spreadsheet.getSheetByName("Log")
  }else{
    Logger.log(config.logId);
    log = SpreadsheetApp.openById(config.logId);
  }
}


/**
* Sends Emails and parses inbox for labeled emails
*/
function run() {
  if(authorize() >= 1) {
    initGlobalVariables();
   
    if(dateObj.getHours() >= 0 && dateObj.getHours() <= 1){
      archive(); 
    }else{
      if(((dateObj.getHours() >= config.startTime && dateObj.getHours() < config.endTime)||(config.startTime=="" && config.endTime== "")) 
      && config.sendTo != 'No'){
        emailValues.forEach(sendEmails);
      }
  
      if(slot8Obj.trigger != '' && slot8Obj.trigger != null){
        var slot8users = getUserDataFromLabel(slot8Obj.trigger);
        handleUsersForSlot(slot8users, 8);
      }
      if(slot7Obj.trigger != '' && slot7Obj.trigger != null){
        var slot7users = getUserDataFromLabel(slot7Obj.trigger);
        handleUsersForSlot(slot7users, 7);
      }
      if(slot6Obj.trigger != '' && slot6Obj.trigger != null){
        var slot6users = getUserDataFromLabel(slot6Obj.trigger);
        handleUsersForSlot(slot6users, 6);
      }
      if(slot5Obj.trigger != '' && slot5Obj.trigger != null){
        var slot5users = getUserDataFromLabel(slot5Obj.trigger);
        handleUsersForSlot(slot5users, 5);
      }
      if(slot4Obj.trigger != '' && slot4Obj.trigger != null){
        var slot4users = getUserDataFromLabel(slot4Obj.trigger);
        handleUsersForSlot(slot4users, 4);
      }
      if(slot3Obj.trigger != '' && slot3Obj.trigger != null){
        var slot3users = getUserDataFromLabel(slot3Obj.trigger);
        handleUsersForSlot(slot3users, 3);
      }
      if(slot2Obj.trigger != '' && slot2Obj.trigger != null){
        var slot2users = getUserDataFromLabel(slot2Obj.trigger);
        handleUsersForSlot(slot2users, 2);
      }
      if(slot1Obj.trigger != '' && slot1Obj.trigger != null){
        var users = getUserDataFromLabel(slot1Obj.trigger);
        users.forEach(handleNewUser);
      }
    }
    
    sort();  
  }
}


/**
* Checks if user is already on sheet, then adds them if they aren't
* @param {String[]} user
*/
function handleNewUser(user) {
  var addLead = true;
  for(var j = 0; j < range; j++){
    if((user.firstname == emailValues[j][1] && user.lastname == emailValues[j][2]) || user.email == emailValues[j][3]) {
      addLead = false;
      break;
    }
  }
  if(addLead){
    for(var j = 0; j < doneRange; j++){
      if((user.firstname == doneUsers[j][0] && user.lastname == doneUsers[j][1]) || user.email == doneUsers[j][2]) {
        addLead = false;
        break;
      }
    }
  }
  
  if(addLead){
    for(var j = 0; j < inactiveRange; j++){
      if((user.firstname == inactiveUsers[j][0] && user.lastname == inactiveUsers[j][1]) || user.email == inactiveUsers[j][2]) {
        addLead = false;
        break;
      }
    }
  }

  if(addLead) {
    var emailData = ["", '=HYPERLINK("https://mail.google.com/mail/u/0/#search/'+user.email+'","'+user.firstname+'")',
                     user.lastname,user.email,
                     "Yes", "", "", "SEND", "","","","","","","", user.dateAdded, user.password, user.phone];
    emailSheet.appendRow(emailData);
  }
}

function handleUsersForSlot(users, slot){
  var len = users.length;
  
  for(var i = 0; i < len; i++){
    var user = users[i];
    for(var j = 0; j < range; j++){
      if(user.email == emailValues[j][3]) {
        if(emailValues[j][6 + slot] == ""){
          emailSheet.getRange(j + 2, 7 + slot).setValue("SEND");
        }
        break;
      }
    }
  }
}

function sendEmails(user, index) {
  var varData;
  var version = PropertiesService.getDocumentProperties().getProperty("version");
    
  if(config.sendTo == "Client") {
    varData = {
      firstname: emailValues[index][1],
      lastname: emailValues[index][2],
      email: emailValues[index][3],
      dateAdded: emailValues[index][15],
      phone: emailValues[index][17],
      password: emailValues[index][16],
    };
  }else if(config.sendTo == "Test") {
    varData = {
      firstname: emailValues[index][1],
      lastname: emailValues[index][2],
      email: config.testEmail,
      dateAdded: emailValues[index][15],
      phone: emailValues[index][17],
      password: emailValues[index][16],
    };
  }
  
  sendSlotFor(1, user, index, varData, slot1Obj, subType);
  sendSlotFor(2, user, index, varData, slot2Obj, subType);
  sendSlotFor(3, user, index, varData, slot3Obj, subType);
  sendSlotFor(4, user, index, varData, slot4Obj, subType);
  sendSlotFor(5, user, index, varData, slot5Obj, subType);
  sendSlotFor(6, user, index, varData, slot6Obj, subType);
  sendSlotFor(7, user, index, varData, slot7Obj, subType);
  sendSlotFor(8, user, index, varData, slot8Obj, subType);
  
  sendPokeFor(user, index, varData, poke1Obj, subType);
  sendPokeFor(user, index, varData, poke2Obj, subType);
  sendPokeFor(user, index, varData, poke3Obj, subType);
  sendPokeFor(user, index, varData, poke4Obj, subType);
  sendPokeFor(user, index, varData, poke5Obj, subType);
  sendPokeFor(user, index, varData, poke6Obj, subType);
  sendPokeFor(user, index, varData, poke7Obj, subType);
  sendPokeFor(user, index, varData, poke8Obj, subType);
  sendPokeFor(user, index, varData, poke9Obj, subType);
  sendPokeFor(user, index, varData, poke10Obj, subType);
  sendPokeFor(user, index, varData, poke11Obj, subType);
  sendPokeFor(user, index, varData, poke12Obj, subType);
  sendPokeFor(user, index, varData, poke13Obj, subType);
  sendPokeFor(user, index, varData, poke14Obj, subType);
  sendPokeFor(user, index, varData, poke15Obj, subType);
  if(subType >= 2) {
    sendPokeFor(user, index, varData, poke16Obj, subType);
    sendPokeFor(user, index, varData, poke17Obj, subType);
    sendPokeFor(user, index, varData, poke18Obj, subType);
    sendPokeFor(user, index, varData, poke19Obj, subType);
    sendPokeFor(user, index, varData, poke20Obj, subType);
    sendPokeFor(user, index, varData, poke21Obj, subType);
    sendPokeFor(user, index, varData, poke22Obj, subType);
    sendPokeFor(user, index, varData, poke23Obj, subType);
    sendPokeFor(user, index, varData, poke24Obj, subType);
    sendPokeFor(user, index, varData, poke25Obj, subType);
  }
  if(subType >= 3) {
    sendPokeFor(user, index, varData, poke26Obj, subType);
    sendPokeFor(user, index, varData, poke27Obj, subType);
    sendPokeFor(user, index, varData, poke28Obj, subType);
    sendPokeFor(user, index, varData, poke29Obj, subType);
    sendPokeFor(user, index, varData, poke30Obj, subType);
    sendPokeFor(user, index, varData, poke31Obj, subType);
    sendPokeFor(user, index, varData, poke32Obj, subType);
    sendPokeFor(user, index, varData, poke33Obj, subType);
    sendPokeFor(user, index, varData, poke34Obj, subType);
    sendPokeFor(user, index, varData, poke35Obj, subType);
    
    
    sendTeamFor(user, index, varData, team1Obj, subType);
    sendTeamFor(user, index, varData, team2Obj, subType);
    sendTeamFor(user, index, varData, team3Obj, subType);
    sendTeamFor(user, index, varData, team4Obj, subType);
    sendTeamFor(user, index, varData, team5Obj, subType);
    sendTeamFor(user, index, varData, team6Obj, subType);
    sendTeamFor(user, index, varData, team7Obj, subType);
    sendTeamFor(user, index, varData, team8Obj, subType);
    sendTeamFor(user, index, varData, team9Obj, subType);
    sendTeamFor(user, index, varData, team10Obj, subType);
    sendTeamFor(user, index, varData, team11Obj, subType);
    sendTeamFor(user, index, varData, team12Obj, subType);
    sendTeamFor(user, index, varData, team13Obj, subType);
    sendTeamFor(user, index, varData, team14Obj, subType);
    sendTeamFor(user, index, varData, team15Obj, subType);
    sendTeamFor(user, index, varData, team16Obj, subType);
    sendTeamFor(user, index, varData, team17Obj, subType);
    sendTeamFor(user, index, varData, team18Obj, subType);
    sendTeamFor(user, index, varData, team19Obj, subType);
    sendTeamFor(user, index, varData, team20Obj, subType);
    sendTeamFor(user, index, varData, team21Obj, subType);
    sendTeamFor(user, index, varData, team22Obj, subType);
    sendTeamFor(user, index, varData, team23Obj, subType);
    sendTeamFor(user, index, varData, team24Obj, subType);
    sendTeamFor(user, index, varData, team25Obj, subType);
    sendTeamFor(user, index, varData, team26Obj, subType);
    sendTeamFor(user, index, varData, team27Obj, subType);
    sendTeamFor(user, index, varData, team28Obj, subType);
    sendTeamFor(user, index, varData, team29Obj, subType);
    sendTeamFor(user, index, varData, team30Obj, subType);
  } 
}

function sendSlotFor(slot, user, index, varData, slotObj, subType) {
  if(user[slot + 6] == "SEND" && user[4] == "Yes") {
    try{
      var id;
      var templates = "";
      var messages = GmailApp.getMessagesForThreads(GmailApp.getUserLabelByName("AC Step Emails").getThreads());
      
      for(var i = 0; i < messages.length; i++) {
        for(var j = 0; j < messages[i].length; j++) {
          if(slotObj.subject == messages[i][j].getSubject().replace(/'/, "")) {
            id = messages[i][j].getId();
            break;
          }
        }
      }
      var canned = GmailApp.getMessageById(id);
      var options = {htmlBody: fillInTemplateFromObject(canned.getBody(), varData)};
      if(config.sendTo == "Client"){
        options.cc = canned.getCc();
      } 
      if(config.sendTo == "Client" && subType >= 3){
        options.bcc = canned.getBcc();
      }
      
      options.attachments = canned.getAttachments();
      
      if(Session.getActiveUser().getEmail() == 'al@mydotcomeducation.net'){
        options.name = 'Al Monteiro'; 
      }
    
      GmailApp.sendEmail(varData.email, canned.getSubject(), "", options);
      emailSheet.getRange(index + 2, slot + 7).setValue(curDate).setBackgroundRGB(183,225,205);
      var logData = [date, '=HYPERLINK("https://mail.google.com/mail/u/0/#search/'+varData.email+'","'+varData.firstname+'")', varData.lastname,
                     '=HYPERLINK("https://coaches.mobe.com/contacts?contact='+varData.email+'","'+varData.email+'")', slotObj.log];
    }catch(err){
      var slotName = emailSheet.getRange(1, slot + 7).getValue();
      emailSheet.getRange(index + 2, slot + 7).setValue("Error").setBackgroundRGB(183,225,205);
      var logData = [date,'=HYPERLINK("https://mail.google.com/mail/u/0/#search/'+varData.email+'","'+varData.firstname+'")', varData.lastname,
                     '=HYPERLINK("https://coaches.mobe.com/contacts?contact='+varData.email+'","'+varData.email+'")', "Error sending email (" + slotName + ")\n" + err]; 
    }
    if(subType >= 2) {
      log.appendRow(logData);
    }
  }
}

function sendPokeFor(user, index, varData, emailObj, subType) {
  
  if(user[5] == emailObj.code){
    try{
    
      var id;
      var templates = "";
      var messages = GmailApp.getMessagesForThreads(GmailApp.getUserLabelByName("AC Engmt Emails").getThreads());
      
      for(var i = 0; i < messages.length; i++) {
        for(var j = 0; j < messages[i].length; j++) {
          if(emailObj.subject == messages[i][j].getSubject().replace(/'/, "")) {
            id = messages[i][j].getId();
            break;
          }
        }
      }
    
      var canned = GmailApp.getMessageById(id);
      var options = {htmlBody: fillInTemplateFromObject(canned.getBody(), varData)};
      
      if(config.sendTo == "Client"){
        options.cc = canned.getCc();
      } 
      if(config.sendTo == "Client" && subType >= 3){
        options.bcc = canned.getBcc();
      }
      
      options.attachments = canned.getAttachments();
      
      if(Session.getActiveUser().getEmail() == 'al@mydotcomeducation.net'){
        options.name = 'Al Monteiro'; 
      }
    
      GmailApp.sendEmail(varData.email, canned.getSubject(), "", options);
      emailSheet.getRange(index + 2, 6).setValue(emailObj.code + " (SENT)");
    
      var val = emailSheet.getRange(index + 2, 4).getNote();
      emailSheet.getRange(index + 2, 4).setNote(val + engDate + " " +  emailObj.code + "\n");
    
      var logData = [date,'=HYPERLINK("https://mail.google.com/mail/u/0/#search/'+varData.email+'","'+varData.firstname+'")', varData.lastname,
                     '=HYPERLINK("https://coaches.mobe.com/contacts?contact='+varData.email+'","'+varData.email+'")', emailObj.log];
    }catch(err){
      emailSheet.getRange(index + 2, 6).setValue("(Error) " + emailObj.code);
      var logData = [date,'=HYPERLINK("https://mail.google.com/mail/u/0/#search/'+varData.email+'","'+varData.firstname+'")', varData.lastname,
                     '=HYPERLINK("https://coaches.mobe.com/contacts?contact='+varData.email+'","'+varData.email+'")', "Error sending poke email (" + emailObj.code + ")\n" + err];
    }
    
    if(subType >= 2) {
      log.appendRow(logData);
    }
  }  
}

function sendTeamFor(user, index, varData, emailObj, subType) {
  
  if(user[6] == emailObj.code){
    try{
    
      var id;
      var templates = "";
      var messages = GmailApp.getMessagesForThreads(GmailApp.getUserLabelByName("AC Team Emails").getThreads());
      
      for(var i = 0; i < messages.length; i++) {
        for(var j = 0; j < messages[i].length; j++) {
          if(emailObj.subject == messages[i][j].getSubject().replace(/'/, "")) {
            id = messages[i][j].getId();
            break;
          }
        }
      }
      
      var canned = GmailApp.getMessageById(id);
      var options = {htmlBody: fillInTemplateFromObject(canned.getBody(), varData)};
      var to;
      
      to = canned.getTo();
      
      if(config.sendTo == "Client"){
        options.cc = canned.getCc();
      } 
      if(config.sendTo == "Client" && subType >= 3){
        options.bcc = canned.getBcc();
        
      }
      
      options.attachments = canned.getAttachments();
      
      
      if(Session.getActiveUser().getEmail() == 'alwyn@mobe.com'){
        options.name = 'Al Monteiro'; 
      }
    
      GmailApp.sendEmail(to, canned.getSubject(), "", options);
      emailSheet.getRange(index + 2, 7).setValue(emailObj.code + " (SENT)");
    
      var val = emailSheet.getRange(index + 2, 4).getNote();
      emailSheet.getRange(index + 2, 4).setNote(val + engDate + " " +  emailObj.code + "\n");
    
      var logData = [date,'=HYPERLINK("https://mail.google.com/mail/u/0/#search/'+varData.email+'","'+varData.firstname+'")', varData.lastname,
                     '=HYPERLINK("https://coaches.mobe.com/contacts?contact='+varData.email+'","'+varData.email+'")', emailObj.log];
    }catch(err){
      emailSheet.getRange(index + 2, 7).setValue("(Error) " + emailObj.code);
      var logData = [date,'=HYPERLINK("https://mail.google.com/mail/u/0/#search/'+varData.email+'","'+varData.firstname+'")', varData.lastname,
                     '=HYPERLINK("https://coaches.mobe.com/contacts?contact='+varData.email+'","'+varData.email+'")',
                     "Error sending team communication email (" + emailObj.code + ")\n" + err];
    }
    
    if(subType >= 2) {
      log.appendRow(logData);
    }
  }  
}



/**
* Pulls data from emails marked with a specified label
* @param {String} labeltext
* @return {Object[][]} users
*/
function getUserDataFromLabel(labeltext) {
  var users = []
  
  var label = GmailApp.getUserLabelByName(labeltext);
  var threads = label.getThreads(0, 2);

  
  for(var j = 0; j < threads.length; j++){
    var messages = threads[j].getMessages()
    var messageCount = messages.length
    for (var i = 0; i < messageCount; i++) {
      
      var tmp;
      var content = messages[i].getBody();
      
      content = content.replace(/(&nbsp;)/g," ");
      // Get the plain text body of the email message
      // Implement Parsing rules using regular expressions
      
      tmp = content.match(/Name:\s*([A-Za-z0-9\s.í]+)/g)[0].substring(6);
      Logger.log(tmp);
      if(tmp.split(" ")[0] != null){
        var firstname = tmp.split(" ")[0].capitalizeFirstLetter();
        firstname = firstname.trim();
      }else{
        var firstname = "" 
      }
      
      if(tmp.substring(firstname.length + 1) != null){
        var lastname = tmp.substring(firstname.length + 1);
        lastname = lastname.trim() 
      }else{
        var lastname = "" 
      }
      
      tmp = content.match(/Email:\s*([A-Za-z0-9_@.-]+)/g);
      if(tmp != null){
        var email = tmp[0].substring(7);
        if(email.search(/<w?br>?/) != -1){
          email = email.replace(/<w?br>?/g, "")
        }
      }else{
        var email = "Couldn't Retrieve Email" 
      }
      email = email.toString()
      
      tmp = messages[i].getDate().toString().split(" ")
      var dateAdded = tmp[1] + " " + tmp[2] + " " + tmp[3] + " " + tmp[4]
      
      tmp = content.match(/Phone:\s*[^\s]([0-9-\s*()]+)/)
      var phone
      if(tmp != null){
        phone = tmp[0];
        phone = phone.replace(/(Phone: )/, "");
      }else{
        phone = "";
      }
      
      
      tmp = content.match(/Password:\s*(&nbsp;)*([0-9A-Za-z]+)|MTTB Password:\s*(&nbsp;)*([0-9A-Za-z]+)/g)
      
      if(tmp != null) {
        var password = tmp.toString().trim().split(":")[1];
        password = password.replace(/&nbsp;/, "");
      }else{
        var password = null;
      }
      
      //Put All Data In One Object
      var userData = {
        firstname: firstname,
        lastname: lastname,
        email: email,
        dateAdded: dateAdded,
        phone: phone,
        password: password,
      };
      
      Logger.log(userData);
     
      users.push(userData);
    }
  }
  return users;
}

function archive(){
  var len = emailValues.length;
  var counter = 0;
  for(var i=len-1;i>=0;i--){
    var user = emailValues[i];
    if(user[0] == 'Archive: Completed' || user[0] == 'Arc:C') {
      var tmp = emailSheet.getRange(i + 2, 1, 1, 18).getValues()[0];
      doneSheet.appendRow([tmp[1], tmp[2], tmp[3], tmp[15], tmp[17], tmp[16]]);
      emailSheet.deleteRow(i + 2);
      counter++;
    }else if(user[0] == 'Archive: Inactive' || user[0] == 'Arc:I'){
      var tmp = emailSheet.getRange(i + 2, 1, 1, 17).getValues()[0];
      inactiveSheet.appendRow([tmp[1], tmp[2], tmp[3], tmp[15], tmp[17], tmp[16]]);
      emailSheet.deleteRow(i + 2);
      counter++;
    }
    
    if(counter >= 20){
      break;
    }
  }
}
