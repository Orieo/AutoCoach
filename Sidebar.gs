function onInstall(e) {
  onOpen();
}

function onOpen() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('Run manually', 'run')
  .addItem('Automated emails settings', 'slotMenu')
  .addItem('Engagement email settings', 'engageMenu')
  .addItem('Team Communication settings', 'teamMenu')
  .addItem('Manually Add Client', 'addUser')
  .addItem('CSV Import Clients', 'addBatchCSV')
  .addItem('Run Archive Manually', 'manualArchive')
  .addItem('Enter Authorization Key', 'enterAuthKey')
  .addItem("Reset (Admin Only)", 'reset')
  .addToUi();
}


function reset() {
  var response = SpreadsheetApp.getUi().alert('WARNING: Reset Properties', 'Resetting will erase all your configurations for your emails and more. Do not\
reset unless instructed to by an admin. Do you wish to continue? ', SpreadsheetApp.getUi().ButtonSet.YES_NO);
  
  if(response == SpreadsheetApp.getUi().Button.YES){
    installProperties();
  }
}

function addUser(){
  var response = Browser.inputBox('Manually Add Client', "Enter the client's first name, last name, and email seperated by commas and no spaces (John,Smith,john.smith@email.com)",
                                               SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  
  var dateObj = new Date();
  var dateSplit = dateObj.toString().split(" ");
  var date = dateSplit[1] + " " + dateSplit[2] + " " + dateSplit[3] + " " + dateSplit[4];
  
  var info = response.split(",");
  var formula = '=HYPERLINK("https://mail.google.com/mail/u/0/#search/'+info[2]+'","'+info[0]+'")';
  var emailForm = info[2];
  var lname = info[1];
  var emailData = ["",formula,lname,emailForm, "Yes", "", "", "SEND", "","","","","","","", date, "", ""];
  
  var emailSheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email tracker");
  emailSheet.appendRow(emailData);
  
  emailSheet.sort(16, false);
  
}

function addBatchCSV() {
  var response = Browser.inputBox('Batch Add Clients', "Copy and Paste CSV file into input box below",SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  
  var dateObj = new Date();
  var dateSplit = dateObj.toString().split(" ");
  var date = dateSplit[1] + " " + dateSplit[2] + " " + dateSplit[3] + " " + dateSplit[4];
  
  var csvData = response.replace(/\s/g,'').split(",");
  
  for(var i = 0; i < csvData.length - 1; i+=5) {
      var formula = '=HYPERLINK("https://mail.google.com/mail/u/0/#search/'+csvData[i+2]+'","'+csvData[i]+'")';
      var lname = csvData[i+1];
      var emailForm = csvData[i+2];
      var pass = csvData[i+3];
      var phone = csvData[i+4];
    
      var emailData = ["",formula,lname,emailForm, "Yes", "", "", "SEND", "","","","","","","", date, pass, phone];
  
      var emailSheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email tracker");
      emailSheet.appendRow(emailData);
  }
  
  emailSheet.sort(16, false);
  
  
}


function installProperties(){
  PropertiesService.getDocumentProperties().setProperty('config', '{"sendTo":"","startTime":"","endTime":"","testEmail":"","timer":"Off","logId":"", "archive":"No"}');
  PropertiesService.getDocumentProperties().setProperty('slot1', '{"trigger":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('slot2', '{"trigger":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('slot3', '{"trigger":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('slot4', '{"trigger":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('slot5', '{"trigger":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('slot6', '{"trigger":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('slot7', '{"trigger":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('slot8', '{"trigger":"","subject":"","log":""}');

  PropertiesService.getDocumentProperties().setProperty('poke1', '{"code":"Poke 1","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke2', '{"code":"Poke 2","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke3', '{"code":"Poke 3","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke4', '{"code":"Poke 4","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke5', '{"code":"Poke 5","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke6', '{"code":"Poke 6","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke7', '{"code":"Poke 7","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke8', '{"code":"Poke 8","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke9', '{"code":"Poke 9","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke10', '{"code":"Poke 10","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke11', '{"code":"Poke 11","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke12', '{"code":"Poke 12","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke13', '{"code":"Poke 13","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke14', '{"code":"Poke 14","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke15', '{"code":"Poke 15","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke16', '{"code":"Poke 16","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke17', '{"code":"Poke 17","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke18', '{"code":"Poke 18","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke19', '{"code":"Poke 19","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke20', '{"code":"Poke 20","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke21', '{"code":"Poke 21","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke22', '{"code":"Poke 22","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke23', '{"code":"Poke 23","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke24', '{"code":"Poke 24","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke25', '{"code":"Poke 25","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke26', '{"code":"Poke 26","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke27', '{"code":"Poke 27","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke28', '{"code":"Poke 28","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke29', '{"code":"Poke 29","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke30', '{"code":"Poke 30","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke31', '{"code":"Poke 31","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke32', '{"code":"Poke 32","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke33', '{"code":"Poke 33","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke34', '{"code":"Poke 34","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('poke35', '{"code":"Poke 35","subject":"","log":""}');
  
  PropertiesService.getDocumentProperties().setProperty('team1', '{"code":"Team 1","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team2', '{"code":"Team 2","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team3', '{"code":"Team 3","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team4', '{"code":"Team 4","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team5', '{"code":"Team 5","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team6', '{"code":"Team 6","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team7', '{"code":"Team 7","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team8', '{"code":"Team 8","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team9', '{"code":"Team 9","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team10', '{"code":"Team 10","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team11', '{"code":"Team 11","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team12', '{"code":"Team 12","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team13', '{"code":"Team 13","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team14', '{"code":"Team 14","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team15', '{"code":"Team 15","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team16', '{"code":"Team 16","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team17', '{"code":"Team 17","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team18', '{"code":"Team 18","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team19', '{"code":"Team 19","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team20', '{"code":"Team 20","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team21', '{"code":"Team 21","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team22', '{"code":"Team 22","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team23', '{"code":"Team 23","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team24', '{"code":"Team 24","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team25', '{"code":"Team 25","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team26', '{"code":"Team 26","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team27', '{"code":"Team 27","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team28', '{"code":"Team 28","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team29', '{"code":"Team 29","to":"","subject":"","log":""}');
  PropertiesService.getDocumentProperties().setProperty('team30', '{"code":"Team 30","to":"","subject":"","log":""}');
}

function slotMenu() {
  loadPage("public/emailMenu");
}

function engageMenu() {
  loadPage("public/poke");
}

function teamMenu() { 
  loadPage("public/teamComm");
}

function getLabels() {
  var htmlLabels = "";
  var labels = GmailApp.getUserLabels();
  for(var i = 0; i < labels.length; i++) {
    htmlLabels += ("<option value='" + labels[i].getName().replace(/'/, "") + "'>" + labels[i].getName() + "</option>");
  }
  return htmlLabels;
}

function getTemplates(label) {
  var templates = "";
  var messages = GmailApp.getMessagesForThreads(GmailApp.getUserLabelByName(label).getThreads());
  for(var i = 0; i < messages.length; i++) {
    for(var j = 0; j < messages[i].length; j++) {
      if(messages[i][j].getSubject().substring(0,4) != "Re:" || messages[i][j].getSubject().substring(0,5) != "Fwd:"){
        templates += ("<option value='" + messages[i][j].getSubject().replace(/'/, "") + "'>" + messages[i][j].getSubject().replace(/'/, "") + "</option>");
      }
    }
  }
  return templates;
}

function setSlotProperty(slot, string) {
 PropertiesService.getDocumentProperties().setProperty(slot, string); 
}

function setTimer(state) {
  if(state == "On" && ScriptApp.getUserTriggers(SpreadsheetApp.getActive())[0] == undefined) {
    trigger = ScriptApp.newTrigger("run").timeBased().everyHours(1).create();
  }else if(state == "Off") {
    ScriptApp.deleteTrigger(ScriptApp.getUserTriggers(SpreadsheetApp.getActive())[0]);
    
  }
}

function enterAuthKey() {
  var key = Browser.inputBox('Enter Authorization Key', 'Enter the key sent to your email to continue using AutoCoach.', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  PropertiesService.getDocumentProperties().setProperty('authKey', key);
  authorize();  
}

function authorize() {
  var authKey = PropertiesService.getDocumentProperties().getProperty('authKey');
  var email = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
  var date = new Date();
  
  var thisMonth = date.getMonth() + 1;
  
  var lastMonth = date.getMonth();
  if(lastMonth == 0){
    lastMonth = 12;
  }
  
  var letterCode = email.charCodeAt(0);
  
  var algorithm = 98888888 -(email.length * thisMonth * date.getFullYear()) - letterCode - 7;
  var lastAlg = 98888888 -(email.length * lastMonth * date.getFullYear()) - letterCode;
  
  if(authKey == "moad9165"){
    Logger.log("Master");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Tracker").getRange(1, 1).setValue("Master").setBackground("purple");
    return "4";
  }else if(Math.sqrt(Math.sqrt(authKey % algorithm)) == 6){
    Logger.log("Premium+");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Tracker").getRange(1, 1).setValue("Premium+").setBackground("blue");
    return "3";
  }else if(Math.sqrt(Math.sqrt(authKey % lastAlg)) == 6 && date.getDate() <= 7){
    Logger.log("Premium+");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Tracker").getRange(1, 1).setValue("Premium+").setBackground("blue");
    return "3";
  }else if(Math.sqrt(Math.sqrt(authKey % algorithm)) == 4){
    Logger.log("Premium");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Tracker").getRange(1, 1).setValue("Premium").setBackground("green");
    return "2";
  }else if(Math.sqrt(Math.sqrt(authKey % lastAlg)) == 4 && date.getDate() <= 7){
    Logger.log("Premium");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Tracker").getRange(1, 1).setValue("Premium").setBackground("green");
    return "2";
  }else if(Math.sqrt(Math.sqrt(authKey % algorithm)) == 2){
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Tracker").getRange(1, 1).setValue("Lite").setBackground("orange");
    Logger.log('Lite');
    return "1";
  }else if(Math.sqrt(Math.sqrt(authKey % lastAlg)) == 2 && date.getDate() <= 7){
    Logger.log('Lite');
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Tracker").getRange(1, 1).setValue("Lite").setBackground("orange");
    return "1";
  }else{
    Logger.log('false');
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Tracker").getRange(1, 1).setValue("Unauthorized").setBackground("red");
    return "0"; 
  }
  
}

function manualArchive(){
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  emailSheet = spreadsheet.getSheetByName("Email Tracker");
  doneSheet = spreadsheet.getSheetByName("Completed Clients");
  inactiveSheet = spreadsheet.getSheetByName("Inactive Clients");
  range = parseInt(emailSheet.getSheetValues(1, 19, 1, 1), "10") + 1;
  emailValues = emailSheet.getSheetValues(2, 1, range, 18);
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
      var tmp = emailSheet.getRange(i + 2, 1, 1, 18).getValues()[0];
      inactiveSheet.appendRow([tmp[1], tmp[2], tmp[3], tmp[15], tmp[17], tmp[16]]);
      emailSheet.deleteRow(i + 2);
      counter++;
    }
    
    if(counter >= 20){
      break;
    }
  }
}

function updateDropdown() {
  var pokeEmails;
  var data = [];
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(2, 6, sheet.getMaxRows(), 1);
  var rules = range.getDataValidations();
  
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
  
  pokeEmails = [poke1Obj, poke2Obj, poke3Obj, poke4Obj, poke5Obj, poke6Obj, poke7Obj, poke8Obj, poke9Obj, poke10Obj,
               poke11Obj, poke12Obj, poke13Obj, poke14Obj, poke15Obj, poke16Obj, poke17Obj, poke18Obj, poke19Obj, poke20Obj,
               poke21Obj, poke22Obj, poke23Obj, poke24Obj, poke25Obj, poke26Obj, poke27Obj, poke28Obj, poke29Obj, poke30Obj,
               poke31Obj, poke32Obj, poke33Obj, poke34Obj, poke35Obj];

  for(var i = 0; i < 35; i++) {
    if(pokeEmails[i].code.substring(0, 4) != "Poke") {
      data.push(pokeEmails[i].code);
    }
  }
  
  var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInList(data, true).build();
  
  for(var i = 0; i < rules.length; i++) {
    rules[i][0] = rule;
  }
  
  range.setDataValidations(rules);
}

function updateDropdownTeam() {
  var teamEmails;
  var data = [];
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(2, 7, sheet.getMaxRows(), 1);
  var rules = range.getDataValidations();
  
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
  
  teamEmails = [team1Obj, team2Obj, team3Obj, team4Obj, team5Obj, team6Obj, team7Obj, team8Obj, team9Obj, team10Obj,
               team11Obj, team12Obj, team13Obj, team14Obj, team15Obj, team16Obj, team17Obj, team18Obj, team19Obj, team20Obj,
               team21Obj, team22Obj, team23Obj, team24Obj, team25Obj, team26Obj, team27Obj, team28Obj, team29Obj, team30Obj];

  for(var i = 0; i < 30; i++) {
    if(teamEmails[i].code.substring(0, 4) != "Team") {
      data.push(teamEmails[i].code);
    }
  }
  
  var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInList(data, true).build();
  
  for(var i = 0; i < rules.length; i++) {
    rules[i][0] = rule;
  }
  
  range.setDataValidations(rules);
}