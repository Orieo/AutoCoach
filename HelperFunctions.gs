function loadPage(request) {
 var html = doGet(request).setTitle("AutoCoach");
 SpreadsheetApp.getUi().showSidebar(html);
}

function doGet(request) {
  return HtmlService.createTemplateFromFile(request).evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function insert(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent(); 
}

/**
* Capitalize first letter of string
*/
String.prototype.capitalizeFirstLetter = function() {
  return this.charAt(0).toUpperCase() + this.slice(1).toLowerCase()
}

/**
* Finds variable tags in a string and replaces them with data from an object
*/
function fillInTemplateFromObject(template, data) {
  var emailtext = template;
  // Search for all the variables to be replaced, for instance ${Column name}
  var templateVars = emailtext.match(/\$\{[^\"\s]+\}/g);
  
  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  if(templateVars.length > 0){
    for (var i = 0; i < templateVars.length; ++i) {
      // normalizeHeader ignores ${} so we can call it directly here.
      //Logger.log(normalizeHeader(templateVars[i]))
      var variableData = data[normalizeHeader(templateVars[i])];
      emailtext = emailtext.replace(templateVars[i], variableData || "");
    }
  }

  return emailtext;
}

/**
* Returns variable key from inside header tag
* @return {String} key
*/
function normalizeHeader(header) {
  var key = ""
  var upperCase = false
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i]
    if (letter == " " && key.length > 0) {
      upperCase = true
      continue
    }
    if (!isAlnum(letter)) {
      continue
    }
    if (key.length == 0 && isDigit(letter)) {
      continue // first character must be a letter
    }
    if (upperCase) {
      upperCase = false
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key
}

/**
* Returns true if character is alphanumeric.
* @return {Boolean}
*/
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

/**
* Returns true if the character char is a digit, false otherwise.
* @param {Char} char
* @return {Boolean}
*/
function isDigit(char) {
  return char >= '0' && char <= '9';
}

function sort() {
  emailSheet.sort(16, false);
}