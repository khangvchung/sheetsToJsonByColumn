// Includes functions for exporting active sheet or all sheets as JSON object (also Python object syntax compatible).
// Tweak the makePrettyJSON_ function to customize what kind of JSON to export.

var FORMAT_ONELINE   = 'One-line';
var FORMAT_MULTILINE = 'Multi-line';
var FORMAT_PRETTY    = 'Pretty';

var LANGUAGE_JS      = 'JavaScript';
var LANGUAGE_PYTHON  = 'Python';

var STRUCTURE_LIST = 'List';
var STRUCTURE_HASH = 'Hash (keyed by "id" column)';

/* Defaults for this particular spreadsheet, change as desired */
var DEFAULT_FORMAT = FORMAT_PRETTY;
var DEFAULT_LANGUAGE = LANGUAGE_JS;
var DEFAULT_STRUCTURE = STRUCTURE_LIST;


function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {name: "Export JSON for this sheet", functionName: "exportSheet"},
    {name: "Export JSON for all sheets", functionName: "exportAllSheets"}
  ];
  ss.addMenu("Export JSON", menuEntries);
}
 
function makeLabel(app, text, id) {
  var lb = app.createLabel(text);
  if (id) lb.setId(id);
  return lb;
}

function makeListBox(app, name, items) {
  var listBox = app.createListBox().setId(name).setName(name);
  listBox.setVisibleItemCount(1);
  
  var cache = CacheService.getPublicCache();
  var selectedValue = cache.get(name);
  Logger.log(selectedValue);
  for (var i = 0; i < items.length; i++) {
    listBox.addItem(items[i]);
    if (items[1] == selectedValue) {
      listBox.setSelectedIndex(i);
    }
  }
  return listBox;
}

function makeButton(app, parent, name, callback) {
  var button = app.createButton(name);
  app.add(button);
  var handler = app.createServerClickHandler(callback).addCallbackElement(parent);;
  button.addClickHandler(handler);
  return button;
}

function makeTextBox(app, name) { 
  var textArea    = app.createTextArea().setWidth('100%').setHeight('200px').setId(name).setName(name);
  return textArea;
}

function exportAllSheets(e) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetsData = {};
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var rowsData = getRowsData_(sheet, getExportOptions(e));
    var sheetName = sheet.getName(); 
    sheetsData[sheetName] = rowsData;
  }
  var json = makeJSON_(sheetsData, getExportOptions(e));
  displayText_(json);
}

function exportSheet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var rowsData = getRowsDataReversed_(sheet, getExportOptions(e));
  var json = makeJSON_(rowsData, getExportOptions(e));
  displayText_(json);
}
  
function getExportOptions(e) {
  var options = {};
  
  options.language = e && e.parameter.language || DEFAULT_LANGUAGE;
  options.format   = e && e.parameter.format || DEFAULT_FORMAT;
  options.structure = e && e.parameter.structure || DEFAULT_STRUCTURE;
  
  var cache = CacheService.getPublicCache();
  cache.put('language', options.language);
  cache.put('format',   options.format);
  cache.put('structure',   options.structure);
  
  Logger.log(options);
  return options;
}

function makeJSON_(object, options) {
  if (options.format == FORMAT_PRETTY) {
    var jsonString = JSON.stringify(object, null, 4);
  } else if (options.format == FORMAT_MULTILINE) {
    var jsonString = Utilities.jsonStringify(object);
    jsonString = jsonString.replace(/},/gi, '},\n');
    jsonString = prettyJSON.replace(/":\[{"/gi, '":\n[{"');
    jsonString = prettyJSON.replace(/}\],/gi, '}],\n');
  } else {
    var jsonString = Utilities.jsonStringify(object);
  }
  if (options.language == LANGUAGE_PYTHON) {
    // add unicode markers
    jsonString = jsonString.replace(/"([a-zA-Z]*)":\s+"/gi, '"$1": u"');
  }
  return jsonString;
}

function displayText_(text) {
  var output = HtmlService.createHtmlOutput("<textarea style='width:100%;' rows='20'>" + text + "</textarea>");
  output.setWidth(400)
  output.setHeight(300);
  SpreadsheetApp.getUi()
      .showModalDialog(output, 'Exported JSON');
}


// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader_(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum_(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit_(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
    
  }
  return key;
}
``
function nestKeysUnderParent(parentName, keys) {
  var result = {};

  result[parentName] = {};

  for (var i = 0; i < keys.length; i++) {
    result[parentName][keys[i]] = "";
  }

  return result;
}

// Reversed version of the getRowsData_ function to read data vertically
function getRowsDataReversed_(sheet, options) {
  var headersRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getFrozenColumns());
  var headers = headersRange.getValues();
  var dataRange = sheet.getRange(1, sheet.getFrozenColumns() + 1, sheet.getMaxRows(), sheet.getMaxColumns());
  var objects = getObjectsReversed_(dataRange.getValues(), normalizeHeadersReversed_(headers));
  return objects;
}

function renameAndNestKey(obj, keyToRename, parentKey, newKey, defaultValue = "REPLACE_TEXT") {
  if (obj.hasOwnProperty(keyToRename)) {
    if (parentKey === "" || parentKey === null) {
      obj[newKey] = obj[keyToRename];
    } else if (parentKey in obj && typeof obj[parentKey] === "object") {
      obj[parentKey][newKey] = obj[keyToRename];
    } else {
      obj[parentKey] = { [newKey]: obj[keyToRename] };
    }
    delete obj[keyToRename];
  } else {
    if (parentKey === "" || parentKey === null) {
      obj[newKey] = defaultValue;
    } else {
      if (!obj.hasOwnProperty(parentKey) || typeof obj[parentKey] !== "object") {
        obj[parentKey] = {};
      }
      if (!(newKey in obj[parentKey])) {
        obj[parentKey][newKey] = defaultValue; // You can set any default value for the new key here
      }
    }
  }
  return obj;
}

// Reversed version to read data vertically
function getObjectsReversed_(data, keys) {
  var objects = [];
  var lastNonEmptyColumnIndex = -1;
  var encounteredKeys = {};

  // Find the index of the last non-empty column
  for (var i = data[0].length - 1; i >= 0; i--) {
    var hasData = false;
    for (var j = 0; j < data.length; j++) {
      var cellData = data[j][i];
      if (cellData !== undefined && cellData !== null && cellData !== "") {
        hasData = true;
        break;
      }
    }
    if (hasData) {
      lastNonEmptyColumnIndex = i;
      break;
    }
  }

  // Create objects up to the last non-empty column index
  for (var i = 0; i <= lastNonEmptyColumnIndex; i++) {
    var object = {};
    for (var j = 0; j < data.length; j++) {
      var cellData = data[j][i];
      var property = keys[j];


      if (property === undefined || cellData === undefined || cellData === null) {
        continue; // Skip if header or cellData is undefined or null
      }

      // Check if the key already exists in the object
      if (object.hasOwnProperty(property)) {
        // If it exists, create an array and push the new value
        if (!Array.isArray(object[property])) {
          object[property] = [object[property]];
        }
        object[property].push(cellData);
      } else {
        // If it doesn't exist, set the value directly
        object[property] = cellData;
      }
    }

 
 /* //For implenting a rename function that renames desired keys, then uses a flag to delete all old keys that have not been renamed
  //only necessary if you only want to return all and only the renamed keys. Requires you to create a new function(ex. rename)
    let renameCalled = false;
    rename(object);
    renameCalled = true;

    if (renameCalled == true) {
      // Delete all the original keys from the object
      for (var j = 0; j < keys.length; j++) {
        delete object[keys[j]];
      }
    }
*/
    objects.push(object);

  }
  console.log(objects);
  return objects;
}


// Reversed version of the normalizeHeaders_ function
function normalizeHeadersReversed_(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader_(headers[i][0]);
    if (key.length > 0) {
      // console.log(key);
      keys.push(key);
    }
    else{
      break;
    }
  }
  return keys;
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum_(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit_(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit_(char) {
  return char >= '0' && char <= '9';
}
