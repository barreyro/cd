function testTriggerSet() {
  var test = setTriggerServerSide();
  debugger;
}


function testTriggerUnset() {
  var test = unsetTriggerServerSide();
  debugger;
}


function testGetTriggerState() {
  var test = getTriggerState();
  debugger;
}


function testGetFormUrl() {
  var test = getFormUrl();
  debugger;
}

function clearDocumentProperties() {
  PropertiesService.getDocumentProperties().deleteAllProperties(); 
}

function clearTriggerState() {
  copyDownProperties.deleteCopyDownDocumentProperty('triggerState');
}

function setTestProperty() {
  copyDownProperties.setCopyDownDocumentProperty('trigger', 'set');
 }


function getLastRowNumberFormats() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn());
  var numberFormats = lastRow.getNumberFormats();
  debugger;
}

function testGetAvailableHeaders() {
  var test = getAvailableHeaders(2);
  debugger;
}

