var scriptName = 'copyDown';
var scriptTrackingId = 'UA-48800213-7';

function onInstall(e) {
  buildFullModeMenu(e)
  launchCopyDownUi()
}

function onOpen(e) {
  buildFullModeMenu(e);
}


function buildFullModeMenu(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('copyDown settings', 'launchCopyDownUi');
  menu.addToUi();
}


function launchCopyDownUi() {
  var formUrl = getFormUrl(); 
  if (formUrl) {
    setSid_();
    try {
      var form = FormApp.openByUrl(formUrl);
    } catch (err) {
      SpreadsheetApp.getUi().alert("Oops! It appears you don't have edit rights on the attached form. copyDown can only work with Forms that you have editing rights on.");
      return;
    }
    var template = HtmlService.createTemplateFromFile('interface');
    SpreadsheetApp.getUi().showSidebar(template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle("Copy Down Formulas on Form Submit"));
  } else {
    var template = HtmlService.createTemplateFromFile('formCreator');
    SpreadsheetApp.getUi().showModalDialog(template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setHeight(100), "Attach a form to this Spreadsheet?");
  }
}

// replace with new Sheets method once it is added.
function getFormUrl(ss) {
  if (ss !== true) {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var formUrl = call(function() { return ss.getFormUrl(); });
  return formUrl;
}


function createForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formUrl = ss.getFormUrl();
  if (!formUrl) {
    var form = FormApp.create("Untitled Form");
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    var formEditUrl = form.getEditUrl();
    getSetFirstFormSubmission(form);
    return formEditUrl;
  } else {
    formUrlEdit = "Error: Looks like there is already a form attached to this spreadsheet";
  }
}


function testRunCopyDown() {
  var e = {};
   e.range = SpreadsheetApp.getActiveRange();
   runCopyDown(e);
}


function runCopyDown(e) {
  var lock = LockService.getDocumentLock();
  var hasLock = lock.tryLock(10000);
  //if (hasLock) {  //removed lock -- treats the lock like a rate limiter instead
    try {
      var authStatus = checkAuthStatus();
      if (authStatus) {  
        try {
          var range = e.range;
          var sheet = range.getSheet();
          var ss = sheet.getParent();
          var props = PropertiesService.getDocumentProperties();
          var sheetId = sheet.getSheetId().toString();
          props.setProperty('formSheetId', sheetId);
          var formulaRow = props.getProperty('formulaRow');
          formulaRow = formulaRow ? formulaRow : 2;
          formulaRow = !isNaN(formulaRow) ? parseInt(formulaRow) : 2;
          var row = range.getRow();
          var formUrl = getFormUrl(ss);
          var statusCol = getStatusCol(sheet);
          var excludeCols = checkAutoCratMergeCol(sheet);
          var copyDownPairs = getFormulaRangePairs(sheet, formulaRow, excludeCols);
          var asValuesPairs = getAsValuesPairs(sheet);
          var values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues();
          var message = '';
          for (var i=0; i<values.length; i++) {
            if ((values[i][statusCol-1] === "")&&(row!==formulaRow)) {
              message = '';
              var error = copyDownRow(sheet, row, copyDownPairs, asValuesPairs, formulaRow);
              if (error.indexOf('error')!==-1) {
                //var message = "Due to limitations in Apps Script, copyDown is not compatible with filters. \nDeleting this status message, removing all filters, and submitting a new form response will allow copyDown to this row.";
                var message = "copyDown could not complete " + error;
                sheet.getRange(row, statusCol).setValue(message);
              } else {
                var message = constructMessage(copyDownPairs, asValuesPairs, formulaRow);
                sheet.getRange(row, statusCol).setValue(message);
              }
              call(function() { SpreadsheetApp.flush(); });
            }
          }
          try {
            logFormulasCopiedDown_();
          } catch(err) {
            lock.releaseLock();
            return;
          }
        } catch(err) {
          lock.releaseLock();
          var errInfo = catchToString_(err);
          logErrInfo_(errInfo);
          return;
        }
      } else {
        lock.releaseLock();
        logErrInfo_("Authorization function failure");
        return;
      }
    } catch(err) {
      lock.releaseLock();
      var errInfo = catchToString_(err);
      logErrInfo_(errInfo);
      return;
    }
    lock.releaseLock();
    return;
  //} else {
  //  logErrInfo_("Failed to obtain lock");
  //  return;
 // }
}


function testCheckAutoCratMergeCol() {
  var sheetId = PropertiesService.getDocumentProperties().getProperty('formSheetId');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetById(sheetId, ss);
  var autoCratMergeCol = checkAutoCratMergeCol(sheet);
}


function checkAutoCratMergeCol(sheet) {
  var headers = getSheetHeaders(sheet);
  var colTextIncludes = "Link to merged Doc";
  var autoCratMergeCols = [];
  for (var i=0; i<headers.length; i++) {
    if (headers[i].indexOf(colTextIncludes)!==-1) {
      autoCratMergeCols.push(i+1);
    }
  }
  return autoCratMergeCols;
}


function checkAuthStatus() {
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  // Check if the actions of the trigger require authorizations
  // that have not been supplied yet -- if so, warn the active
  // user via email (if possible).  This check is required when
  // using triggers with add-ons to maintain functional triggers.
  var authStatus = authInfo.getAuthorizationStatus();
  if (authStatus == ScriptApp.AuthorizationStatus.REQUIRED) {
    // Re-authorization is required. In this case, the user
    // needs to be alerted that they need to reauthorize; the
    // normal trigger action is not conducted, since it authorization
    // needs to be provided first. Send at most one
    // 'Authorization Required' email a day, to avoid spamming
    // users of the add-on.
    var props = PropertiesService.getDocumentProperties();
    var lastAuthEmailDate = props.getProperty('lastAuthEmailDate');
    var today = new Date().toDateString();
    if (lastAuthEmailDate != today) {
      if (MailApp.getRemainingDailyQuota() > 0) {
        var html = HtmlService.createTemplateFromFile('AuthorizationEmail');
        html.url = authInfo.getAuthorizationUrl();
        html.addonTitle = addonTitle;
        var message = html.evaluate();
        MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
            'Authorization Required',
            message.getContent(), {
                name: addonTitle,
                htmlBody: message.getContent()
            }
        );
        logAuthEmailSent_();
      }
      props.setProperty('lastAuthEmailDate', today);
    }
    return false;
  } else {
    return true;
  }
}



function constructMessage(copyDownPairs, asValuesPairs, formulaRow) {
  if (!formulaRow) {
    formulaRow = 2;
  }
  var message = 'Copied down all formats, and formulas from row ' + formulaRow + ' in columns ';
  var count = 0;
  for (var i=0; i<copyDownPairs.length; i++) {
    if (count>0) {
      message += ", ";
    }
    if (copyDownPairs[i].start === copyDownPairs[i].end) {
      message += columnToLetter(copyDownPairs[i].start);
    } else {
      message += columnToLetter(copyDownPairs[i].start) + "-" + columnToLetter(copyDownPairs[i].end);
    }
    count++;
  }
  var count2 = 0;
  for (var i=0; i<asValuesPairs.length; i++) {
    if (count2===0) {
      message += ". Copied and pasted back values in column(s) ";
    } else {
      message += ", ";
    }
    if (asValuesPairs[i].start === asValuesPairs[i].end) {
      message += columnToLetter(asValuesPairs[i].start);
    } else {
      message += columnToLetter(asValuesPairs[i].start) + "-" + columnToLetter(asValuesPairs[i].end);
    }
    if (count2 === asValuesPairs.length - 1) {
      message += ".";
    }
    count2++;
  }
  if (count === 0) {
    message = "Copied down formats only from row " + formulaRow + ". No formulas detected."
  }
  return message;
}


function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


function getStatusCol(sheet) {
  var headers = getSheetHeaders(sheet);
  var statusColText = "Formula Copy Down Status";
  var colIndex = headers.indexOf(statusColText);
  var statusCol = 0;
  if (colIndex === -1) {
    var lastCol = sheet.getLastColumn();
    sheet.insertColumnAfter(lastCol);  
    statusCol = lastCol + 1;
    sheet.getRange(1, statusCol, sheet.getMaxRows(), 1).setWrap(true);
    sheet.setColumnWidth(statusCol, 250);
    var statusHeader = sheet.getRange(1, statusCol);
    statusHeader.setValue(statusColText).setBackground('purple').setFontColor('white').setFontWeight('bold').setNote('This column is needed by the copyDown Add-on');
    getSetFirstFormSubmission();
    SpreadsheetApp.flush();
    sheet.getRange(2, statusCol).setValue("Master formula row. Do not sort.");
    SpreadsheetApp.flush();
  } else {
    statusCol = colIndex + 1;
  }
  return statusCol;
}


function getSetFirstFormSubmission(form) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!form) {
     var formUrl = getFormUrl();
    if (formUrl) {
      var form = FormApp.openByUrl(formUrl);
    }
  }
  if (form) {
    var responses = form.getResponses();
    if (responses.length === 0) {
      var items = form.getItems();
      var response = form.createResponse().submit();
      Utilities.sleep(3000);  
      ss.toast("An empty form submission was just submitted to row 2. Plase build your formulas in that row");
    }
  }
}


function setAsValuesCols(pasteAsValues, selectAllSet) {
   PropertiesService.getDocumentProperties().setProperty('pasteAsValues', JSON.stringify(pasteAsValues));
   PropertiesService.getDocumentProperties().setProperty('selectAll', selectAllSet);
   return;
}


function getAsValuesCols() {
  var pasteAsValues = PropertiesService.getDocumentProperties().getProperty('pasteAsValues');
  if ((pasteAsValues)&&(pasteAsValues !== '')) {
    pasteAsValues = JSON.parse(pasteAsValues);
  } else {
    pasteAsValues = [];
  }
  return pasteAsValues;
}


function getAvailableHeaders(formulaRow, reset) {
  var storedFormulaRow = PropertiesService.getDocumentProperties().getProperty('formulaRow');
  if (!formulaRow) {
    formulaRow = storedFormulaRow ? storedFormulaRow: 2;
  }
  PropertiesService.getDocumentProperties().setProperty('formulaRow', formulaRow);
  var formUrl = getFormUrl(true);
  var sheet = getFormDestinationSheet(formUrl, reset);
  var excludeCols = checkAutoCratMergeCol(sheet);
  var headers = getSheetHeaders(sheet);
  var formulaRowRange = sheet.getRange(formulaRow,1,1,sheet.getLastColumn());
  sheet.setActiveRange(formulaRowRange);
  var formulas = formulaRowRange.getFormulas()[0];
  var availableHeaders = [];
  var storedAsValuesCols = getAsValuesCols();
  for (var i=0; i<formulas.length; i++) {
    if ((formulas[i] !== "")&&(excludeCols.indexOf(i+1)===-1)) {
      var thisHeader = {};
      thisHeader.header = headers[i];
      thisHeader.formula = formulas[i];
      if (storedAsValuesCols.indexOf(headers[i]) !== -1) {
        thisHeader.state = 1;
      } else {
        thisHeader.state = 0;
      }
      availableHeaders.push(thisHeader);
    }
  }
  var returnObj = {};
  returnObj.availableHeaders = availableHeaders; 
  returnObj.formulaRow = formulaRow;
  returnObj.selectAllSet = PropertiesService.getDocumentProperties().getProperty('selectAll');
  return returnObj;
}


function getSheetHeaders(sheet) {
  var headers = [];
  var lastRow = sheet.getLastRow();
  if (lastRow !== 0) {
    var lastCol = sheet.getLastColumn();
    var headersRange = sheet.getRange(1, 1, 1, lastCol);
    headers = convertRange(headersRange)[0];
  }
  return headers;
}


function getFormDestinationSheet(formUrl, reset) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var excludeTypes = ['PARAGRAPH_TEXT','SECTION_HEADER','IMAGE', 'PAGE_BREAK'];
  if (!reset) {
    var formSheetId = PropertiesService.getDocumentProperties().getProperty('formSheetId');
  }
  if (!formSheetId) {
    var sheets = ss.getSheets();
    var form = FormApp.openByUrl(formUrl);
    var items = form.getItems();
    var questionTitles = [];
    for (var i=0; i<items.length; i++) {
      var type = items[i].getType().toString();
      if (excludeTypes.indexOf(type) === -1) {
        questionTitles.push(items[i].getTitle());
      }
    }
    var exclude = false;
    var formSheet;
    var sheetsWithTimestamp = [];
    for (var i=0; i<sheets.length; i++) {
      exclude = false;
      var theseHeaders = getSheetHeaders(sheets[i]);
      for (var j=0; j<questionTitles.length; j++) {
        if (theseHeaders.indexOf("Timestamp") !== -1) {
          sheetsWithTimestamp.push(sheets[i]);
        }
        if (theseHeaders.indexOf(questionTitles[j]) === -1) {
          exclude = true;
          break;
        }
      }
      if (exclude) {
        continue;
      } else {
        formSheet = sheets[i];
        break;
      }
    }
    if (!formSheet) {
      formSheet = sheetsWithTimestamp[0];
    }
    getStatusCol(formSheet);
    ss.setActiveSheet(formSheet);
    PropertiesService.getDocumentProperties().setProperty('formSheetId', formSheet.getSheetId().toString());
  } else {
    formSheet = getSheetById(formSheetId, ss);
  }
  return formSheet;
}




function getFormulaRangePairs(sheet, formulaRow, excludeCols) {
  if (!formulaRow) {
    formulaRow = 2;
  }
  var row2Formulas = sheet.getRange(formulaRow, 1, 1, sheet.getLastColumn()).getFormulas()[0];
  var colsWithFormulas = [];
  for (var i=0; i<row2Formulas.length; i++) {
    if ((row2Formulas[i]!=="")&&(excludeCols.indexOf(i+1)===-1)) {
      colsWithFormulas.push(i+1);
    }
  }
  var rangePairs = formContiguousPairs(colsWithFormulas);
  return rangePairs; 
}


function getAsValuesPairs(sheet) {
  var properties = PropertiesService.getDocumentProperties().getProperties();
  var headers = getSheetHeaders(sheet);
  var pasteAsValues =  properties.pasteAsValues ? JSON.parse(properties.pasteAsValues) : []; 
  var asValuesCols = [];
  for (var i=0; i<headers.length; i++) {
    if (pasteAsValues.indexOf(headers[i])!==-1) {
      asValuesCols.push(i+1);
    }
  }
  var rangePairs = formContiguousPairs(asValuesCols);
  return rangePairs;
}


function waitAndGiveValue() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Lookup Sheet');
  var values = sheet.getDataRange().getValues();
  for (var i=0; i<values.length; i++) {
    if (values[i][0] === 2000225194) {
      return values[i][1];
    }
  }
}


function copyDownRow(sheet, row, copyDownPairs, asValuesPairs, formulaRow) {
  try {
    if (!formulaRow) {
      formulaRow = 2;
    }
    var lastColumn = sheet.getLastColumn();
    var sourceRange = sheet.getRange(formulaRow, 1, 1, lastColumn);
    var destRange = sheet.getRange(row, 1, 1, lastColumn);
    sourceRange.copyTo(destRange, {formatOnly: true});
    for (var i=0; i<copyDownPairs.length; i++) {
      var sourceRange = sheet.getRange(formulaRow, copyDownPairs[i].start, 1, (copyDownPairs[i].end - copyDownPairs[i].start + 1));
      var destRange = sheet.getRange(row, copyDownPairs[i].start, 1, (copyDownPairs[i].end - copyDownPairs[i].start + 1));
      sourceRange.copyTo(destRange);
    }
    for (var i=0; i<asValuesPairs.length; i++) {
      var destRange = sheet.getRange(row, asValuesPairs[i].start, 1, (asValuesPairs[i].end - asValuesPairs[i].start + 1));
      var values = destRange.getValues();
      destRange.setValues(values);
    }
    return "success";
  } catch(err) {
    return "error: " + err.message;
  }
}



function formContiguousPairs(colsWithFormulas) {
  var contiguousPairs = [];
  var initiated = false;
  for (var i=0; i<colsWithFormulas.length; i++) {
    if (initiated === false) {
      var thisPair = {};
      thisPair.start = colsWithFormulas[i];
      initiated = true;
    }
    //columns are contiguous
    if ((colsWithFormulas[i] == colsWithFormulas[i+1]-1)) {
      continue;
    } else {
      thisPair.end = colsWithFormulas[i];
      initiated = false;
      contiguousPairs.push(thisPair);
    }  
  }
  return contiguousPairs;
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}



function unsetTriggerServerSide() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = call(function() { return ScriptApp.getUserTriggers(ss);});
  var found = false;
    for (var i=0; i<triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "runCopyDown") {
        ScriptApp.deleteTrigger(triggers[i]);
        copyDownProperties.deleteCopyDownDocumentProperty('triggerState');
        return "removed trigger";
      }
    }
    copyDownProperties.deleteCopyDownDocumentProperty('triggerState');
    return "no trigger found";
}



function setTriggerServerSide() {
  var triggerState = triggerIsSet();
  var user = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = call(function() { return ScriptApp.getUserTriggers(ss);});
  var found = false;
  for (var i=0; i<triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runCopyDown") {
      found = true;
    }
  }
  if (!triggerState) {
    triggerState = {user: user};
    if (!found) {
      ScriptApp.newTrigger('runCopyDown').forSpreadsheet(ss).onFormSubmit().create();
      copyDownProperties.setCopyDownDocumentProperty('triggerState', JSON.stringify(triggerState));
      return "success";
    } else {
      copyDownProperties.setCopyDownUserProperty('triggerState', JSON.stringify(triggerState));
      return "already set";
    }
  } else if (triggerState.user === user) {
    if (!found) {
      ScriptApp.newTrigger('runCopyDown').forSpreadsheet(ss).onFormSubmit().create();
      copyDownProperties.setCopyDownDocumentProperty('triggerState', JSON.stringify(triggerState));
      return "set by this user";
    }
    return "set by this user";
  } else {
    return triggerState.user;
  }
}


function getTriggerState() {
  var returnObj = {};
  returnObj.triggerState = triggerIsSet();
  var formulaRow = PropertiesService.getDocumentProperties().getProperty('formulaRow');
  returnObj.formulaRow = formulaRow ? parseInt(formulaRow) : 2;
  var user = Session.getActiveUser().getEmail();
  if (!returnObj.triggerState) {
    return returnObj;
  } else if (returnObj.triggerState.user === user) {
    returnObj.triggerState = "this_user";
    return returnObj;
  } else {
    return returnObj;
  }
}


function triggerIsSet() {
  var triggerState = copyDownProperties.getCopyDownDocumentProperty('triggerState');
  if (triggerState) {
    triggerState = JSON.parse(triggerState);
    return triggerState;
  } else {
    return false;
  }
}


function getSheetById(sheetId, spreadsheet) {
  try {
    sheetId = parseFloat(sheetId);
    var sheets = spreadsheet.getSheets();
    for (var i=0; i<sheets.length; i++) {
      if (sheets[i].getSheetId() == sheetId) {
        return sheets[i];
      }
    }
    PropertiesService.getDocumentProperties().deleteProperty('formSheetId');
    return;
  } catch(err) {
    throw(err.message);
    return;
  }
}



/**
* Invokes a function, performing up to 5 retries with exponential backoff.
* Retries with delays of approximately 1, 2, 4, 8 then 16 seconds for a total of 
* about 32 seconds before it gives up and rethrows the last error. 
* See: https://developers.google.com/google-apps/documents-list/#implementing_exponential_backoff 
* <br>Author: peter.herrmann@gmail.com (Peter Herrmann)
<h3>Examples:</h3>
<pre>//Calls an anonymous function that concatenates a greeting with the current Apps user's email
var example1 = GASRetry.call(function(){return "Hello, " + Session.getActiveUser().getEmail();});
</pre><pre>//Calls an existing function
var example2 = GASRetry.call(myFunction);
</pre><pre>//Calls an anonymous function that calls an existing function with an argument
var example3 = GASRetry.call(function(){myFunction("something")});
</pre><pre>//Calls an anonymous function that invokes DocsList.setTrashed on myFile and logs retries with the Logger.log function.
var example4 = GASRetry.call(function(){myFile.setTrashed(true)}, Logger.log);
</pre>
*
* @param {Function} func The anonymous or named function to call.
* @param {Function} optLoggerFunction Optionally, you can pass a function that will be used to log 
to in the case of a retry. For example, Logger.log (no parentheses) will work.
* @return {*} The value returned by the called function.
*/
function call(func, optLoggerFunction) {
  for (var n=0; n<6; n++) {
    try {
      return func();
    } catch(e) {
      if (optLoggerFunction) {optLoggerFunction("GASRetry " + n + ": " + e)}
      if (n == 5) {
        throw e;
      } 
      Utilities.sleep((Math.pow(2,n)*1000) + (Math.round(Math.random() * 1000)));
    }    
  }
}




function catchToString_(err) {
  var errInfo = "Caught something:\n"; 
  for (var prop in err)  {  
    errInfo += "  property: "+ prop+ "\n    value: ["+ err[prop]+ "]\n"; 
  } 
  errInfo += "  toString(): " + " value: [" + err.toString() + "]"; 
  return errInfo;
}

function logErrInfo_(errInfo) {
  var ss = SpreadsheetApp.openById('1PvAoUzC9o1d99HSY8x-G-olznhixZZobGpkSBGWjU-0');
  var sheet = ss.getSheets()[0];
  var date = new Date();
  sheet.getRange(sheet.getLastRow()+1, 1, 1, 2).setValues([[date, errInfo]]);;
}

