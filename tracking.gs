function logFormulasCopiedDown_()
{
  var documentProperties = PropertiesService.getDocumentProperties();
  var systemName = documentProperties.getProperty("systemName");
  NVAddOns.log("Formulas%20Copied%20Down", scriptName, scriptTrackingId, systemName)
}


function logAuthEmailSent_()
{
  var documentProperties = PropertiesService.getDocumentProperties();
  var systemName = documentProperties.getProperty("systemName");
  NVAddOns.log("Reauthorization%20Email%20Sent", scriptName, scriptTrackingId, systemName)
}


function logRepeatInstall_() {
  var docProperties = PropertiesService.getDocumentProperties();
  var systemName = docProperties.getProperty('systemName');
  NVAddOns.log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
}

function logFirstInstall_() {
  var docProperties = PropertiesService.getDocumentProperties();
  var systemName = docProperties.getProperty('systemName');
  NVAddOns.log("First%20Install", scriptName, scriptTrackingId, systemName)
}

// Call this function from within the first major UI Setup step.
// Multiple calls to this function will not result in multiple install analytics

function setSid_() { 
  var docProperties = PropertiesService.getDocumentProperties();
  var userProperties = PropertiesService.getUserProperties();
  var scriptNameLower = scriptName.toLowerCase();
  var sid = docProperties.getProperty(scriptNameLower + "_sid");

  if (sid == null || sid == "")
  {
    incrementNumUses_();
    var dt = new Date();
    var ms = dt.getTime();
    var ms_str = ms.toString();
    docProperties.setProperty(scriptNameLower + "_sid", ms_str);
    var uid = userProperties.getProperty(scriptNameLower + "_uid");
    if (uid) {
      logRepeatInstall_();
    } else {
      logFirstInstall_();
      userProperties.setProperty(scriptNameLower + "_uid", ms_str);
    }      
  }
}


function incrementNumUses_() {
  try {
    var numCopyDownUses = PropertiesService.getUserProperties().getProperty('numCopyDownUses');
    if (parseInt(numCopyDownUses)) {
      numCopyDownUses = parseInt(numCopyDownUses) + 1;
    } else {
      numCopyDownUses = 1;
    }
    PropertiesService.getUserProperties().setProperty('numCopyDownUses', numCopyDownUses);
  } catch(err) {
    var errInfo = catchToString_(err);
    Browser.msgBox(errInfo);
  }
}
