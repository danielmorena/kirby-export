/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Kirby Export')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui);
}


function getId() {
  return DocumentApp.getActiveDocument().getId();
}

function getOauthToken() {
  return ScriptApp.getOAuthToken();
}


function getAuthStatus() {
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var status = authInfo.getAuthorizationStatus()
  
  if (status == ScriptApp.AuthorizationStatus.REQUIRED)
    return "AUTH REQUIRED";
  else 
    return "AUTHORIZED";
}

function getExportLink() {
  var file = Drive.Files.get(getId());
  return file['exportLinks']['text/html'];
}

function getName() {
  return DocumentApp.getActiveDocument().getName();
}

function getExportParams() {
  var params = {
    source: getExportLink(),
    fname: getName(),
    token: getOauthToken()
  };
  
  return params;
}

function insertMetaInfoTable() {
  var body = DocumentApp.getActiveDocument().getBody();
      
  if (body.getChild(0).isAtDocumentEnd()) { 
    insertTable();
  } else {
  
    var c0 = body.getChild(0);
    var c1 = body.getChild(1);
  
    if (c0.getType() == DocumentApp.ElementType.PARAGRAPH && c0.asParagraph().getText() == "" &&
        c1.getType() == DocumentApp.ElementType.TABLE) {
      throw "To insert a new table please delete table at the beginning of the document."; 
    } else {
      insertTable();
    }  
  }
}

function insertTable() {
  var body = DocumentApp.getActiveDocument().getBody();
  var cells = [
       ['Directory', ''],
       ['Title', ''],
       ['Description', ''],
       ['Tags', ''],
       ['TopNews', ''],
       ['Author', ''],
       ['Category', ''],
       ['Position', ''],
       ['Location', ''],
       ['Advertorial', ''],
       ['Picture', '']
     ];
     
  var table = body.insertTable(0, cells);
  table.setColumnWidth(0, 90);
     
  body.insertPageBreak(1);
}


function saveScriptProperties(prefs) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties(prefs);
}

function saveDocumentProperties(prefs) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperties(prefs);
}

function getScriptProperties() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperties();
}

function getDocumentProperties() {
  var documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperties();
}/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Kirby Export')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui);
}


function getId() {
  return DocumentApp.getActiveDocument().getId();
}

function getOauthToken() {
  return ScriptApp.getOAuthToken();
}


function getAuthStatus() {
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var status = authInfo.getAuthorizationStatus()
  
  if (status == ScriptApp.AuthorizationStatus.REQUIRED)
    return "AUTH REQUIRED";
  else 
    return "AUTHORIZED";
}

function getExportLink() {
  var file = Drive.Files.get(getId());
  return file['exportLinks']['text/html'];
}

function getName() {
  return DocumentApp.getActiveDocument().getName();
}

function getExportParams() {
  var params = {
    source: getExportLink(),
    fname: getName(),
    token: getOauthToken()
  };
  
  return params;
}

function insertMetaInfoTable() {
  var body = DocumentApp.getActiveDocument().getBody();
      
  if (body.getChild(0).isAtDocumentEnd()) { 
    insertTable();
  } else {
  
    var c0 = body.getChild(0);
    var c1 = body.getChild(1);
  
    if (c0.getType() == DocumentApp.ElementType.PARAGRAPH && c0.asParagraph().getText() == "" &&
        c1.getType() == DocumentApp.ElementType.TABLE) {
      throw "To insert a new table please delete table at the beginning of the document."; 
    } else {
      insertTable();
    }  
  }
}

function insertTable() {
  var body = DocumentApp.getActiveDocument().getBody();
  var cells = [
       ['Directory', ''],
       ['Title', ''],
       ['Description', ''],
       ['Tags', ''],
       ['TopNews', ''],
       ['Author', ''],
       ['Category', ''],
       ['Position', ''],
       ['Location', ''],
       ['Advertorial', ''],
       ['Picture', '']
     ];
     
  var table = body.insertTable(0, cells);
  table.setColumnWidth(0, 90);
     
  body.insertPageBreak(1);
}


function saveScriptProperties(prefs) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties(prefs);
}

function saveDocumentProperties(prefs) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperties(prefs);
}

function getScriptProperties() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperties();
}

function getDocumentProperties() {
  var documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperties();
}