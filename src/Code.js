/*
    The MIT License (MIT)
    
    Copyright (c) 2016 Tobias Klevenz (tobias.klevenz@gmail.com)
    
    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:
    
    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.
    
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
*/
/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Beta 7', 'showSidebar')
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

function getExportLink() {
  var file = Drive.Files.get(getId());
  return file['exportLinks']['text/html'];
}

function getFileName() {
  var table = DocumentApp.getActiveDocument().getBody().getTables()[0];
  var fileName = table.getCell(0, 1).getText() + " " + table.getCell(table.getNumRows() - 1, 1).getText();
  return fileName.replace(/\s/g, "_");
}

function getName() {
  return DocumentApp.getActiveDocument().getName();
}

function getExportParams() {
  var params = {
    source: getExportLink(),
    fname: getFileName(),
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
       ['Updated', ''],
       ['Disqus', ''],
       ['Author', ''],
       ['Category', ''],
       ['Position', ''],
       ['Location', ''],
       ['Advertorial', ''],
       ['Picture', ''],
       ['Template', '']
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