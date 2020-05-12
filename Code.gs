var HDNS = "http://www.hotdocs.com/schemas/component_library/2009";
var HDPREFIX = "hd";
var hdNamespace = XmlService.getNamespace(HDPREFIX,HDNS);

function onInstall(e) {
  onOpen(e);
}

/**
 * Create the initial Hotdocs menus for the Sheet
 */
function onOpen() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var hdTitle = documentProperties.getProperty('hdTitle');
  var ui = SpreadsheetApp.getUi(); 
  var menu = ui.createMenu('HotDocs Computation Generator');
  if (hdTitle) {
    menu.addSubMenu(ui.createMenu('File')
                    .addItem('Open...','showPicker')
                    .addItem('Refresh Preview of "' + hdTitle + '"','refreshHDFile')
                    .addItem('Show ID','displayDocId'));
    menu.addSubMenu(ui.createMenu('Component')
                    .addItem('Properties','displayComponent')
                    .addItem('Load Dialog into Sheet', 'refreshDialog'));
    menu.addSubMenu(ui.createMenu('Computation')
                    .addItem('Text Before/After', 'showBeforeOrAfter')
                    .addItem('Script conditional variables (IF A SET B...)','showIfA')
                    .addItem('Script conditional variables (IF A = [x] SET B...)','showIfAeqX')
                    .addItem('Summarize (IF A RESULT + B)', 'showSummarize')
                    .addItem('Repeat (SET ColA[1] TO Row1)', 'showRepeat'));
  } else {
    menu.addItem('Select or Upload Hotdocs Component File','showPicker');
  }
  menu.addItem('Help','showHelp');
  
  // .addItem('Text before/after selection','showBeforeOrAfter')
  //    menu.addSubMenu(ui.createMenu('Component File')
  //                .addItem('Unlink All Components to Dialogs','showSidebar')
  
  menu.addToUi();
}

function showIfA() {
  showSidebar('IF A THEN SET B TO TRUE','IfAThenB');
}

function showIfAeqX() {
  showSidebar('IF A = [X] THEN SET B TO TRUE','IfAeqX');
}

function showBeforeOrAfter() {
  showSidebar('Add text before/after selection','BeforeOrAfter');
}

function showRepeat() {
  showSidebar('Generate REPEAT', 'Repeat');
}

function showSummarize() {
  showSidebar('Summarize columns','Summarize');
}

function showHelp() {
  showSidebar('About the HotDocs Computation Generator','Help');
}

function showSidebar(title,filename) {
  var html = HtmlService.createHtmlOutputFromFile(filename)
      .setTitle(title)
      .setWidth(450);
  SpreadsheetApp.getUi() 
      .showSidebar(html);
}


/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('Picker.html')
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a HotDocs Component File');
}

function displayDocId() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var ui = SpreadsheetApp.getUi();
  var hdID = documentProperties.getProperty('hdID');
  var hdTitle = documentProperties.getProperty('hdTitle');
  ui.alert('Document Title: ' + hdTitle + ' ID: ' + hdID);
}

/**
 * Search for the component in the current HD component file and display its properties
 */
function displayComponent() {
  var range = SpreadsheetApp.getActiveRange();
  var cell = range.getValue().trim();
  
  var documentProperties = PropertiesService.getDocumentProperties();
  var hdID = documentProperties.getProperty('hdID');
    
  var component =  getComponentAttributes(getComponent(hdID,cell));
  
  var header = '<html><head><link rel="stylesheet" href="https://unpkg.com/purecss@1.0.0/build/pure-min.css" integrity="sha384-nn4HPE8lTHyVtfCBi5yW9d20FjT8BJwUXyWZT9InLYax14RDjBj46LmSztkmNP9w" crossorigin="anonymous"></head><body>';
  var footer = '</body></html>';
  
  var cFlat = '<table class="pure-table pure-table-bordered"><caption>Properties of component "' + cell + '"' + '<thead><tr><td>Property</td><td>Value</td></tr></thead>';
  for (var e in component) {
    cFlat += '<tr><td>' + e + '</td><td>'+ component[e] + '</td></tr>';
  }
  cFlat += '</table>';
  
  var html = HtmlService.createHtmlOutput(header + cFlat + footer)
      .setTitle('Component Properties')
      .setWidth(450);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function displayDownloadURL() {
  
  var documentProperties = PropertiesService.getDocumentProperties();
  var ui = SpreadsheetApp.getUi();
  var hdID = documentProperties.getProperty('hdID');
  var hdTitle = documentProperties.getProperty('hdTitle');
  var file = DriveApp.getFileById(hdID);
  var url = file.getDownloadUrl();
  
  var html = HtmlService.createHtmlOutput('Download <a href="' + url + '">' + hdTitle + '</a>')
             .setWidth(150)
             .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html,'Download Component File');
    
}

function displayDocId() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var ui = SpreadsheetApp.getUi();
  var hdID = documentProperties.getProperty('hdID');
  var hdTitle = documentProperties.getProperty('hdTitle');
  ui.alert('Document Title: ' + hdTitle + ' ID: ' + hdID);
}


function processIfAThenBForm(formObject) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1; // set to 1 to skip header row, 2 to skip an example data row as well.
  return getHotdocsLinkComputation(sheet, startRow);
}

function processIfAeqXForm(formObject) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1; // set to 1 to skip header row, 2 to skip an example data row as well.
  return getHotdocsIfAeqXComputation(sheet, startRow);

}

function processBeforeAfterForm(formObject) {
  var cr = "<br/>";
  var text = "";
  var before = formObject.before.trim();
  var after = formObject.after.trim();
  var sheet = SpreadsheetApp.getActiveSheet();  
  var rows = sheet.getActiveRange().getValues();
  
  for (var i = 0; i<rows.length; i++) {
    text += before + ' ' + rows[i] + ' ' + after + cr;
  }
  return text;
}

function processRepeatForm(formObject) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var omitA = (formObject.skip == "skip");
  return getHotdocsRepeatComputation(sheet, omitA);
}

function processSummarizeForm(formObject) {
  var sheet = SpreadsheetApp.getActiveSheet();
  return getHotdocSummarizeComputation(sheet);
}


function getHotdocSummarizeComputation(sheet) {
  var cr = "<br/>";
  var header = '"«.p ""a, b, and c""»"' + cr;
  var text = "";
  var footer = 'RESULT + "«.pe»"'
  
  var rows = sheet.getDataRange().getValues();
  
  for (var i = 1; i<rows.length; i++) {
    text += 'IF VALUE(' + rows[i][0] + ') RESULT +"' + rows[i][1] +  '"' + cr;
  }
  
  return header + text + footer;

}


function getHotdocsRepeatComputation(sheet, omitA) {
  
  var text = "";
  var cr = "<br/>";
  
  var rows = sheet.getDataRange().getValues();
  var header = rows[0];
  var colStart = 0;
  if (omitA) { colStart = 1;}
  
  for (var i = 1; i < rows.length; i++ ) {
    var index = i-1; // start the HotDocs array with 0 although we're skipping row 0 in the Sheet
    for (var j = colStart; j < header.length; j++ ) {
      var val = rows[i][j];
      var end = '';
      // Check for special cases: date and boolean values must be transformed to HotDocs types
      if (isValidDate(val)) {
        end = hotdocsDate(val);        
      } else if (typeof(val) === typeof(true)) {
        if (val) {
          end = 'TRUE';
        } else {
          end = 'FALSE';
        }
      } else {
        end = '"' + val + '"';
      }
      text += 'SET ' + header[j] + '[' + index + '] TO ' + end + cr;
    }
  }
  return text;
}

function getHotdocsLinkComputation(sheet, startRow) {
  var text = "";
  var tab = "&nbsp;&nbsp;&nbsp;&nbsp;";
  var cr = "<br/>";
  if (!startRow) {
    var startRow = 1;
  }

  var data = sheet.getDataRange().getValues();
 
  for (var j = startRow; j < data.length; j++) { 
    i = j - startRow;
    if (data[j][0] && data[j][1]) { // skip rows with empty cells in either col A or B
      text += 'IF VALUE(' + data[j][0] +')' + cr;
      text += tab + "SET " + data[j][1] + ' TO TRUE' + cr;
      text += 'END IF' + cr;
    }
  }
  return text;
}

function getHotdocsIfAeqXComputation(sheet, startRow) {
  var text = "";
  var tab = "&nbsp;&nbsp;&nbsp;&nbsp;";
  var cr = "<br/>";
  if (!startRow) {
    var startRow = 1;
  }

  var data = sheet.getDataRange().getValues();
 
  for (var j = startRow; j < data.length; j++) { 
    i = j - startRow;
    if (data[j][0] && data[j][1]) { // skip rows with empty cells in either col A or B
      text += 'IF VALUE(' + data[j][0] +') = ' + data[j][1] + cr;
      text += tab + "SET " + data[j][2] + ' TO TRUE' + cr;
      text += 'END IF' + cr;
    }
  }
  return text;
}

/**
 * Iteratively search for a HotDocs component, using the name attribute of the component as a reference.
 * May be slow on large component files.
 */
function getComponent(docID,name) {
  var components = getComponents(docID);
  return getComponentFromXML(components, name);
}

/** 
 * Iteratively search for a HotDocs component, using the name attribute of the component as a reference.
 * Helper function to reduce file loads for repeated calls
 * @param xml the XmlService root Element to search
 * @param name the name attribute to match
 * @return a component attributes array [type, name, prompt, title]
 */
function getComponentFromXML(xmlRoot,name) { 
  var xml = xmlRoot.getChildren();
  
  for (var i = 0; i < xml.length; i++) { 
    var x = xml[i];
    var xname = x.getAttribute('name').getValue();
    if (xname == name) {
      return x;
    }
  }
}

function processComponentFile(docID) {
  var hotDocsXML = loadDriveXMLStripBOM(docID);
  var root = hotDocsXML.getRootElement();
  return root.getChild('components',hdNamespace).getValue();
}

/**
 * Load the element containing all components in a HotDocs XML file
 * @return Xml Element
 */
function getComponents(docID) {
  var hotDocsXML = loadDriveXMLStripBOM(docID);
  var root = hotDocsXML.getRootElement();
  return root.getChild('components',hdNamespace);
}

/** 
 * Store the file title and ID using Google Properties Service
 * This will persist through file save/loads until a new component file is 
 * associated with the document.
 */
function setHDFile(docID,title) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperties({'hdID': docID,
                                    'hdTitle': title});
  onOpen();
}

/**
 * Parameter free wrapper function - reloads the HD components into a labeled sheet
 */
function refreshHDFile() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var hdID = documentProperties.getProperty('hdID');
  var hdTitle = documentProperties.getProperty('hdTitle');
  
  loadHDFileIntoSheet(hdID,hdTitle);
}

/**
 * Create a new sheet with a list of the components, or load 
 * into an existing sheet named HDComponents
 */
function loadHDFileIntoSheet(docID, title) {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var componentsRoot = getComponents(docID);
  try {
    var sheet = spreadsheet.insertSheet(title);
  } catch (e) {
    var sheet = spreadsheet.getSheetByName(title);
    sheet.clear();
  }
  
  var components = componentsRoot.getChildren();
  sheet.appendRow(['Component Type','Name','Prompt','Title']);
  sheet.setFrozenRows(1);
  for (var i=0; i<components.length; i++) {
    var x = getComponentAttributes(components[i]);
    var a = [x['type'],x['name'], x['prompt'], x['title']];
    sheet.appendRow(a);
  }  
}

function refreshDialog() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var hdID = documentProperties.getProperty('hdID');
  var cell = SpreadsheetApp.getActiveRange().getValue().trim();
       
  loadDialogIntoSheet(hdID, cell);
}

/**
 * Create a new sheet with the contents of the selected dialog
 */
function loadDialogIntoSheet(docID, title) {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var componentsRoot = getComponents(docID);
  var dialog = getComponentFromXML(componentsRoot,title);
  var sheet = createOrClear(spreadsheet, title);
  sheet.appendRow(['Component Type','Name','Prompt','Title']);
  sheet.setFrozenRows(1);
  
  var items = getDialogContentsXML(dialog, title);
  
  for (var i=0; i<items.length; i++) {
    var z = getComponentAttributes(items[i]);
    var c = getComponentFromXML(componentsRoot, z['name']);  
    var x = getComponentAttributes(c);
    var a = [x['type'],x['name'], x['prompt'], x['title']];
    sheet.appendRow(a);
  }  

}

function createOrClear(sheet, title) {
  try {
    var sheet = spreadsheet.insertSheet(title);
  } catch (e) {
    var sheet = spreadsheet.getSheetByName(title);
    sheet.clear();
  }
  return sheet;
}

/**
 * Return an array of components (XML elements) in a dialog based on the active component file
 * @return Element[]
 */
function getDialogContents(docID, dialogName) {
  var components = getComponents(docID);
  return getDialogContentsXML(components, dialogName);
}

/** 
 * Internal function: return an array of the components in a dialog
 */
function getDialogContentsXML(xmlRoot, dialogName) {
  Logger.log('About to check for dialog ' + dialogName);
  //var dialogXML  = getComponentFromXML(xmlRoot,dialogName); // type Element
  Logger.log(xmlRoot.getAttribute('name').getValue());
  
  var contentsParent = xmlRoot.getChild('contents',hdNamespace); // look for child element named 'contents'
  //var contents = null;
  
  if (contentsParent) {
    return contentsParent.getChildren('item', hdNamespace); // dialog components in HotDocs XML are named 'item'
  }
}


/**
 * Returns the type, name, prompt, and title for a HotDocs Component (XmlService element) as an array
 */
function getComponentAttributes(element) {
  var x = element;
  var prompt = null;
  var title = null;
  var script = null;
    
  var type = x.getName();
  var name = x.getAttribute('name').getValue();
  var promptParent = x.getChild('prompt',hdNamespace);
  if (promptParent) {
    var prompt = promptParent.getValue(); 
  }
  var titleParent = x.getChild('title',hdNamespace);
  if (titleParent) { 
    var title = titleParent.getValue(); 
  }
  var scriptParent = x.getChild('script',hdNamespace);
  if (scriptParent) {
    var script = scriptParent.getValue();
  }
  
  return {'type': type, 'name': name, 'prompt': prompt, 'title': title, 'script': script};
}

/**
 * Returns the XML contents for a HotDocs Component File using Google XmlService class
 */
function loadDriveXMLStripBOM(docID) {
  var rawXml = DriveApp.getFileById(docID).getBlob().getDataAsString();
  // Remove the BOM marker. See: https://stackoverflow.com/questions/13024978/removing-bom-characters-from-ajax-posted-string
  // also see: https://stackoverflow.com/questions/19756252/google-apps-script-how-to-fix-content-is-not-allowed-in-prolog
  // It looks like HotDocs 11 automatically marks its XML with a BOM, which GAS doesn't like.
  var xmlNoBOM = rawXml.replace("\ufeff", "");
  
  return XmlService.parse(xmlNoBOM);
}

/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

function hotdocsDate(date) {  
  var monthnames = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG", "SEP","OCT","NOV","DEC"];
  return date.getDate() + " " + monthnames[date.getMonth()] + " " + date.getFullYear();
}

// From http://stackoverflow.com/questions/1353684
// Returns 'true' if variable d is a date object.
function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}
