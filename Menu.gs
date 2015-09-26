/*
The purpose of this function is to control the layout and content of the Menu bar at the top of the sheet. 
It is an onOpen function which means it is triggered whenever user opens/refreshes the document.
All other functions in this script file are triggered from the onOpen function.
*/

function onOpen() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ss.getSheetByName("Spreadsheet restoration").hideSheet();
  ss.getSheetByName("Pocketguide 1").hideSheet();
  ss.getSheetByName("Pocketguide 2").hideSheet();
  ss.getSheetByName("Deliverables reference").hideSheet();
  ss.getSheetByName("Gateways reference").hideSheet();
  ss.getSheetByName("Features reference").hideSheet();
  ss.getSheetByName("Data update").hideSheet();
  
  //ADDS THE MENU ENTRIES IN THE "CUSTOMISATION" MENU TAB AND ADDS THE MENU
  var menuCustomiserEntries = [ {name: "Features customisation", functionName: "featureopener"},
                      {name: "Deliverables customisation", functionName: "deliverableopener"}, 
                      {name: "Gateways customisation", functionName: "gatewayopener"},
                      null,
                      {name: "Spreadsheet restoration", functionName: "spreadsheetrestoration"}];
  ss.addMenu("Spreadsheet management", menuCustomiserEntries);
  
  //ADDS THE MENU ENTRIES IN THE "DOCUMENTATION" MENU TAB AND ADDS THE MENU
  var menuDocumentationEntries = [ {name: "Software Documentation", functionName: "documentation"},
                      null,
                      {name: "GPDS Pocketguide 1", functionName: "pocketguide1"},
                      {name: "GPDS Pocketguide 2", functionName: "pocketguide2"}];
  ss.addMenu("Documentation", menuDocumentationEntries);
}

//OPENS THE "FEATURES REFERENCE" SHEET
function featureopener(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Features reference"));
}

//OPENS THE "DELIVERABLES REFERENCE" SHEET
function deliverableopener(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Deliverables reference"));
}

//OPENS THE "GATEWAYS REFERENCE" SHEET
function gatewayopener(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Gateways reference"));
}

//OPENS THE "POCKETGUIDE 1" SHEET
function pocketguide1(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Pocketguide 1"));
}

//OPENS THE "POCKETGUIDE 2" SHEET
function pocketguide2(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Pocketguide 2"));
}

//OPENS THE "SPREADSHEET RESTORATION" SHEET
function spreadsheetrestoration(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Spreadsheet restoration"));
}

//DISPLAYS THE HYPERLINK TO THE DOCUMENTATION
function documentation(){
  showURL("https://docs.google.com/a/jaguarlandrover.com/document/d/13gPPJ7YixcIBxx4pdd1e5RQlRbG9w-dtczJ-4wn6sqI/edit");
}

//CREATES A USER INTERFACE (MESSAGE BOX) FOR THE HYPERLINK ABOVE
function showURL(href){
  var app = UiApp.createApplication().setHeight(50).setWidth(150);
  app.setTitle("Software Documentation");
  var link = app.createAnchor('Open in a new tab.', href).setId("link");
  app.add(link);  
  var doc = SpreadsheetApp.getActive();
  doc.show(app);
 }