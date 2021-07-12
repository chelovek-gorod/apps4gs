'use strict';

// ALL UI FUNCTIONS

const assApp = SpreadsheetApp.getActiveSpreadsheet();
const sheets = assApp.getSheets();

const ui = SpreadsheetApp.getUi();

const widget = HtmlService.createHtmlOutputFromFile("sidebar").setTitle("Consolidate sheets");

let spreadsheetsArr = [getSheetsListAtStart()];
let addedSpreadsheetNamesArr = [assApp.getName()];

// CACHE
const cache = CacheService.getDocumentCache();
/*
//cache.removeAll(["SsheetsArr", "NamesArr"]);

let cachedS = cache.get("cachedS");
let cachedN = cache.get("cachedN");
if (cachedS !== null && cachedN !== null) {
  showmessage('GET CACHE');
  spreadsheetsArr = JSON.parse(cachedS);
  addedSpreadsheetNamesArr = JSON.parse(cachedN);
} else {
  showmessage('NO CACHE');
  cache.put("cachedS", JSON.stringify(spreadsheetsArr));
  cache.put("cachedN", JSON.stringify(addedSpreadsheetNamesArr));
}
*/

// UI FUNCTIONS
function onOpen() {
  cache.removeAll(["cachedS", "cachedN"]);
  ui.createMenu("Consolidation")
    .addItem("Create new consolidation", "showSidebar")
    .addSeparator()
    .addItem("Fast consolidation sheets", "fastConsolidation")
    .addToUi();
}

function showSidebar() {
  cache.removeAll(["cachedS", "cachedN"]);
  ui.showSidebar(widget);
}

function showModal() {
  let htmlOutput = HtmlService.createHtmlOutputFromFile("modal")
  .setWidth(400).setHeight(500);
  ui.showModalDialog(htmlOutput, 'Add spreadsheets');
}

function showmessage(message) {
  ui.alert(message);
}

// CALLBACK FUNCTIONS
function getSheetsListAtStart() {
  let obj = {ss:assApp.getName(), id:assApp.getId(), si:[], sa:[]};
  for (let s of sheets) {
    obj.si.push(s.getName()); // including sheets
    obj.sa.push(s.getName()); // all sheets
  }
  return obj;
}

function getSheetsArr() {
  let cachedS = cache.get("cachedS");
  if (cachedS !== null) spreadsheetsArr = JSON.parse(cachedS);
  return spreadsheetsArr;
}

function addSpraedsheetsList() {
  let filesArr = [];
  let files = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
  while(files.hasNext()) {
    let f = files.next();
    filesArr.push({name:f.getName(),id:f.getId()});
  }
  let cachedN = cache.get("cachedN");
  if (cachedN !== null) addedSpreadsheetNamesArr = JSON.parse(cachedN);
  return {files: filesArr, names: addedSpreadsheetNamesArr};
}

function apdateCache(arr) {
  let cachedS = cache.get("cachedS");
  let cachedN = cache.get("cachedN");
  if (cachedS !== null && cachedN !== null) {
    spreadsheetsArr = JSON.parse(cachedS);
    addedSpreadsheetNamesArr = JSON.parse(cachedN);
  }
  for (let i of arr) {
    addedSpreadsheetNamesArr.push(i.name);
    let ss = SpreadsheetApp.openById(i.id).getSheets(); // ss{}
    let obj = {ss:i.name, id:i.id, si:[], sa:[]};
    for (let s of ss) {
      obj.si.push(s.getName()); // including sheets
      obj.sa.push(s.getName()); // all sheets
    }
    spreadsheetsArr.push(obj);
  }
  cache.put("cachedS", JSON.stringify(spreadsheetsArr));
  cache.put("cachedN", JSON.stringify(addedSpreadsheetNamesArr));
  ui.showSidebar(widget);
}

function resetCache() {
  cache.put("cachedS", JSON.stringify([]));
  cache.put("cachedN", JSON.stringify([]));
  //cache.removeAll(["cachedS", "cachedN"]);
}
