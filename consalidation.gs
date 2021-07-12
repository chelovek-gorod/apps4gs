'use strict';

let scriptStartTime = new Date();

//const assApp = SpreadsheetApp.getActiveSpreadsheet();
//const sheets = assApp.getSheets();
//const ui = SpreadsheetApp.getUi();

const namesArr = [];
const propsArr = [];
const linesArr = [];

let resultTitle = '';

class Line {
  constructor (name) {
    this.name = name;
  }

  readLine (s, line) {
    let ceil = 2;
    while (s.getRange(1, ceil).getValue() !== '') {
      let property = s.getRange(1, ceil).getValue();
      if (propsArr.indexOf(property) < 0) propsArr.push(property);
      let value = s.getRange(line, ceil).getValue();
      if (!(property in this)) this[property] = '';
      if (value !== '') this.crossingCeil(property, this[property], value);
      ceil++;
    }
  }

  crossingCeil (property, value, newValue) {
    // SUMM
    if (value === '') this[property] = newValue;
    else this[property] = value + newValue;
  }
}

// START FAST CONSOLIDATION SCRIPT
function fastConsolidation() {
  for (let s of sheets) {
    if (resultTitle === '') resultTitle = s.getRange(1,1).getValue();
    getName(s); //traid to fined line name in namesArr
  }
  fillResult();
  ui.alert('All done! \n Just in ' + (new Date() - scriptStartTime)/1000 + 'seconds');
}
// START CONSOLIDATION SCRIPT
function startConsolidation(arrSheetsIn) {
  /* [{ss: "tz4", id: "1_bYrB2xCoVPU_piiVVdPppeqL1y9L7toQ7YTHJn4564", si: Array(2), sa: Array(2)}] */
  let arrSheets = [];
  for (let i of arrSheetsIn) {
    for (let ii of i.si) {
      arrSheets.push(SpreadsheetApp.openById(i.id).getSheetByName(ii));
    }
  }
  /**/
  for (let s of arrSheets) {
    if (resultTitle === '') resultTitle = s.getRange(1,1).getValue();
    getName(s); //traid to fined line name in namesArr
  }
  fillResult();
  // ui.alert('All done! \n Just in ' + (new Date() - scriptStartTime)/1000 + 'seconds');
  return 'All done! \n Just in ' + (new Date() - scriptStartTime)/1000 + 'seconds';
}

function allDoneResponce(responce) {
  ui.alert(responce);
}

function getName(s) {
  let line = 2;
  while (s.getRange(line, 1).getValue() !== '') {
    let name = s.getRange(line, 1).getValue();
    let index = namesArr.indexOf(name);
    if (index < 0) {
      index = namesArr.length;
      namesArr.push(name);
      linesArr.push(new Line (name));
    }
    linesArr[index].readLine(s, line);
    line++;
  }
}

function fillResult() {
  let resultSheet;
  // create or clear resultSheet 
  if (assApp.getSheetByName("Results") !== null) {
    resultSheet = assApp.getSheetByName("Results");
    resultSheet.clear();
  } else {
    resultSheet = assApp.insertSheet(sheets.length+1);
    resultSheet.setName("Results");
  }

  resultSheet.getRange(1, 1).setValue(resultTitle).setFontWeight("bold");

  for (let prop = 0; prop < propsArr.length; prop++) {
    resultSheet.getRange(1, prop + 2,1).setValue(propsArr[prop]).setFontWeight("bold");
  }
  for (let line = 0; line < namesArr.length; line++) {
    resultSheet.getRange(line + 2,1).setValue(namesArr[line]).setFontWeight("bold");
    for (let prop in linesArr[line]) {
      resultSheet.getRange(line + 2, propsArr.indexOf(prop)+2).setValue(linesArr[line][prop]);
    }
  }
}
