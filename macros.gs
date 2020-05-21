// Author:   Margaret Swift
// Contact:  swift.margaret.e@gmail.com
// Created:  19 May 2020
//
// Macros to run a game of Consequences with the Dinwiddie crew. 


//--------------------------------------------------------------------------------
// SETUP
//--------------------------------------------------------------------------------

function onOpen(e) {
  // Add items to game menu
  SpreadsheetApp.getUi()
    .createMenu('Game Menu') // .createAddonMenu()
    .addItem('Add Player', 'ADD')
    .addItem('Fold Paper', 'FOLD')
    .addItem('Pass Paper', 'PASS')
    .addItem('Reveal Answers', 'REVEAL')
    .addItem('Hide Answers', 'HIDE')
    .addItem('Reset Card', 'RESET')
    .addItem('Erase Player Data', 'TOTALRESET')
    .addToUi();
}

//--------------------------------------------------------------------------------
// GAME FUNCTIONS
//--------------------------------------------------------------------------------

function ADD() {
  // Ask player for name and an animal
  var ss = SpreadsheetApp.getActive();
  var name = Browser.inputBox('Enter your name.');
  var card = Browser.inputBox('Give your card a fun animal name.');
  
  // Control for duplicate card names
  var itt = ss.getSheetByName(card);
  if (itt) { var card = card + Math.floor(Math.random() * 100) }
  
  // Add player and create new sheet
  addPlayer(name, card);
  createSheet(name, card);
}
function FOLD() {
  // Fold paper down to lock in your answer
  
  var ss = getSS();
  var stem = 'val' + getCurSheet().toFixed();
  var props = PropertiesService.getScriptProperties();
  
  // get current properties
  var curRow = parseInt(props.getProperty(stem + 'curRow'));
  var curSheet = getCurSheet();
  var nextRow = curRow + 1
  var val = ss.getRange(curRow, 2).getValue()
  
  
  // Get value just written
  var setRow = curRow-1
  props.setProperty(stem + setRow, val)  
  ss.getRange('B2:B7').clear({contentsOnly: true});
  
  // Reset colors to next
  var colorcode = ss.getRange('A1').getValue()
  ss.getRange('B2:B7').setBackground('#cccccc');
  if (curRow < 7){ 
    ss.getRange(nextRow, 2).setBackground(colorcode) 
  };
  
  // Set cur and next and status to Ready
  props.setProperty(stem + 'curRow', nextRow); 
  setStatus('ready', curSheet);
  
  // check if we need to pass
  checkPass()
}
function PASS() {
  // passes player to the next sheet in the lineup.
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var curSheet = getCurSheet();
  var sheets = getPlayerSheets()

  // Swap to next sheet, circling back if last.
  var nextSheet = curSheet + 1
  if (nextSheet == sheets.length) { nextSheet = 0 }
  next = ss.getSheetByName(sheets[nextSheet]);
  next.activate();
    
  // Reveal message if game is over
  var stem = 'val' + getCurSheet().toFixed();
  var props = PropertiesService.getScriptProperties();
  var curRow = parseInt(props.getProperty(stem + 'curRow'));
  if (curRow >= 8){ 
    Browser.msgBox('Time to reveal your answers! Click the "Reveal" button in the game menu.');
  } else { setStatus('waiting', curSheet) };
}
function RESET() {
  // Resets spreadsheet content and properties
  
  // get spreadsheet and properties
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties();
  var curSheet = getCurSheet();
  // set colors
  var colorcode = ss.getRange('A1').getValue();
  ss.getRange('B2:B7').setBackground('#cccccc');
  ss.getRange('B2').setBackground(colorcode);
  // clear content
  ss.getRange('B2:B7').clear({contentsOnly: true});
  // set status
  setStatus('waiting', curSheet);
  
  // reset cached values
  var stem = 'val' + getCurSheet().toFixed();
  props.setProperty(stem + 'curRow', 2);
  props.setProperty(stem + '1', 'no answer yet');
  props.setProperty(stem + '2', 'no answer yet');
  props.setProperty(stem + '3', 'no answer yet');
  props.setProperty(stem + '4', 'no answer yet');
  props.setProperty(stem + '5', 'no answer yet');
  props.setProperty(stem + '6', 'no answer yet');
}
function TOTALRESET() {
  // resets game completely to just template and instructions.
  
  Browser.msgBox("Make sure you want to erase all players before proceeding!")
  
  // Resets whole game and removes players
  var ss = SpreadsheetApp.getActive()
  var sheets = getPlayerSheets();
  var inst = ss.getSheetByName('Instructions');
  
  // reset and remove sheets
  for (inx in sheets) {
    var sheet = ss.getSheetByName(sheets[inx]);
    if (sheet) {
      sheet.activate();
      RESET();
      ss.deleteSheet(sheet);
    }
  }
  
  // clear player info
  inst.getRange('B5:D14').clear({contentsOnly: true}).setBackground('#ffffff');
}
function REVEAL() {
  // reveal everyone's answers
  
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties();
  var stem = 'val' + getCurSheet().toFixed();
  
  // remove colors
  ss.getRange('B2:B7').setBackground('#ffffff');
  
  // get values
  ss.getRange('B2').setValue(props.getProperty(stem + '1'));
  ss.getRange('B3').setValue(props.getProperty(stem + '2'));
  ss.getRange('B4').setValue(props.getProperty(stem + '3'));
  ss.getRange('B5').setValue(props.getProperty(stem + '4'));
  ss.getRange('B6').setValue(props.getProperty(stem + '5'));
  ss.getRange('B7').setValue(props.getProperty(stem + '6'));
}
function HIDE() {
  // undo reveal answers
  
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties();
  var stem = 'val' + getCurSheet().toFixed();
  var curRow = parseInt(props.getProperty(stem + 'curRow'));
  var colorcode = ss.getRange('A1').getValue()
  
  // set colors
  ss.getRange('B2:B7').setBackground('#cccccc');
  ss.getRange('B'+curRow).setBackground(colorcode);
  
  // clear answers
  ss.getRange('B2:B7').clear({contentsOnly: true});
}


//--------------------------------------------------------------------------------
// PLAYER FUNCTIONS
//--------------------------------------------------------------------------------

function createSheet(name, card) {
  // Create a new sheet with player's name and animal
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Template').copyTo(ss);
  var inst = ss.getSheetByName('Instructions');
  
  // Color
  var colorcode = '#' + Math.ceil(Math.random() * 0xFFFFFF).toString(16);

  // Update sheet with player info.
  sheet.activate();
  ss.getActiveSheet().setName(card);
  ss.getRange('B1').setValue(card).setFontColor(colorcode);
  ss.getRange('B2').setBackground(colorcode);
  ss.getRange('A1').setValue(colorcode);
};
function addPlayer(name, card) {
  // Add player to list at beginning.
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Instructions');
  var inx = getFirstEmptyRow('B5:B14') + 5;
  
  sheet.getRange('B' + inx).setValue(name);
  sheet.getRange('C' + inx).setValue(card);
  setStatus('waiting', inx-5);
}
function getPlayerSheets() {
  // find all the sheets that are actual gameplay (not instructions/template)
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Instructions');
  var sheets = sheet.getRange('C5:C14').getValues();
  var inx = findMatchingInList('', sheets)
  var sheets = sheets.slice(0, inx)
  return(sheets);
}
function getPlayerNames() {
  // like it says on the tin
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Instructions');
  var names = sheet.getRange('B5:B14').getValues();
  var inx = findMatchingInList('', names)
  var names = names.slice(0, inx)
  return(names);
}
function bumpPlayers() {
  // move everyone's name to the right card.
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Instructions');
  var names = getPlayerNames()
  
  // Rearrange names to bump last to first
  var result = [];
  var L = names.length-1
  result.push(names[L])
  for (i in names) {
    if (i < L) result.push(names[i])
  }
  
  // reset names in cells.
  var inx = 5 + L
  var range = 'B5:B' + inx
  sheet.getRange(range).setValues(result);
}


//--------------------------------------------------------------------------------
// UTILITY FUNCTIONS
//--------------------------------------------------------------------------------

function findMatchingInList(name, arr) {
  // to find which item in the list matches key
  var ct = 0;
  while ( arr[ct][0] != name ) { ct++ };
  return (ct);
}
function getFirstEmptyRow(range) {
  // find the end of the list
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var column = ss.getRange(range);
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct][0] != "" ) { ct++ };
  return (ct);
}
function getCurSheet() {
  // get current sheet's index
  var name = getSS().getName();
  var sheets = getPlayerSheets();
  var inx = findMatchingInList(name, sheets);
  return inx
}
function getSS() {
  // shortcut for this gobbledygook
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}
function setStatus(status, inx) {
  // set status to ready or waiting
  var inx = 5 + inx
  var cell = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange('D' + inx)
  if (status=='ready') {
    cell.setValue('READY');
    cell.setBackground('#0bb337');
  } else {
    cell.setValue('Waiting...');
    cell.setBackground('#fcba03');
  }
}
function checkPass() {
  // If everyone is ready, tell players to pass
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Instructions');
  if (sheet.getRange('E5').getValue() == true) { 
    Browser.msgBox('TIME TO PASS!') 
    var sheets = getPlayerSheets();
    for (i in sheets) { setStatus('waiting', i) }
    bumpPlayers();
  };
}


