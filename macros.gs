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
    .createMenu('Game Menu')
    .addItem('Add Player', 'ADD')
    .addSeparator()
    .addItem('Fold Paper', 'FOLD')
    .addItem('Pass Paper', 'PASS')
    .addSeparator()
    .addItem('Reveal Answers', 'REVEAL')
    .addItem('Hide Answers', 'HIDE')
    .addSeparator()
    .addItem('Reset Card', 'RESET')
    .addItem('Erase Player Data', 'TOTALRESET')
    .addToUi();
}

//--------------------------------------------------------------------------------
// GAME FUNCTIONS
//--------------------------------------------------------------------------------

function ADD() {
  // Ask player for name and an animal
  var name = Browser.inputBox('Enter your name.');
  var card = Browser.inputBox('Give your card a fun animal name.');
  
  // Control for duplicate card names
  var itt = gA().getSheetByName(card);
  if (itt) { var card = card + Math.floor(Math.random() * 100) }
  
  // Add player and create new sheet for them
  addPlayer(name, card);
  createSheet(name, card);
}
function FOLD() {
  // Fold paper down to lock in your answer
  var ss = gAS();
  var stem = getStem();
  var props = getProps();
  var curSheet = getCurSheet();
  var curRow  = getCurRow();
  var nextRow = curRow + 1
  var setRow  = curRow - 1
  var val = ss.getRange(curRow, 2).getValue()
  
  // Get value just written
  props.setProperty(stem + setRow, val)  
  ss.getRange('B2:B7').clear({contentsOnly: true});
  
  // Reset colors to next
  var colorcode = props.getProperty(stem + 'color')
  ss.getRange('B2:B7').setBackground('#cccccc');
  if (curRow < 7){ ss.getRange(nextRow, 2).setBackground(colorcode) };
  
  // Set cur and next and status to Ready
  props.setProperty(stem + 'curRow', nextRow); 
  setStatus('ready', curSheet);
  
  // check if we need to pass
  checkPass()
}
function PASS() {
  // passes player to the next sheet in the lineup.
  var ss = gASS();
  var sheets = getPlayerSheets();
  var curSheet = getCurSheet();
  var nextSheet = curSheet + 1
  var stem = getStem();
  var props = getProps();
  var curRow = getCurRow();
  
  // Swap to next sheet, circling back if last.
  if (nextSheet == sheets.length) { nextSheet = 0 }
  next = ss.getSheetByName(sheets[nextSheet]);
  next.activate();
    
  // Reveal message if game is over
  if (curRow >= 8){ 
    Browser.msgBox('Time to reveal your answers! Click the "Reveal" button in the game menu.');
  } else { setStatus('waiting', curSheet) };
}
function RESET() {
  // Resets spreadsheet content and properties
  var ss = gA();
  var props = getProps();
  var curSheet = getCurSheet();
  var stem = getStem();
  
  // set colors
  var colorcode = props.getProperty(stem + 'color')
  ss.getRange('B2:B7').setBackground('#cccccc');
  ss.getRange('B2').setBackground(colorcode);
  
  // clear content
  ss.getRange('B2:B7').clear({contentsOnly: true});
  
  // set status
  setStatus('waiting', curSheet);
  
  // reset cached values
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
  var ss = gA();
  var sheets = getPlayerSheets();
  
  // reset and remove sheets
  for (inx in sheets) {
    var sheet = ss.getSheetByName(sheets[inx]);
    if (sheet) {
      sheet.activate();
      RESET();
      ss.deleteSheet(sheet);
    }
  }
  
  // clear player info on main sheet
  gSS(0).getRange('B5:D14').clear({contentsOnly: true}).setBackground('#ffffff');
}
function REVEAL() {
  // reveal everyone's answers
  var ss = gA();
  var props = getProps();
  var stem = getStem();
  
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
  
  var ss = gA();
  var props = getProps();
  var stem = getStem();
  var curRow = getCurRow();
  var colorcode = props.getProperty(stem + 'color');
  
  // set colors
  ss.getRange('B2:B7').setBackground('#cccccc');
  if (curRow < 8) { ss.getRange('B'+curRow).setBackground(colorcode) };
  
  // clear answers
  ss.getRange('B2:B7').clear({contentsOnly: true});
}


//--------------------------------------------------------------------------------
// PLAYER FUNCTIONS
//--------------------------------------------------------------------------------

function createSheet(name, card) {
  // Create a new sheet with player's name and animal
  var ss = gA();
  var props = getProps();
  var sheet = gSS(1).copyTo(ss);
  var inst = gSS(0);
  
  // Color
  var colorcode = randColor();
  
  // Update sheet with player info.
  sheet.activate();
  ss.getActiveSheet().setName(card);
  ss.getRange('B1').setValue(card).setFontColor(colorcode);
  ss.getRange('B2').setBackground(colorcode);
  
  var stem = getStem();
  props.setProperty(stem + 'color', colorcode)
};
function addPlayer(name, card) {
  // Add player to list at beginning.
  var sheet = gSS(0);
  sheet.activate();
  var inx = getFirstEmptyRow('B5:B14') + 5;
  
  sheet.getRange('B' + inx).setValue(name);
  sheet.getRange('C' + inx).setValue(card);
  setStatus('waiting', inx-5);
}
function bumpPlayers() {
  // Moves everyone's name to the right card.
  var sheet = gSS(0);
  var names = getPlayerNames()
  
  // Create shifted list
  var result = [];
  var L = names.length - 1
  result.push(names[L]) // move last to first
  for (i in names) { if (i < L) result.push(names[i]) };
  
  // Reset names in cells.
  var inx = 5 + L
  var range = 'B5:B' + inx
  sheet.getRange(range).setValues(result);
}


//--------------------------------------------------------------------------------
// UTILITY FUNCTIONS
//--------------------------------------------------------------------------------

// GET functions
function gA() { return SpreadsheetApp.getActive(); }
function gAS() { return SpreadsheetApp.getActiveSheet(); }
function gASS() { return SpreadsheetApp.getActiveSpreadsheet(); }
function gSS(inx) { return gASS().getSheets()[inx]; }
function getStem() { return 'val' + getCurSheet().toFixed() }
function getProps() { return PropertiesService.getScriptProperties() }
function getCurRow() { return parseInt(getProps().getProperty(getStem() + 'curRow')) }
function getPlayerSheets() { return getList('C5:C14') }
function getPlayerNames()  { return getList('B5:B14') }
function getCurSheet() { return findMatchingInList(gAS().getName(), getPlayerSheets()) }
function getFirstEmptyRow(range) {
  var values = gASS().getRange(range).getValues();
  var ct = 0;
  while ( values[ct][0] != "" ) { ct++ };
  return ct
}
function getList(range) {
  var sheet = gSS(0);
  var vals = sheet.getRange(range).getValues();
  var inx = findMatchingInList('', vals)
  return vals.slice(0, inx)
}


// MISC
function randColor() { return '#' + Math.ceil(Math.random() * 0xFFFFFF).toString(16); }
function checkPass() {
  // If everyone is ready, tell players to pass
  if (gSS(0).getRange('E5').getValue() == true) { 
    Browser.msgBox('Time to pass your paper! Click "Pass Paper" in the Game Menu.') 
    for (i in getPlayerSheets()) { setStatus('waiting', i) }
    bumpPlayers();
  };
}
function findMatchingInList(name, arr) {
  // to find which item in the list matches key
  var ct = 0;
  while ( arr[ct][0] != name ) { ct++ };
  return ct
}
function setStatus(status, inx) {
  // set status to ready or waiting
  var inx = 5 + inx
  var cell = gSS(0).getRange('D' + inx)
  if (status=='ready') {
    cell.setValue('READY');
    cell.setBackground('#0bb337');
  } else {
    cell.setValue('Waiting...');
    cell.setBackground('#FF5733');
  }
}

// eof