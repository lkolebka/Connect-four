var spreadsheet = SpreadsheetApp.getActive();
var color1 = spreadsheet.getRange('L7').getBackground() // get color player A
var color2 = spreadsheet.getRange('N7').getBackground() // get color player B
var player1 =  spreadsheet.getRange('L7').getValue() // get name player A n
var player2 =  spreadsheet.getRange('N7').getValue() // get name player B name


/// 


function checkwin(color1,color2){
  var spreadsheet = SpreadsheetApp.getActive();
let ss = SpreadsheetApp.getActiveSheet().getRange('c3:i8').getBackgrounds()
let colrex1 = new RegExp(color1, "g");
let colrex2 = new RegExp(color2, "g");

let txt = ss.join('-').replace(/#ffffff/g, "0").replace(colrex1, "1").replace(colrex2, "2").replace(/,/g, "") + '-'
let chk1 = /1{4}|(1.{7}){3}1|(1.{6}){3}1|(1.{8}){3}1/
let chk2 = /2{4}|(2.{7}){3}2|(2.{6}){3}2|(2.{8}){3}2/
if(chk1.test(txt)){
return player1 + " win"
  }
if(chk2.test(txt)){
return player2 + " win" 
}
return null
}

function test(){
console.log(SpreadsheetApp.getActiveSheet().getRange('c7').getBackground())
}


function AA () {
spreadsheet.getRange('L8').activate();
  spreadsheet.getActiveRangeList().setBackground(color1);
};


function showPrompt() {             //Show the welcome message
  var ui = SpreadsheetApp.getUi(); // Same variations.
 var spreadsheet = SpreadsheetApp.getActive();

  var result = ui.prompt(
      'Puissance 4',
      'Please enter the name of player A',
      ui.ButtonSet.OK);
 var result2 = ui.prompt(
      'Puissance 4',
      'Please enter the name of player B',
      ui.ButtonSet.OK);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
 var button2 = result2.getSelectedButton();
  var text2 = result2.getResponseText();
var spreadsheet = SpreadsheetApp.getActive();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert('Hi ' + text + ' & ' + text2 + ' !' +'\n May the best of you win.');
    
  spreadsheet.getRange('L7').activate(); //active nom player A in the right cell
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getActiveRange().mergeVertically();
  spreadsheet.getCurrentCell().setValue(text);
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontWeight('bold');

  spreadsheet.getRange('N7').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getActiveRange().mergeVertically();
  spreadsheet.getCurrentCell().setValue(text2);
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontWeight('bold');
 spreadsheet.getRange('L7').activate();

 spreadsheet.getRange('C3:I8').activate();
  spreadsheet.getActiveRangeList().setBackground(null)
  .clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('D10').activate();

  spreadsheet.getRange('L4').activate();
 var text = spreadsheet.getCurrentCell().setValue('Before you start, each of you must choose a color of token by colouring the square with your name. \n'+'***\n Once it\'s done, ' + text + ', you can start playing : select a cell and push your "play" button.');
  spreadsheet.getActiveRangeList().setBackground('#76A5B0');
   spreadsheet.getActiveRangeList().setFontColor('#ffffff')

  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('Don\'t wanna play ?');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('Don\'t wanna play ? ');
  spreadsheet.getRange('L4').activate();
 var text = spreadsheet.getCurrentCell().setValue('No player loaded :(');
  spreadsheet.getActiveRangeList().setBackground('#76A5B0');
   spreadsheet.getActiveRangeList().setFontColor('#ffffff')
  }
else if (button2 == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('Don\'t wanna play ? ');
     spreadsheet.getRange('L4').activate();
 var text = spreadsheet.getCurrentCell().setValue('No player loaded :(');
  spreadsheet.getActiveRangeList().setBackground('#76A5B0');
   spreadsheet.getActiveRangeList().setFontColor('#ffffff')

  }
else if (button2 == ui.Button.CANCEL) {
    // User clicked X in the title bar.
    ui.alert('Don\'t wanna play ? ');
     spreadsheet.getRange('L4').activate();
 var text = spreadsheet.getCurrentCell().setValue('No player loaded :(');
  spreadsheet.getActiveRangeList().setBackground('#76A5B0');
   spreadsheet.getActiveRangeList().setFontColor('#ffffff')

  }
 spreadsheet.getRange('L7').activate();  
spreadsheet.getActiveRangeList().setBackground('#FFFFFF');
 spreadsheet.getRange('N7').activate();  
spreadsheet.getActiveRangeList().setBackground('#FFFFFF');
     spreadsheet.getRange('L4').activate();
}


function A() { ///Player A turn
  var spreadsheet = SpreadsheetApp.getActive();
  var player1 =  spreadsheet.getRange('L7').getValue() 
  var player2 =  spreadsheet.getRange('N7').getValue()
  var range = spreadsheet.getActiveSheet().getSelection().getActiveRange(); 
  var cell = range.getA1Notation(); // sert à trouver les cordonnées de la cellule

  spreadsheet.getActiveRangeList().setBackground(color1);
  spreadsheet.getRange('L4').activate();
  spreadsheet.getCurrentCell().setValue(player1 + ' has just played '+'['+ cell+']' + ' \n\nIt\'s ' + player2 +'\'s turn !');
  spreadsheet.getActiveCell().activate();
  spreadsheet.getActiveRangeList().setBackground(color2);
 spreadsheet.getActiveRangeList().setFontColor('#000000')
  spreadsheet.getRange('N7').activate();
  SpreadsheetApp.flush()
  
  // check win 
let chk = checkwin(color1,color2);
  spreadsheet.getRange("L4").setBackground('#76A5B0');
chk&&spreadsheet.getRange('L4').setValue(chk);
};


function B() {///Player B turn
  var spreadsheet = SpreadsheetApp.getActive();
  var player1 =  spreadsheet.getRange('L7').getValue() 
  var player2 =  spreadsheet.getRange('N7').getValue()
  var activeCell =  spreadsheet.getActiveCell().activate();

   var range = spreadsheet.getActiveSheet().getSelection().getActiveRange(); // sert à trouver les corddoné de la cellule
   var cell = range.getA1Notation(); // same
  

  activeCell.setBackground(color2);
  activeCell.getDataSourceTables()
  spreadsheet.getRange('L4').activate();
  spreadsheet.getCurrentCell().setValue(player2 + ' has just played '+'['+ cell+']' + ' \n\nIt\'s ' + player1 +'\'s turn !');
  spreadsheet.getActiveCell().activate();
  spreadsheet.getActiveRangeList().setBackground(color1);
spreadsheet.getActiveRangeList().setFontColor('#000000')
  spreadsheet.getRange('L7').activate();
// spreadsheet.getRange('L4').setValue(checkwin());
SpreadsheetApp.flush()
let chk = checkwin(color1,color2);
spreadsheet.getRange("L4").setBackground('#76A5B0');
chk&&spreadsheet.getRange('L4').setValue(chk);
  
};

 function reset() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C3:I8').activate();
  spreadsheet.getActiveRangeList().setBackground(null)
  .clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('D10').activate();
  spreadsheet.getRange('L4').activate();
  spreadsheet.getCurrentCell().setValue('The game has been reset')
  spreadsheet.getActiveRangeList().setFontColor('#ffffff')
  spreadsheet.getActiveRangeList().setBackground('#76A5B0');

};


