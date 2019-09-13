/** @OnlyCurrentDoc */

function runningTotal (){
  var spreadsheet = SpreadsheetApp.getActiveSheet(); // will need to exception handle this in case of null value.
  var interface = SpreadsheetApp.getUi();
  var lastRow = spreadsheet.getLastRow();
  var column;
  var startRow;
  var step;
  var lastTotal;
  
  //User input loop.
  while(true){
    
    //Column input
    var userInputColumn = interface.prompt('Please enter a column number.','Please enter a column number.' , interface.ButtonSet.OK_CANCEL);
    if (userInputColumn.getSelectedButton() == interface.Button.CANCEL || userInputColumn.getSelectedButton() == interface.Button.CLOSE){
      return;
    }
    column = userInputColumn.getResponseText();
    while(isNaN(column) || column < 1){
      var userInputColumn = interface.prompt('The column address must be a positive, non-zero intager.','Please enter an number.' , interface.ButtonSet.OK_CANCEL);
      if (userInputColumn.getSelectedButton() == interface.Button.CANCEL || userInputColumn.getSelectedButton() == interface.Button.CLOSE){
        return; 
      }
      column = userInputColumn.getResponseText();
    }
     
    //Row input.
    var userInputRow = interface.prompt('Please enter the number of the row to start from.','Please enter a number.' , interface.ButtonSet.OK_CANCEL);
    if (userInputRow.getSelectedButton() == interface.Button.CANCEL || userInputRow.getSelectedButton() == interface.Button.CLOSE){
      return;
    }
    startRow = userInputRow.getResponseText();
    while(isNaN(startRow) || startRow <= 0){
      var userInputRow = interface.prompt('The row address must be a positive, non-zero intager.','Please enter a row number.' , interface.ButtonSet.OK_CANCEL);
      if (userInputRow.getSelectedButton() == interface.Button.CANCEL || userInputRow.getSelectedButton() == interface.Button.CLOSE){
        return; 
      }
      startRow = userInputRow.getResponseText();
    }
    // If startRow is at the top of the column.
    if (startRow < 2){
      var InputLastTotal = interface.prompt("Cannot find last total. Would you like to enter a starting total?", interface.ButtonSet.OK_CANCEL);
      if (InputLastTotal.getSelectedButton() == interface.Button.CANCEL || InputLastTotal.getSelectedButton() == interface.Button.CLOSE){
        return;
      }
      lastTotal = InputLastTotal.getResponseText();
      while(isNaN(lastTotal)){
        var InputLastTotal = interface.prompt('Sorry, the starting total must be a number.','Please enter a number.' , interface.ButtonSet.OK_CANCEL);
        if (InputLastTotal.getSelectedButton() == interface.Button.CANCEL || InputLastTotal.getSelectedButton() == interface.Button.CLOSE){
          return; 
        }
        lastTotal = InputLastTotal.getResponseText();
      }
    }
    
    //Step input
    var userInputStep = interface.prompt("How much would you like to decrement the values in the column by?",'Please enter a number.' , interface.ButtonSet.OK_CANCEL);
    if (userInputStep.getSelectedButton() == interface.Button.CANCEL || userInputStep.getSelectedButton() == interface.Button.CLOSE){
      return;
    }
    step = userInputStep.getResponseText();
    while(isNaN(step)){
      var userInputStep = interface.prompt('Invalid input. Please enter an number to decrement by.','Please enter a number.' , interface.ButtonSet.OK_CANCEL);
      if (userInputStep.getSelectedButton() == interface.Button.CANCEL || userInputStep.getSelectedButton() == interface.Button.CLOSE){
        return; 
      }
      step = userInputStep.getResponseText();
    }
    
    //Verifying user inputs.
    var confirm = interface.alert('Acting on column '+column+', starting with row '+startRow+', and decreminting by '+step+'. Is this correct?', interface.ButtonSet.YES_NO_CANCEL);
    if (confirm == interface.Button.CLOSE || confirm == interface.Button.CANCEL){
      return;
    }else if (confirm == interface.Button.YES){
      break;
    }
  }
  
  //Getting a value for lastTotal
  if (!lastTotal){
    for (var x = startRow -1; x >= 1; x--){
      lastTotal = spreadsheet.getRange(x,column).getValue();
      if (lastTotal && !isNaN(lastTotal)){
        break;
      }
    }
    if (!lastTotal){
      var InputLastTotal = interface.prompt("Cannot find last total. Would you like to enter a starting total?", interface.ButtonSet.OK_CANCEL);
      if (InputLastTotal.getSelectedButton() == interface.Button.CANCEL || InputLastTotal.getSelectedButton() == interface.Button.CLOSE){
        return;
      }
      lastTotal = InputLastTotal.getResponseText();
      while(isNaN(lastTotal)){
        var InputLastTotal = interface.prompt('Sorry, the starting total must be a number.','Please enter a number.' , interface.ButtonSet.OK_CANCEL);
        if (InputLastTotal.getSelectedButton() == interface.Button.CANCEL || InputLastTotal.getSelectedButton() == interface.Button.CLOSE){
          return; 
        }
        lastTotal = InputLastTotal.getResponseText();
      }
    } 
  }
  
  //Writing to spreadsheet.
  var count = 1;
  for (var i = startRow; i<= lastRow; i++){
    var newTotal = lastTotal - step * count;
    spreadsheet.getRange(i, column).setValue(newTotal);
    count++;
  }
  
}

