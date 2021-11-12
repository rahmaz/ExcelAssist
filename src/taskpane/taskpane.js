/* eslint-disable no-undef */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    //run when ready
    document.getElementById("go").addEventListener("click", executeCommand);
  }
});

export async function executeCommand() {
  const commandVal = document.getElementById("command").value; //get the user input text
  const commandLow = commandVal.toLowerCase(); //LowerCase
  var parsedCommand = commandLow.split(" "); //split the commands into words
  var verb = parsedCommand[0]; //get the first word to know which command to execute
  var verbnext = parsedCommand[1]; //if next word is needed ex: move screen
  var verblast = parsedCommand[parsedCommand.length]; //get the last word in case verb is last
  //depending on the verb, execute one of the following functions
  //Group 1
  if (verb == "sort") sortCommand(parsedCommand);
  else if (verb == "swap") swapCommand(parsedCommand);
  //Group 2
  else if (verb == "clear") clearCmd(parsedCommand);
  else if (verb == "delete") deleteCmd(parsedCommand);
  else if (verb == "copy") copyCmd(parsedCommand);
  else if (verb == "highlight") highlightCmd(parsedCommand);
  else if (verb == "move" && (verbnext == "cell" || verbnext == "column" || verbnext == "row" || verbnext == "table"))
    moveCmd(parsedCommand);
  else if (verblast == "bold") boldCmd(parsedCommand);
  else if (verb == "adjust") adjustCmd(parsedCommand);
  //Group 3
  else if (verb == "paste") pasteCmd(parsedCommand);
  //Group 4
  else if (verb == "revert") revertCmd(parsedCommand);
  else if (verb == "save") saveCmd(parsedCommand);
  //Group 5
  else if (verb == "zoom") zoomCmd(parsedCommand);
  else if (verb == "move" && (verbnext == "screen" || verbnext == "sheet")) moveScreenCmd(parsedCommand);
  //Group 6
  else if (verb == "sum" || verb == "add") sumCmd(parsedCommand);
  else if (verb == "average") avgCmd(parsedCommand);
  else if (verblast == "fraction") fractionCmd(parsedCommand);
  else if (verblast == "decimal") decimalCmd(parsedCommand);
  //Group 7
  else if (verb == "hide") hideCmd(parsedCommand);
  else if (verb == "insert") insertCmd(parsedCommand);
  return context.sync();
}

export async function sortCommand(parsedCommand) {
  //TO DO: Handle case where it's not a Table
  console.log("parsedCommand", parsedCommand);
  //var myTable = "test"; //TO DO: get the active table name from the worksheet if it exists, or get table name from command if user mentioned it
  var commandLength = parsedCommand.length;
  var RC = 1; //Row 0 or Column 1
  var myColumn;
  var myRow;
  //checking if it's sort column or sort row, if row is not mentioned it assumes it's column
  for (var k = 0; k < commandLength; k++) {
    if (parsedCommand[k] == "row") {
      RC = 0;
    }
  }
  //Get Column Name
  if (RC == 1) {
    for (var i = 0; i < commandLength; i++) {
      if (parsedCommand[i] == "column") {
        myColumn = parsedCommand[i + 1];
        //TO DO: handle case where column name is mentioned before the word column
        //TO DO: handle all types of arguments
      }
    }
  } /* Get Row Name */ else {
    for (var r = 0; r < commandLength; r++) {
      if (parsedCommand[r] == "row") {
        myRow = parsedCommand[i + 1];
        //TO DO: handle case where row name is mentioned before the word row
        //TO DO: handle all types of arguments
      }
    }
  }
  //Check which order for sorting
  var OrderAsc; //true is ascending
  for (var j = 0; j < commandLength; j++) {
    /* Not sure if includes() will work */
    if (parsedCommand[j].includes("desc") == true) {
      OrderAsc = false;
      //TO DO: handle spelling mistakes
    }
  }
  console.log(myColumn, myRow, OrderAsc);
  //TO DO: Handle row sort
  //Column Sort
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem("test");
    var columnRange = myTable.getDataBodyRange();
    console.log(columnRange);
    columnRange.sort.apply([
      {
        key: myColumn,
        //TO DO: what is myColumn is the name?
        ascending: OrderAsc,
      },
    ]);
  });
  return context.sync();
}

export async function swapCommand(parsedCommand) {
  console.log("parsedCommand", parsedCommand);
  var commandLength = parsedCommand.length;
  //TO DO: Handle case where it's not a Table
  var RC = 1; //Row 0 or Column 1
  var myColumn1;
  var myColumn2;
  var myRow1;
  var myRow2;
  //checking if it's sort column or sort row, if row is not mentioned it assumes it's column
  for (var k = 0; k < commandLength; k++) {
    if (parsedCommand[k] == "rows" || parsedCommand[k] == "row") {
      RC = 0;
    }
  }
  //Get Column Names
  if (RC == 1) {
    for (var i = 0; i < commandLength; i++) {
      if (parsedCommand[i] == "columns" || parsedCommand[i] == "column") {
        myColumn1 = parsedCommand[i + 1];
        if (parsedCommand[i + 2] == "and") myColumn2 = parsedCommand[i + 3];
        else myColumn2 = parsedCommand[i + 2];
        //TO DO: handle case where first column name is mentioned before the word column
        //TO DO: handle all types of arguments
      }
    }
  } /* Get Row Names */ else {
    for (var r = 0; r < commandLength; r++) {
      if (parsedCommand[r] == "rows" || parsedCommand[k] == "row") {
        myRow1 = parsedCommand[i + 1];
        if (parsedCommand[r + 2] == "and") myRow2 = parsedCommand[r + 3];
        else myRow2 = parsedCommand[i + 2];
        //TO DO: handle case where row name is mentioned before the word row
        //TO DO: handle all types of arguments
      }
    }
  }
  //Swap Columns
  //moveColumnByIndex(fromIndex, toIndex)
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem("test");
    var columnRange = myTable.getDataBodyRange();
    //columnRange.moveColumnByIndex(,);
    //either move the columns and store or traditional copy temp
  });
}

export async function clearCmd(parsedCommand) {
  var sheet = context.workbook.worksheets.getActiveWorksheet();
  var range;
  if (parsedCommand[1] == "table") {
    //get table range
    range = sheet.getRange("X");
    range.clear();
  } else if (parsedCommand[1] == "column") {
    //get column range
    range = sheet.getRange("X");
    range.clear();
  } else if (parsedCommand[1] == "row") {
    //get row range
    range = sheet.getRange("X");
    range.clear();
  } else if (parsedCommand[1] == "cell") {
    //get cell range
    range = sheet.getRange("X");
    range.clear();
  } else if (parsedCommand[1] == "sheet") {
    range = sheet.getUsedRange(); //gets the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them
    range.clear();
  }
  return context.sync();
}

export async function deleteCmd(parsedCommand) {
  var sheet = context.workbook.worksheets.getActiveWorksheet();
  var range;
  if (parsedCommand[1] == "table") {
    //get table range
    range = sheet.getRange("X");
    range.clear();
  } else if (parsedCommand[1] == "column") {
    //get column range
    range = sheet.getRange("X");
    range.delete(Excel.DeleteShiftDirection.left);
  } else if (parsedCommand[1] == "row") {
    //get row range
    range = sheet.getRange("X");
    range.delete(Excel.DeleteShiftDirection.up);
  } else if (parsedCommand[1] == "cell") {
    //get cell range
    range = sheet.getRange("X");
    range.clear();
  } else if (parsedCommand[1] == "sheet") {
    //delete the whole worksheet
    sheet.removeWorksheet(sheet.id);
  }
  return context.sync();
}

export async function copyPasteCmd(parsedCommand) {
  var sheet = context.workbook.worksheets.getActiveWorksheet();
  var range;
  var rangePaste;
  if (parsedCommand[1] == "table") {
    //get table range
    range = sheet.getRange("X");
    range.copyFrom(rangePaste);
  } else if (parsedCommand[1] == "column") {
    //get column range
    range = sheet.getRange("X");
    range.copyFrom(rangePaste);
  } else if (parsedCommand[1] == "row") {
    //get row range
    range = sheet.getRange("X");
    range.copyFrom(rangePaste);
  } else if (parsedCommand[1] == "cell") {
    //get cell range
    range = sheet.getRange("X");
    range.copyFrom(rangePaste);
  }
  return context.sync();
}
