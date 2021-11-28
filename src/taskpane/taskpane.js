/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

/* eslint-disable no-undef */
/* eslint-disable no-unused-vars */

// var speechRecognition = document.webkitSpeechRecognition;
// var recognition = new speechRecognition();
// var textbox = "#command";
// var instructions = "#instructions";
// var content = "";
// recognition.continuous = true;
// //recognition started
// recognition.onstart = function () {
//   instructions.text("Voice Recognition is On");
// };
// recognition.onspeechend = function () {
//   instructions.text("No Activity");
// };
// recognition.onerror = function () {
//   instructions.text("Try Again");
// };
// recognition.onresult = function (event) {
//   var current = event.resultIndex;
//   var transcript = event.results[current][0].transcript;
//   content += transcript;
//   textbox.val(content);
// };
// "#btn".click(function (event) {
//   recognition.start();
// });
// textbox.on("input", function () {
//   content = this.val();
// });

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    //run when ready
    document.getElementById("go").addEventListener("click", executeCommand);
    document.querySelector("#command").addEventListener("keypress", function (e) {
      if (e.key === "Enter") {
        executeCommand();
      }
    });
  }
});

// Office.context.document.getSelectedDataAsync("table", function (asyncResult) {
//   if (asyncResult.status == "failed") {
//     console.log("Action failed with error: ", asyncResult.error.message);
//   } else {
//     console.log("Headers: " + asyncResult.value.headers + " Rows: " + asyncResult.value.rows);
//   }
// });

function convertLetterToNumber(str) {
  "use strict";
  var out = 0;
  var len = str.length;
  var pos = len;
  while (--pos > -1) {
    out += (str.charCodeAt(pos) - 64) * Math.pow(26, len - 1 - pos);
  }
  out = out - 1;
  console.log("letter to number", str, out);
  return out;
}

function getColNumber(clmName) {
  if (isNaN(+clmName)) {
    //TO DO get column header after getting active table
    clmName = clmName.toUpperCase();
    clmName = convertLetterToNumber(clmName);
  } else clmName = clmName - 1;
  return clmName;
}

export async function executeCommand() {
  const commandVal = document.getElementById("command").value; //get the user input text
  const commandLow = commandVal.toLowerCase(); //LowerCase
  var parsedCommand = commandLow.split(" "); //split the commands into words
  var verb = parsedCommand[0]; //get the first word to know which command to execute
  // var verbnext = parsedCommand[1]; //if next word is needed ex: move screen
  // var verblast = parsedCommand[parsedCommand.length]; //get the last word in case verb is last
  // depending on the verb, execute one of the following functions
  // Group 1
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
}

export async function sortCommand(parsedCommand) {
  //var myTable = "test"; //TO DO: get the active table name from the worksheet if it exists, or get table name from command if user mentioned it
  var commandLength = parsedCommand.length;
  var RC = 1; //Row? 0 or Column? 1
  var clmName;
  var rowName;
  //checking if it's sort column or sort row, if row is not mentioned it assumes it's column
  for (var k = 0; k < commandLength; k++) {
    if (parsedCommand[k] == "row") {
      RC = 0;
    }
  }
  console.log("RC:", RC);
  //Get Column Name
  if (RC == 1) {
    for (var i = 0; i < commandLength; i++) {
      if (parsedCommand[i].includes("col")) {
        clmName = parsedCommand[i + 1];
        //convert column Letter or Number to number
        clmName = getColNumber(clmName);
      }
    }
  } /* Get Row Name */ else {
    for (var r = 0; r < commandLength; r++) {
      if (parsedCommand[r].includes("row")) {
        rowName = parsedCommand[i + 1];
        console.log("rowName:", rowName);
        //TO DO: handle case where row name is mentioned before the word row
        //TO DO: handle all types of arguments
      }
    }
  }
  //Check which order for sorting
  var OrderAsc = true; //true is ascending
  for (var j = 0; j < commandLength; j++) {
    /* Not sure if includes() will work */
    if (parsedCommand[j].includes("des")) {
      OrderAsc = false;
      //TO DO: handle spelling mistakes
    }
  }
  console.log("Ascending?", OrderAsc);
  //TO DO: Handle row sort
  //Column Sort
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem("test");
    var columnRange = myTable.getDataBodyRange();
    columnRange.sort.apply([
      {
        key: Number(clmName), //change this to variable
        //TO DO: what is myColumn is the name?
        ascending: OrderAsc,
      },
    ]);
    return context.sync();
  });
}

export async function swapCommand(parsedCommand) {
  console.log("parsedCommand", parsedCommand);
  var commandLength = parsedCommand.length;
  //TO DO: Handle case where it's not a Table
  var RC = 1; //Row? 0 or Column? 1
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
  console.log("RC:", RC);
  //Get Column Names
  if (RC == 1) {
    for (var i = 0; i < commandLength; i++) {
      if (parsedCommand[i] == "columns" || parsedCommand[i] == "column") {
        myColumn1 = parsedCommand[i + 1];
        if (parsedCommand[i + 2] == "and") myColumn2 = parsedCommand[i + 3];
        else myColumn2 = parsedCommand[i + 2];
        console.log("Columns: ", myColumn1, myColumn2);
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
        console.log("Rows: ", myRow1, myRow2);
        //TO DO: handle case where row name is mentioned before the word row
        //TO DO: handle all types of arguments
      }
    }
  }
  //Swap Columns
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem("test");
    var columnRange = myTable.getDataBodyRange();
    //either move the columns and store or traditional copy temp
    return context.sync();
  });
}

export async function clearCmd(parsedCommand) {
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range;
    if (parsedCommand[1] == "table") {
      //get table range
      var myTable = sheet.tables.getItem(parsedCommand[2]);
      range = myTable.getDataBodyRange();
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
  });
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
