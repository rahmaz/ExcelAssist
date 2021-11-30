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

// export async function getTable() {
//   Excel.run(function (context) {
//     var range = context.workbook.getSelectedRange();
//     range.load("address");
//     var tables = context.workbook.tables.load("name");
//     context.sync();
//     var found = false;
//     for (var i = 0; i < tables.items.length; i++) {
//       var table = tables.items[i];
//       var intersectionRange = table.getRange().getIntersection(selection).load("address");
//       context.sync();
//       found = true;
//       console.log(
//         `Intersection found with table "${table.name}". ` + `Intersection range: "${intersectionRange.address}".`
//       );
//     }
//     return context.sync().then(function () {
//       console.log(`The address of the selected range is "${range.address}"`, table);
//     });
//   });
// }

export async function executeCommand() {
  //getTable();
  const commandVal = document.getElementById("command").value; //get the user input text

  //Text processing
  const commandLow = commandVal.toLowerCase(); //Lower Case everything
  var parsedCommand = commandLow.split(" "); //split the command by space

  //determine the keyword, get the first word to know which command to execute
  var verb = parsedCommand[0];
  // var verbnext = parsedCommand[1]; //if next word is needed ex: move screen
  // var verblast = parsedCommand[parsedCommand.length]; //get the last word in case verb is last

  // depending on the verb, execute one of the following functions
  //converts range to table
  if (verb == "convert" || verb == "create") convertcmd(parsedCommand);
  //sorts table using one of the columns, can be called by name letter or number
  else if (verb == "sort") sortCommand(parsedCommand);
  //colors cells, columns, rows
  else if (parsedCommand.includes("fill") || parsedCommand.includes("highlight") || parsedCommand.includes("color"))
    fillcmd(parsedCommand);
  //clears data from the wanted range
  else if (verb == "clear") clearCmd(parsedCommand);
  else if (verb == "swap") swapCommand(parsedCommand);
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

export async function convertcmd(parsedCommand) {
  var myTableName = "test2"; //TO DO: get the active table name from the worksheet
  const commandLength = parsedCommand.length;
  if (parsedCommand.includes("table")) {
    Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        var range = sheet.getUsedRange();
      }
      sheet.activate();
      // Convert the range to a table
      var myTable = sheet.tables.add(range, true);
      for (var i = 1; i < commandLength; i++) {
        if (parsedCommand[i].includes("name")) {
          myTable.name = parsedCommand[i + 1];
          break;
        } else myTable.name = myTableName;
      }
      return context.sync();
    });
  }
}

export async function sortCommand(parsedCommand) {
  var myTableName = "test"; //TO DO: get the active table name from the worksheet
  const commandLength = parsedCommand.length;
  var clmName;
  var headerName;
  //Get Column Name
  for (var i = 1; i < commandLength; i++) {
    if (parsedCommand[i].includes("col")) {
      clmName = parsedCommand[i + 1];
      //convert column Letter or Number to number
      clmName = getColNumber(clmName);
      break;
    } else {
      clmName = await Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var myTable = sheet.tables.getItem(myTableName);
        myTable.columns.load("items");
        return context.sync().then(function () {
          for (var k = 0; k < 100; k++) {
            headerName = String(myTable.columns.items[k].name).toLowerCase();
            for (var j = 1; j < commandLength; j++) {
              if (parsedCommand[j].includes(headerName)) {
                return k;
              }
            }
          }
        });
      });
    }
  }
  console.log("clmName :", clmName);
  //Check order of sorting
  var OrderAsc = true; //true is ascending
  for (var j = 0; j < commandLength; j++) {
    if (parsedCommand[j].includes("des")) {
      OrderAsc = false;
    }
  }
  console.log("Ascending?", OrderAsc);
  //Column Sort
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem(myTableName);
    var columnRange = myTable.getDataBodyRange();
    columnRange.sort.apply([
      {
        key: Number(clmName),
        ascending: OrderAsc,
      },
    ]);
    return context.sync();
  });
}

export async function fillcmd(parsedCommand) {
  var myTableName = "test"; //TO DO: get the active table name from the worksheet
  const commandLength = parsedCommand.length;
  var clmName;
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem(myTableName);
    //Get Column Name, Not handling header names
    for (var i = 1; i < commandLength; i++) {
      if (parsedCommand[i].includes("col")) {
        clmName = parsedCommand[i + 1];
        //convert column Letter or Number to number
        clmName = getColNumber(clmName);
        break;
      }
    }
    if (parsedCommand.includes("transp") || parsedCommand.includes("no")) {
      myTable.getDataBodyRange().format.fill.color = "white";
      myTable.getHeaderRowRange().format.fill.color = "white";
    } else {
      if (parsedCommand[1].includes("head")) {
        myTable.getHeaderRowRange().format.fill.color = parsedCommand[2];
      } else if (parsedCommand[1].includes("table")) {
        myTable.getDataBodyRange().format.fill.color = parsedCommand[2];
        myTable.getHeaderRowRange().format.fill.color = parsedCommand[2];
      } else if (parsedCommand[1].includes("col")) {
        myTable.columns.getItemAt(clmName).getDataBodyRange().format.fill.color = parsedCommand[3];
      } else if (parsedCommand[1].includes("row")) {
        myTable.rows.getItemAt(Number(parsedCommand[2]) - 2).getRange().format.fill.color = parsedCommand[3];
      } else if (parsedCommand[1].includes("cell")) {
        sheet.getRange(parsedCommand[2].toUpperCase()).format.fill.color = parsedCommand[3];
      }
    }
    return context.sync();
  });
}

export async function clearCmd(parsedCommand) {
  var myTableName = "test"; //TO DO: get the active table name from the worksheet
  const commandLength = parsedCommand.length;
  var clmName;
  //Get Column Name, Not handling header names
  for (var i = 1; i < commandLength; i++) {
    if (parsedCommand[i].includes("col")) {
      clmName = parsedCommand[i + 1];
      //convert column Letter or Number to number
      clmName = getColNumber(clmName);
      break;
    }
  }
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem(myTableName);
    var range;
    if (parsedCommand[1].includes("table")) {
      //get table range
      myTable = sheet.tables.getItem(parsedCommand[2]);
      range = myTable.getDataBodyRange();
      range.clear();
    } else if (parsedCommand[1].includes("col")) {
      //get column range
      range = myTable.columns.getItemAt(Number(clmName)).getDataBodyRange();
      range.clear();
    } else if (parsedCommand[1] == "row") {
      //get row range
      range = myTable.rows.getItemAt(Number(parsedCommand[2]) - 2).getRange();
      range.clear();
    } else if (parsedCommand[1] == "cell") {
      //get cell range
      range = sheet.getRange(parsedCommand[2].toUpperCase());
      range.clear();
    } else if (parsedCommand[1] == "sheet") {
      range = sheet.getUsedRange(); //gets the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them
      range.clear();
    }
    return context.sync();
  });
}

export async function swapCommand(parsedCommand) {
  var myTableName = "test"; //TO DO: get the active table name from the worksheet
  const commandLength = parsedCommand.length;
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
  console.log("Row=0 or Col=1?", RC);
  //Get Column Names
  if (RC == 1) {
    for (var i = 1; i < commandLength; i++) {
      if (parsedCommand[i].includes("col")) {
        myColumn1 = parsedCommand[i + 1];
        if (parsedCommand[i + 2] == "and" || parsedCommand[i + 2] == "with") myColumn2 = parsedCommand[i + 3];
        else myColumn2 = parsedCommand[i + 2];
        break;
      } else {
        myColumn1 = parsedCommand[i];
        if (parsedCommand[i + 1] == "and" || parsedCommand[i + 1] == "with") myColumn2 = parsedCommand[i + 2];
        else myColumn2 = parsedCommand[i + 1];
        break;
      }
    }
    console.log("Columns: ", myColumn1, myColumn2);
  } /* Get Row Names */ else {
    for (var r = 1; r < commandLength; r++) {
      if (parsedCommand[r].includes("row")) {
        myRow1 = parsedCommand[r + 1];
        if (parsedCommand[r + 2] == "and" || parsedCommand[r + 2] == "with") myRow2 = parsedCommand[r + 3];
        else myRow2 = parsedCommand[r + 2];
        break;
      }
    }
  }
  console.log("Rows: ", myRow1, myRow2);
  //Swap Columns
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem(myTableName);
    var columnRange = myTable.getDataBodyRange();
    //either move the columns and store or traditional copy temp
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
