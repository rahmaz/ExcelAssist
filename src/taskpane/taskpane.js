/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

///* eslint-disable no-undef */
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

  //depending on the verb, execute one of the following functions

  //convert range to table or table to range
  if (verb == "convert" || verb == "create") convertcmd(parsedCommand);
  //sort table using one of the columns, can be called by name letter or number
  else if (verb == "sort") sortCommand(parsedCommand);
  //color table, cells, columns, rows
  else if (parsedCommand.includes("fill") || parsedCommand.includes("highlight") || parsedCommand.includes("color"))
    fillcmd(parsedCommand);
  //clear data from a range
  else if (verb == "clear") clearCmd(parsedCommand);
  //swap two columns or two rows
  else if (verb == "swap") swapCommand(parsedCommand);
  //deletes a table, column, or row
  else if (verb == "delete") deleteCmd(parsedCommand);
  //copy and paste cells, columns and rows
  else if ((verb == "copy" || verb == "move") && (parsedCommand.includes("paste") || parsedCommand.includes("to")))
    copyPasteCmd(parsedCommand);
  // adjust the size of columns or rows or tables according the the data
  else if (verb == "adjust") adjustCmd(parsedCommand);
  //hide row or column
  else if (verb == "hide") hideCmd(parsedCommand);
  // else if (verb == "revert") revertCmd(parsedCommand);
  // else if (verb == "save") saveCmd(parsedCommand);
  // else if (verb == "zoom") zoomCmd(parsedCommand);
  // else if (verb == "move" && (verbnext == "screen" || verbnext == "sheet")) moveScreenCmd(parsedCommand);
  // else if (verb == "sum" || verb == "add") sumCmd(parsedCommand);
  // else if (verb == "average") avgCmd(parsedCommand);
  // else if (verblast == "fraction") fractionCmd(parsedCommand);
  // else if (verblast == "decimal") decimalCmd(parsedCommand);
  // else if (verb == "insert") insertCmd(parsedCommand);
}

//DONE
export async function convertcmd(parsedCommand) {
  var myTableName = "test"; //TO DO: get the active table name from the worksheet
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
  var rangeArea;
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
      sheet.getUsedRange().format.fill.color = "white";
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
      } else {
        rangeArea = parsedCommand[1];
        var foundRanges = sheet.findAll(rangeArea, {
          completeMatch: false, // findAll will match the whole cell value
          matchCase: false, // findAll will not match case
        });
        foundRanges.format.fill.color = parsedCommand[2];
      }
    }
    onWorksheetChanged(sheet);
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
    } else {
      range = sheet.getRange(parsedCommand[1].toUpperCase());
      range.clear();
    }
    return context.sync();
  });
}

export async function swapCommand(parsedCommand) {
  var myTableName = "test"; //TO DO: get the active table name from the worksheet
  const commandLength = parsedCommand.length;
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
        if (parsedCommand[i + 2] == "and" || parsedCommand[i + 2] == "with") {
          if (parsedCommand[i + 3].includes("col")) myColumn2 = parsedCommand[i + 4];
          else myColumn2 = parsedCommand[i + 3];
        } else myColumn2 = parsedCommand[i + 2];
        break;
      } else {
        myColumn1 = parsedCommand[i];
        if (parsedCommand[i + 1] == "and" || parsedCommand[i + 1] == "with") myColumn2 = parsedCommand[i + 2];
        else myColumn2 = parsedCommand[i + 1];
        break;
      }
    }
    myColumn1 = getColNumber(myColumn1);
    myColumn2 = getColNumber(myColumn2);
    console.log("Columns: ", myColumn1, myColumn2);
  } /* Get Row Names */ else {
    for (var r = 1; r < commandLength; r++) {
      if (parsedCommand[r].includes("row")) {
        myRow1 = parsedCommand[r + 1];
        if (parsedCommand[r + 2] == "and" || parsedCommand[r + 2] == "with") {
          if (parsedCommand[r + 3].includes("row")) myRow2 = parsedCommand[r + 4];
          else myRow2 = parsedCommand[r + 3];
        } else myRow2 = parsedCommand[r + 2];
        break;
      }
    }
    myRow2 = myRow2 - 2;
    myRow1 = myRow1 - 2;
    console.log("Rows: ", myRow1, myRow2);
  }

  //Columns
  //get range of column 1
  var range1 = await Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem(myTableName);
    var range1 = myTable.columns.getItemAt(Number(myColumn1)).getRange();
    range1.load("address");
    return context.sync().then(function () {
      var range1add = String(range1.address);
      range1add = range1add.split("!");
      range1add = range1add[1];
      return range1add;
    });
  });
  console.log("range 1", range1);
  //get range of column 2
  var range2 = await Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem(myTableName);
    var range2 = myTable.columns.getItemAt(Number(myColumn2)).getRange();
    range2.load("address");
    return context.sync().then(function () {
      var range2add = String(range2.address);
      range2add = range2add.split("!");
      range2add = range2add[1];
      return range2add;
    });
  });
  console.log("range 2", range2);
  var temp = "AAA1:AAA1000";
  //swap columns
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange(range1).moveTo("AAA1");
    sheet.getRange(range2).moveTo(range1);
    sheet.getRange(temp).moveTo(range2);
    return context.sync();
  });

  //Rows
  //get range of row 1
  var rrange1 = await Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem(myTableName);
    var rrange1 = myTable.rows.getItemAt(Number(myRow1)).getRange();
    rrange1.load("address");
    return context.sync().then(function () {
      var range1add = String(rrange1.address);
      range1add = range1add.split("!");
      range1add = range1add[1];
      return range1add;
    });
  });
  console.log("row range 1", rrange1);
  //get range of row 2
  var rrange2 = await Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var myTable = sheet.tables.getItem(myTableName);
    var rrange2 = myTable.rows.getItemAt(Number(myRow2)).getRange();
    rrange2.load("address");
    return context.sync().then(function () {
      var range2add = String(rrange2.address);
      range2add = range2add.split("!");
      range2add = range2add[1];
      return range2add;
    });
  });
  console.log("range 2 row", rrange2);
  var rowtemp = "AAA1000:CCC1000";
  //swap rows
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange(rrange1).moveTo("AAA1000");
    sheet.getRange(rrange2).moveTo(rrange1);
    sheet.getRange(rowtemp).moveTo(rrange2);
    return context.sync();
  });
}

export async function deleteCmd(parsedCommand) {
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
      range = myTable.getRange();
      range.clear();
      myTable.convertToRange();
    } else if (parsedCommand[1].includes("col")) {
      //get column range
      range = myTable.columns.getItemAt(Number(clmName)).getRange();
      myTable = myTable.convertToRange();
      console.log("range", range);
      range.delete(Excel.DeleteShiftDirection.left);
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        var Usedrange = sheet.getUsedRange();
      }
      sheet.activate();
      // Convert back to a table
      myTable = sheet.tables.add(Usedrange, true);
      myTable.name = myTableName;
    } else if (parsedCommand[1] == "row") {
      //get row range
      range = myTable.rows.getItemAt(Number(parsedCommand[2]) - 2).getRange();
      range.delete(Excel.DeleteShiftDirection.up);
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

export async function copyPasteCmd(parsedCommand) {
  const commandLength = parsedCommand.length;
  var myCell;
  var myRange;

  if (parsedCommand[1] == "range" || parsedCommand[1] == "cell") {
    myRange = parsedCommand[2];
  } else myRange = parsedCommand[1];
  myCell = parsedCommand[commandLength - 1];

  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    if (parsedCommand[0] == "move") sheet.getRange(myRange).moveTo(myCell);
    else sheet.getRange(myCell).copyFrom(myRange);
    return context.sync();
  });
}

// This function would be used as an event handler for the Worksheet.onChanged event.
function onWorksheetChanged(eventArgs) {
  Excel.run(function (context) {
    var details = eventArgs.details;
    var address = eventArgs.address;

    // Print the before and after types and values to the console.
    console.log(
      `Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),` +
        +` now is ${details.valueAfter}(${details.valueTypeAfter})`
    );
    return context.sync();
  });
}

export async function adjustCmd(parsedCommand) {
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
      range.format.autofitColumns();
      range.format.autofitRows();
    } else if (parsedCommand[1].includes("col")) {
      //get column range
      range = myTable.columns.getItemAt(Number(clmName)).getDataBodyRange();
      range.format.autofitColumns();
    } else if (parsedCommand[1] == "row") {
      //get row range
      range = myTable.rows.getItemAt(Number(parsedCommand[2]) - 2).getRange();
      range.format.autofitRows();
    } else if (parsedCommand[1] == "cell") {
      //get cell range
      range = sheet.getRange(parsedCommand[2].toUpperCase());
      range.format.autofitColumns();
      range.format.autofitRows();
    } else if (parsedCommand[1] == "sheet") {
      range = sheet.getUsedRange(); //gets the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them
      range.format.autofitColumns();
      range.format.autofitRows();
    } else {
      range = sheet.getRange(parsedCommand[1].toUpperCase());
      range.format.autofitColumns();
      range.format.autofitRows();
    }
    return context.sync();
  });
}

export async function hideCmd(parsedCommand) {
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
    if (parsedCommand[1].includes("col")) {
      //get column range
      range = myTable.columns.getItemAt(Number(clmName)).getRange();
      range.hidden("true");
    } else if (parsedCommand[1] == "row") {
      //get row range
      range = myTable.rows.getItemAt(Number(parsedCommand[2]) - 2).getRange();
      range.hidden("true");
    } else {
      range = sheet.getRange(parsedCommand[1].toUpperCase());
      range.hidden("true");
    }
    return context.sync();
  });
}
