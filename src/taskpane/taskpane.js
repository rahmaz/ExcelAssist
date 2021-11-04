/* eslint-disable no-undef */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    //run when ready
    document.getElementById("go").addEventListener("click", executeCommand);
  }
});

export async function executeCommand() {
  const commandVal = document.getElementById("command").value;
  //console.log("commandVal=====", commandVal);
  var parsedCommand = commandVal.split(" ");
  //console.log("parsedCommand=====", parsedCommand);
  var verb = parsedCommand[0];
  //console.log("Verb=====", verb);
  if (verb == "sort") sortCommand(parsedCommand);
  else if (verb == "swap") swapCommand(parsedCommand);
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
    if (parsedCommand[k] == "row" || parsedCommand[k] == "Row") {
      RC = 0;
    }
  }
  //Get Column Name
  if (RC == 1) {
    for (var i = 0; i < commandLength; i++) {
      if (parsedCommand[i] == "column" || parsedCommand[i] == "Column") {
        myColumn = parsedCommand[i + 1];
        //TO DO: handle case where column name is mentioned before the word column
        //TO DO: handle all types of arguments
      }
    }
  } /* Get Row Name */ else {
    for (var r = 0; r < commandLength; r++) {
      if (parsedCommand[r] == "row" || parsedCommand[i] == "Row") {
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
    if (parsedCommand[j].includes("Desc") == true || parsedCommand[j].includes("desc") == true) {
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
    if (parsedCommand[k] == "rows" || parsedCommand[k] == "Rows" || parsedCommand[k] == "row") {
      RC = 0;
    }
  }
  //Get Column Names
  if (RC == 1) {
    for (var i = 0; i < commandLength; i++) {
      if (parsedCommand[i] == "columns" || parsedCommand[i] == "Columns" || parsedCommand[i] == "column") {
        myColumn1 = parsedCommand[i + 1];
        if (parsedCommand[i + 2] == "and") myColumn2 = parsedCommand[i + 3];
        else myColumn2 = parsedCommand[i + 2];
        //TO DO: handle case where first column name is mentioned before the word column
        //TO DO: handle all types of arguments
      }
    }
  } /* Get Row Names */ else {
    for (var r = 0; r < commandLength; r++) {
      if (parsedCommand[r] == "rows" || parsedCommand[i] == "Rows" || parsedCommand[k] == "row") {
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
    columnRange.moveColumnByIndex();
    //either move the columns and store or traditional copy temp
  });
}
