/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("table").onclick = table;
  document.getElementById("square").onclick = shape;
  document.getElementById("line").onclick = line;
  document.getElementById("powerpoint").onclick = createPowerPoint;
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "red";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function table() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem("Sheet1");
      let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";

      expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

      expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
      ]);

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

      sheet.activate();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function shape() {
  try {
    // This sample creates a rectangle positioned 100 pixels from the top and left sides
    // of the worksheet and is 150x150 pixels.
    await Excel.run(async (context) => {
      let shapes = context.workbook.worksheets.getItem("Sheet1").shapes;

      let rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
      rectangle.left = 100;
      rectangle.top = 100;
      rectangle.height = 150;
      rectangle.width = 150;
      rectangle.name = "Square";

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function line() {
  try {
    // This sample creates a straight line from [200,50] to [300,150] on the worksheet.
    await Excel.run(async (context) => {
      let shapes = context.workbook.worksheets.getItem("Sheet1").shapes;
      let line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
      line.name = "StraightLine";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function createPowerPoint() {
  try {
    // This sample creates a straight line from [200,50] to [300,150] on the worksheet.
    Office.FileType();
    // powerPoint.isPrototypeOf();
  } catch (error) {
    console.error(error);
  }
}
