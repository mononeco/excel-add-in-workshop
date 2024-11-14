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
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange();
      range.format.rowHeight = 8;
      range.format.columnWidth = 8;
      await context.sync();

      range.conditionalFormats.clearAll();
      await context.sync();

      const rule = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);

      rule.colorScale.criteria = {
        minimum: {
          formula: null,
          type: Excel.ConditionalFormatColorCriterionType.lowestValue,
          color: "white",
        },
        midpoint: {
          formula: "50",
          type: Excel.ConditionalFormatColorCriterionType.percent,
          color: "gray",
        },
        maximum: {
          formula: null,
          type: Excel.ConditionalFormatColorCriterionType.highestValue,
          color: "black",
        },
      };

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
