/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console,  Excel, Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function resetSheet(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      const newSheet = context.workbook.worksheets.add();

      activeSheet.load("name");

      await context.sync();

      const sheetName = activeSheet.name;

      activeSheet.delete();
      newSheet.name = sheetName;
      newSheet.activate();

      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("resetSheet", resetSheet);
