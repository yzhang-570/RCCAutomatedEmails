// something changed :D

/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 * @see https://developers.google.com/apps-script/guides/triggers#onopene
 */
function onOpen(e) {
  // Add a custom menu to the spreadsheet.
  SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp, or FormApp.
    .createMenu("Custom Menu")
    .addItem("First item", "menuItem1")
    .addItem("Second item", "menuItem2")
    .addToUi();
}

/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
function onEdit(e) {
  // Set a comment on the edited cell to indicate when it was changed.
  const range = e.range;
  range.setNote(`HIHIHI! Last modified: ${new Date()}`);
  Logger.log('range: ' + range);
  Logger.log('range a1notion: ' + range.getA1Notation());
}

/**
 * The event handler triggered when the selection changes in the spreadsheet.
 * @param {Event} e The onSelectionChange event.
 * @see https://developers.google.com/apps-script/guides/triggers#onselectionchangee
 */
// function onSelectionChange(e) {
//   // Set background to red if a single empty cell is selected.
//   const range = e.range;
//   if (
//     range.getNumRows() === 1 &&
//     range.getNumColumns() === 1 &&
//     range.getCell(1, 1).getValue() === ""
//   ) {
//     range.setBackground("red");
//   }
// }