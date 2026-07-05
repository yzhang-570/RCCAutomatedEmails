// something changed :D

// watching changes

// a change
// another change at not 10:11am

// test: taking text input via menu
function showPrompt() {
  // if canceled, then don't show subsequent choices;

  var ui = SpreadsheetApp.getUi();
  // Opens a pop-up box with a text field
  var result = ui.prompt('Input Required', 'Please type your text here:', ui.ButtonSet.OK_CANCEL);
  
  // Process the user's input
  if (result.getSelectedButton() == ui.Button.OK) {
    var textInput = result.getResponseText();
    ui.alert('You entered: ' + textInput);
  }

  // no dropdowns so need to strictly enforce input -> Social, Non-social, Ambassador Meeting, Tabling
  var ui2 = SpreadsheetApp.getUi();

  var result2 = ui.prompt('Input Required',
    'Your Event is a:\nSocial, Non-social, Ambassador Meeting, Tabling', ui.ButtonSet.OK_CANCEL);
    // Process the user's input
  if (result2.getSelectedButton() == ui.Button.OK) {
    var textInput2 = result2.getResponseText();
    ui.alert('Your event type: ' + textInput2); 
  }
}

/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
// function onEdit(e) {
//   // Set a comment on the edited cell to indicate when it was changed.
//   const range = e.range;
//   range.setNote(`HIHIHI! Last modified: ${new Date()}`);
//   Logger.log('range: ' + range);
//   Logger.log('range a1notion: ' + range.getA1Notation());
// }

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