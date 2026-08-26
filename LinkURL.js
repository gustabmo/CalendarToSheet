/** 
 * Returns the URL of a hyperlinked cell, if it's entered with control + k. 
 * Author: @Frederico Schardong based on https://support.google.com/docs/thread/28558721?hl=en&msgid=28927581 and https://github.com/andrebradshaw/utilities/blob/master/google_apps/convertHiddenLinks.gs 
 * Supports ranges
 */
function linkURL(reference) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var formula = SpreadsheetApp.getActiveRange().getFormula();
  var args = formula.match(/=\w+\((.*)\)/i);
  try {
    var range = sheet.getRange(args[1]);
  }
  catch(e) {
    throw new Error(args[1] + ' is not a valid range');
  }
  
  var formulas = range.getRichTextValues();
  var output = [];
  for (var i = 0; i < formulas.length; i++) {
    var row = [];
    for (var j = 0; j < formulas[0].length; j++) {
      row.push(formulas[i][j].getLinkUrl());
    }
    output.push(row);
  }
  return output
}