const o = {};

/**
 * Intialize the automation
 */
const _init = () => {
  o.ss = SpreadsheetApp.getActiveSpreadsheet();
  o.adminSheet = o.ss.getSheetByName('Admin');
  o.message = o.adminSheet.getRange('B6');
  o.bgDefault = '#efefef';
  o.bgWarning = '#ee9999';
  o.bgSuccess = '#99ee99';
  o.msgStart =
    'Paste a client folder URL above, and select Create Tax Spreadsheet from the Tax Prep menu.';
  o.message.setValue(o.msgStart);
  o.message.setBackground(o.bgDefault);

  o.logHeaders = o.logSheet
    .getRange(1, 1, 1, o.logSheet.getLastColumn())
    .getVAlues()
    .flat()
    .map((x) => x.toLowerCase().trim());
  o.configSheet = o.ss.getSheetByName('Config');
  if (o.configSheet) {
    o.configSheet
      .getDataRange()
      .getValues()
      .forEach((row) => (o[row[0]] = row[1]));
  }
};
