/* global CacheService _getScriptProperty PropertiesService DriveApp Logger _getTemplates showMessages SpreadsheetApp _populateEmailTemplate */
const o = {};

/**
 * Fetch the email templates.
 * Fires via an installable trigger.
 * @param {*} e
 */
// eslint-disable-next-line no-unused-vars
function onTriggeredOpen(e) {
  _init();
  const templateFolderId = _getScriptProperty('templateFolderId');
  try {
    DriveApp.getFolderById(templateFolderId);
  } catch (e) {
    Logger.log(
      `Unable to locate the templates folder with id: ${templateFolderId}`
    );
    return;
  }
  const templates = _getTemplates(templateFolderId);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(templates).sort())
    .setAllowInvalid(false)
    .setHelpText('Select a template name from the dropdown menu')
    .build();
  o.emailTemplateSelect.setDataValidation(rule);
  const cache = CacheService.getDocumentCache();
  cache.put('templates', JSON.stringify(templates), 21600);
}

/**
 * Respond to spreadsheet edits
 * @param {*} e The edit event object
 */
// eslint-disable-next-line no-unused-vars
function onEdit(e) {
  _init();
  const sheet = e.source.getActiveSheet();
  const thisRow = e.range.getRow();
  const thisCol = e.range.getColumn();
  const thisValue = e.range.getValue();
  if (sheet.getName() == o.threadSheetName && thisCol == 4 && thisRow == 1) {
    showMessages(o.clients[thisValue]);
  }
  if (
    sheet.getName() == o.emailSheetName &&
    thisRow == o.emailTemplateSelect.getRow() &&
    thisCol == o.emailTemplateSelect.getColumn()
  ) {
    _populateEmailTemplate(thisValue);
  }
}

/**
 * Get the ids of all files inside a Google Drive folder
 * @param {String} folderId The id of the folder containing the files
 * @returns An array of file IDs
 */
// eslint-disable-next-line no-unused-vars
const _getFolderFilesIds = (folderId) => {
  const fileIds = [];
  const files = DriveApp.searchFiles(`"${folderId}" in parents`);
  while (files.hasNext()) {
    const file = files.next();
    fileIds.push(file.getId());
  }
  return fileIds;
};

/**
 * Get the value of a given key in the Script properties
 * @param {String} k the key to look for in ScriptProperty
 * @returns the value associated with the key
 */
// eslint-disable-next-line no-unused-vars
const _getScriptProperty = (k) => {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty(k);
};

/**
 * Set key and value in ScriptProperties
 * @param {String} k the key to store in ScriptProperties
 * @param {String} v the value to store in ScriptProperties
 */
// eslint-disable-next-line no-unused-vars
const _setScriptProperty = (k, v) => {
  const scriptProperties = PropertiesService.getScriptProperties();
  const obj = {};
  obj[k] = v;
  scriptProperties.setProperties(obj);
};

/**
 * Initialize the automation
 */
// eslint-disable-next-line no-unused-vars
const _init = () => {
  o.ss = SpreadsheetApp.getActiveSpreadsheet();
  o.threadSheetName = 'Easy Read';
  o.threadSheet = o.ss.getSheetByName(o.threadSheetName);
  o.threadDealSelector = o.threadSheet.getRange('D1');
  o.logSheet = o.ss.getSheetByName('Log');
  o.logData = o.logSheet.getDataRange().getValues();
  o.emailSheetName = 'Send Email';
  o.emailSheet = o.ss.getSheetByName(o.emailSheetName);
  o.emailTemplateSelect = o.emailSheet.getRange('C2');
  o.emailSubject = o.emailSheet.getRange('A2');
  o.emailBody = o.emailSheet.getRange('A4');
  o.emailMessage = o.emailSheet.getRange('A14');
  o.transcriptSheetName = 'Transcript';
  o.transcriptSheet = o.ss.getSheetByName(o.transcriptSheetName);
  o.toCol = o.logData[0].indexOf('to');
  o.fromCol = o.logData[0].indexOf('from');
  o.dateCol = o.logData[0].indexOf('date');
  o.bodyCol = o.logData[0].indexOf('body');
  o.threadIdCol = o.logData[0].indexOf('threadId');
  o.subjectCol = o.logData[0].indexOf('subject');
  o.employeeSheetName = 'Employee Email to Name';
  o.employeeSheet = o.ss.getSheetByName(o.employeeSheetName);
  o.employees = {};
  o.employeeSheet
    .getDataRange()
    .getValues()
    .forEach((r) => (o.employees[r[0]] = r[1]));
  o.clientSheetName = 'Email Addresses to Log';
  o.clientSheet = o.ss.getSheetByName(o.clientSheetName);
  o.clients = {};
  o.clientSheet
    .getDataRange()
    .getValues()
    .forEach((r) => (o.clients[r[1]] = r[0]));
};

const dev = () => {
  _setScriptProperty('templateFolderId', '1eT3Kiyh_g9WsZuma7taFNIYOkvSL7aJy');
};
