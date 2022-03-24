/* global o _init DocumentApp _getFolderFilesIds CacheService GmailApp Logger SpreadsheetApp */
/**
 * Shows messages with the given customer
 * in the Easy to Read tab
 * @param {*} email The customer email
 */
// eslint-disable-next-line no-unused-vars
const showMessages = (email) => {
  if (!email) {
    // user didn't select a customer. abort.
    return;
  }
  // clear the panels from prior conversation
  const cache = CacheService.getDocumentCache();
  cache.remove('threadId');
  cache.remove('threadSubject');
  o.emailTemplateSelect.clearContent();
  o.emailSubject.clearContent();
  o.emailBody.clearContent();
  o.emailMessage.setValue('');
  o.threadSheet.getRange(4, 1, o.threadSheet.getLastRow(), 4).clearContent();
  o.threadSheet.getRange(4, 6, o.threadSheet.getLastRow(), 5).clearContent();
  let messageRows = o.logData
    .filter((row) => [row[o.toCol], row[o.fromCol]].includes(email))
    .sort((a, b) => new Date(a) - new Date(b));

  if (messageRows.length == 0) {
    return;
  }
  // cache the first thread subject so that it can be used
  // to populate the subject field in the email sender
  cache.put('threadId', messageRows[0][o.threadIdCol]);
  cache.put('threadSubject', `Re: ${messageRows[0][o.subjectCol]}`);

  const transcriptRows = messageRows.map((row) => {
    const nr = new Array(4).fill('');
    nr[0] = row[o.dateCol];
    nr[1] = row[o.fromCol];
    nr[2] = row[o.toCol];
    nr[3] = row[o.bodyCol];
    return nr;
  });

  messageRows = messageRows.map((row) => {
    const nr = new Array(10).fill('');
    if (row[o.fromCol] == email) {
      nr[0] = row[o.dateCol];
      nr[1] = row[o.bodyCol];
      // nr[3] = o.clients[email];
    } else {
      const sender = row[o.fromCol];
      nr[5] = row[o.bodyCol];
      nr[8] = row[o.dateCol];
      nr[9] = o.employees[sender] || sender;
    }
    return nr;
  });
  o.threadSheet.getRange(4, 1, messageRows.length, 10).setValues(messageRows);
  o.transcriptSheet
    .getRange(2, 1, o.transcriptSheet.getLastRow(), 4)
    .clearContent();
  o.transcriptSheet
    .getRange(2, 1, transcriptRows.length, 4)
    .setValues(transcriptRows);
  SpreadsheetApp.flush();
};

/**
 * Fetch the templates data from the Google Drive Folder
 * @param {String} templateFolderId The id of the Google Drive folder containing the Google Doc templates
 * @returns an object of the templates
 */
// eslint-disable-next-line no-unused-vars
const _getTemplates = (templateFolderId) => {
  const templates = {};
  const fileIds = _getFolderFilesIds(templateFolderId);
  fileIds.forEach((fileId) => {
    const doc = DocumentApp.openById(fileId);
    const body = doc.getBody();
    const tables = body.getTables();
    const templateName = doc.getName();
    templates[templateName] = {
      subject: tables[0].getText(),
      body: tables[1].getText(),
    };
  });
  return templates;
};

/**
 * Populate subject and body of email client
 * with template data
 */
// eslint-disable-next-line no-unused-vars
const _populateEmailTemplate = (name) => {
  let subject = '';
  let body = '';
  if (name) {
    const cache = CacheService.getDocumentCache();
    const templates = JSON.parse(cache.get('templates'));
    const template = templates[name];
    subject = cache.get('threadSubject') || template.subject;
    body = template.body;
  }
  o.emailSubject.setValue(subject);
  o.emailBody.setValue(body);
};

/**
 * Send an email based on the data in the
 * spreadsheet email client
 */
// eslint-disable-next-line no-unused-vars
const sendEmail = () => {
  _init();
  const body = o.emailBody.getValue();
  const cache = CacheService.getDocumentCache();
  const threadId = cache.get('threadId');
  if (threadId) {
    Logger.log('replied');
    const thread = GmailApp.getThreadById(threadId);
    thread.replyAll(body, { htmlBody: body });
  } else {
    Logger.log('emailed');
    const subject = o.emailSubject.getValue();
    const to = o.clients[o.threadDealSelector.getValue()];
    const ccRowIdx =
      o.clientSheet
        .getDataRange()
        .getValues()
        .findIndex((row) => row[0] == to) + 1;
    const cc = o.clientSheet
      .getRange(ccRowIdx, 7, 1, o.clientSheet.getLastColumn())
      .getValues()
      .flat()
      .filter(Boolean)
      .join(',');
    GmailApp.sendEmail(to, subject, body, { htmlBody: body, cc });
  }
  o.emailTemplateSelect.clearContent();
  o.emailSubject.clearContent();
  o.emailBody.clearContent();
  o.emailMessage.setValue('Email Sent');
};
