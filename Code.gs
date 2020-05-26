/** Copyright Martin Hawksey 2020
 * Modified by Min-Yen KAN (2020) <knmnyn@comp.nus.edu.sg>
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License.  You may obtain a copy
 * of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  See the
 * License for the specific language governing permissions and limitations under
 * the License.
*/

/**
 * Change these to match the column names you are using for email
 * recepient addresses and email sent column.
*/
var METADATA_SHEET = "Metadata"; // default value
var EMAIL_SHEET = "Form responses 1"; // default value

// GLOBALS - Metadata fields
const STATUS_COL = "Status";
const SCHEDULE_SEND_COL = "Schedule Send";
const MERGE_SHEET_COL = "Merge Sheet Name";
const SEARCH_RESTRICTIONS_COL = "Search Restrictions";
const SEARCH_SUBJECT_LINE_COL = "Search Subject Line";
const TEMPLATE_SUBJECT_LINE_COL = "Template Subject Line";
const SENDER_NAME_COL = "Sender Name";
const REPLY_TO_COL = "Reply To";
const BCC_COL = "BCC";                     // [BUG] doesn't work
const CC_COL = "CC";
const DEBUG_TO_COL = "Debug To";

// GLOBALS - Email/Form fields
const RECIPIENT_EMAIL_ADDRESS_COL  = "Recipient Email Address";
const EMAIL_SENT_COL = "Email Sent";

/**
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Check Quota', 'checkQuota')
      .addItem('Send Emails From Specific Row in Metadata Sheet', 'sendEmailsFromRow')
      .addItem('Send Scheduled Emails', 'sendScheduledEmails')
      .addToUi();
}

/**
 * Send scheduled emails only.
 * Uses overloaded sentEmailsFromMetadata() to run with sendScheduled flag set to 'true'
 *
 * @param {metadataSheet}: sheet to read data from
*/
function sendEmailsFromRow(metadataSheet = SpreadsheetApp.getActive().getSheetByName(METADATA_SHEET)) {
  // Get the row that the user wants to send
  const ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    "Please enter the row number on the Metadata spreadsheet that you want to send",
    ui.ButtonSet.OK_CANCEL);
  sendEmailsFromMetadata(metadataSheet,false, 0 + result.getResponseText());
}

/**
 * Send scheduled emails only.
 * Uses overloaded sentEmailsFromMetadata() to run with sendScheduled flag set to 'true'
 *
 * @param {metadataSheet}: sheet to read data from
*/
function sendScheduledEmails(metadataSheet = SpreadsheetApp.getActive().getSheetByName(METADATA_SHEET)) {
  sendEmailsFromMetadata(metadataSheet,true,-1); // not sending a specific row, but on scheduled mode
}

/**
 * Send emails from sheet data.
 * @param {metadataSheet}: sheet to read data from
 * @param {sendScheduled}: Boolean flag: run to send out scheduled emails
 * @param {rowNum}: which row to send, if not in scheduled mode
*/
function sendEmailsFromMetadata(metadataSheet = SpreadsheetApp.getActive().getSheetByName(METADATA_SHEET), sendScheduled = false, rowNum = -1) {
  // check mode for running, check Boolean sendScheduled
  if (rowNum == -1 && sendScheduled == false) { throw new Error("Must run with either a row number or scheduled mode"); }
  if (sendScheduled == false) { rowNum -= 2; } // remove header rows for row indexing if run in row mode

  // attempt to retrieve metadata
  if (metadataSheet == null) { throw new Error("Can't find metadata sheet '" + METADATA_SHEET + "'. Check your constants."); }

  // get the data from the passed sheet
  const metadataRange = metadataSheet.getDataRange();
  const metadata = metadataRange.getDisplayValues();

  // assuming Row 1 contains our column headings; changes `metadata` var
  const metaheads = metadata.shift();

  metadata.forEach(function(row, rowIdx){
    var status = row[metaheads.indexOf(STATUS_COL)];
    var scheduleSend = row[metaheads.indexOf(SCHEDULE_SEND_COL)];
    var mergeSheet = row[metaheads.indexOf(MERGE_SHEET_COL)];
    var searchRestrictions = row[metaheads.indexOf(SEARCH_RESTRICTIONS_COL)];
    var searchSubjectLine = row[metaheads.indexOf(SEARCH_SUBJECT_LINE_COL)];
    var templateSubjectLine = row[metaheads.indexOf(TEMPLATE_SUBJECT_LINE_COL)];
    var debugTo = "" + row[metaheads.indexOf(DEBUG_TO_COL)];
    var name = row[metaheads.indexOf(SENDER_NAME_COL)];
    var bcc = row[metaheads.indexOf(BCC_COL)];
    var replyTo = row[metaheads.indexOf(REPLY_TO_COL)];
    var cc = row[metaheads.indexOf(CC_COL)];
    var targetedRow = false;

    if (rowNum == rowIdx) { targetedRow = true; } // targeted row if specifically asked to run
    if (sendScheduled == true && scheduleSend != "") {
      try {
        timeScheduled = new Date(scheduleSend).valueOf();
        timeNow = new Date().valueOf();
        if (timeNow > timeScheduled) { targetedRow = true; }
      } catch (e) {
        throw(e);
      }
    }

    if (targetedRow) { // targeted row?
      if (searchSubjectLine != "" && status == "") { // valid line?
        var debug = false;
        if (debugTo != "FALSE" && debugTo != "") { debug = true; }

        // get the Gmail message to use as a template
        var searchString = "subject:\"" + searchSubjectLine + "\"";
        if (searchRestrictions != "") { searchString += " " + searchRestrictions; }

        const emailTemplate = getGmailTemplateFromMail_(searchString);

        // used to record sent emails
        const metadataOut = [];

        var msgOptionsHash = {};
        msgOptionsHash['attachments'] = emailTemplate.attachments;
        if (name != "") { msgOptionsHash['name'] = name; }
        if (cc != "") { msgOptionsHash['cc'] = cc; }
        // if (bcc != "") { msgOptionsHash['bcc'] = bcc; }
        if (replyTo != "") { msgOptionsHash['replyTo'] = replyTo; }

        const returnMsg = sendEmails_(mergeSheet, emailTemplate, templateSubjectLine, debug, debugTo, msgOptionsHash);
        // report stats to user, update the metadatasheet with output
        metadataOut.push(returnMsg);
        metadataSheet.getRange(rowIdx+2, metaheads.indexOf(STATUS_COL)+1,1).setValue(metadataOut);
      } else {
        if (sendScheduled == false) {
          throw new Error("Targeted row has either no subject or has filled status column!");
        }
      }
    }
  });

  /** HELPER FUNCTIONS BELOW ************/

  /**
   * Helper function to send emails according to a merge sheet.
   * @param {mergeSheet}: the name of the worksheet that contains the merge email information.
   * @param {emailTemplate}: the template to use to send the message, with the message data member having "{{}}" fields for merging.
   * @param {templateSubjectLine}: the subject line to use, having "{{}}" fields for merging.
   * @param {debug}: Boolean flag to declare whether debugging or not.
   * @param {debugTo}: email address to send to debugging merged mail output to (if debug == true)
   * @param {msgHash}: partially filled-out messageHash that contains all options that are common to all emails for this function call.
   *
   * @return: overall status string for the email send.
  */
  function sendEmails_(mergeSheetName, emailTemplate, templateSubjectLine, debug, debugTo, msgOptionsHash = {}) {
  // get the data from the address sheet
    sheet = SpreadsheetApp.getActive().getSheetByName(mergeSheetName);
    if (sheet == null) {
      throw new Error("Can't find mail merge sheet '" + mergeSheet + "'. Check your metadata constants.");
    }

    const dataRange = sheet.getDataRange();
    // Fetch displayed values for each row in the Range HT Andrew Roberts
    // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
    // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
    const data = dataRange.getDisplayValues();

    // assuming row 1 contains our column headings
    const heads = data.shift();

    // get the index of column named 'Email Status' (or equivalent as set by constant; assume header names are unique)
    // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
    const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

    // convert 2d array into object array
    // @see https://stackoverflow.com/a/22917499/1027723
    // for pretty version see https://mashe.hawksey.info/?p=17869/#comment-184945
    const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

    // used to record sent emails
    const out = [];
    const metadataOut = [];

    // loop through all the rows of data
    var numSuccess = 0;
    var numTotal = 0;
    var debugMsg = "";
    var debugCount = 0;
    obj.forEach(function(row, rowIdx){
      // only send emails is email_sent cell is blank and not hidden by filter
      if (row[EMAIL_SENT_COL] == ''){
        try {
          var msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
          var subjectString = fillInTemplateFromObject_(templateSubjectLine, row);
          if (templateSubjectLine == "") { subjectString = searchSubjectLine; }
          if (debug) {
            subjectString = "[DEBUGGING] " + subjectString;
            debugMsg = "Debugging Run\n";
          }

          // See documentation for message options: https://developers.google.com/apps-script/reference/mail/mail-app#advanced-parameters_1
          msgOptionsHash['htmlBody'] = msgObj.html;
          msgOptionsHash['attachments'] = emailTemplate.attachments;
          // if (name != "") { msgOptionsHash['name'] = name; }
          // if (cc != "") { msgOptionsHash['cc'] = cc; }
          // if (bcc != "") { msgOptionsHash['bcc'] = bcc; }
          // if (replyTo != "") { msgOptionsHash['replyTo'] = replyTo; }

          // Use MailApp (over GmailApp) that allows sending of Emojis.
          // See https://developers.google.com/apps-script/reference/mail/mail-app
          if (debug) {
            if (debugCount == 0) {
              MailApp.sendEmail(debugTo, subjectString, msgObj.text, msgOptionsHash);
             }
             debugCount = 1;
          } else {
            // GmailApp.sendEmail(row[RECIPIENT_EMAIL_ADDRESS_COL], subjectString, msgObj.text, msgOptionsHash)
            MailApp.sendEmail(row[RECIPIENT_EMAIL_ADDRESS_COL], subjectString, msgObj.text, msgOptionsHash)
          }

          // modify cell to record email sent date
          out.push([debugMsg + new Date()]);
          numSuccess++;
        } catch(e) {
          // modify cell to record error
          out.push([debugMsg + e.message]);
        }
      } else {
        out.push([row[EMAIL_SENT_COL]]);
      }
      numTotal++;
    });

    // updating the mail merge sheet with new data
    sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);

    // return the mailing status to the caller, as an array of size 1
    return ([debugMsg + Date() + "\nSent Emails: " + numSuccess + "\nTotal Lines Seen: " + numTotal]);
  }

  /**
   * Get a Gmail message by matching the subject line.
   * @param {string} subject_line to search for message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromMail_(search_string){
    try {
      // Retrieve all threads of the specified label
      var threads = GmailApp.search(search_string);

      // Retrieve the first message of the first thread
      if (threads.length != 1) { throw new Error("Wrong number (" + threads.length + ") of matching threads, should be just 1 (unique)"); }
      var messages = threads[0].getMessages();
      if (threads.length != 1) { throw new Error("Wrong number (" + messages.length + ") of matching messages, should be just 1 (unique)"); }
      var msg = messages[0];

      const attachments = msg.getAttachments();
      return {message: {subject: msg.getSubject(), text: msg.getPlainBody(), html:msg.getBody()},
              attachments: attachments};
    } catch(e) {
      throw(e);
    }
  }

  /**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
  */
  function fillInTemplateFromObject_(template, data) {
    // we have two templates one for plain text and the html body
    // stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);

    // token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      return data[key.replace(/[{}]+/g, "")] || "";
    });
    return  JSON.parse(template_string);
  }
}

/**
* Populates a dialog box with the information about amount of email allowed to send
*/
function checkQuota() {
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  SpreadsheetApp.getUi().alert("Remaining email quota (refreshes daily): " + emailQuotaRemaining);
}
