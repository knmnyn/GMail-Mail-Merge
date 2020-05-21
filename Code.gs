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
const METADATA_SHEET = "Metadata";
const EMAIL_SHEET = "Emails";

// Metadata fields
const SEARCH_SUBJECT_LINE_COL = "Search Subject Line";
const TEMPLATE_SUBJECT_LINE_COL = "Template Subject Line";
const SENDER_NAME_COL = "Sender Name";
const REPLY_TO_COL = "Reply To";
const BCC_COL = "BCC";
const CC_COL = "CC";
const DEBUG_TO_COL = "Debug To";
const SEARCH_RESTRICTIONS_COL = "Search Restrictions";

// Email/Form fields
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";

/**
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Check Quota', 'checkQuota')
      .addItem('Send Emails', 'sendEmails')
      .addToUi();
}

/**
 * Send emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
*/
function sendEmails(sheet = SpreadsheetApp.getActive().getSheetByName(EMAIL_SHEET)) {

  // attempt to retrieve metadata
  var metadataSheet = SpreadsheetApp.getActive().getSheetByName(METADATA_SHEET)
  if (metadataSheet == null) {
    throw new Error("Oops - can't find metadata sheet '" + METADATA_SHEET + "'. Check your constants.");
  }

  // get the data from the passed sheet
  const metadataRange = metadataSheet.getDataRange();
  const metadata = metadataRange.getDisplayValues();

  // assuming Row 1 contains our column headings; changes metadata var
  const metaheads = metadata.shift();

  var searchSubjectLine = metadata[0][metaheads.indexOf(SEARCH_SUBJECT_LINE_COL)];
  var templateSubjectLine = metadata[0][metaheads.indexOf(TEMPLATE_SUBJECT_LINE_COL)];
  var debugTo = metadata[0][metaheads.indexOf(DEBUG_TO_COL)];
  var name = metadata[0][metaheads.indexOf(SENDER_NAME_COL)];
  var bcc = metadata[0][metaheads.indexOf(BCC_COL)];
  var replyTo = metadata[0][metaheads.indexOf(REPLY_TO_COL)];
  var cc = metadata[0][metaheads.indexOf(CC_COL)];

  var debug = false;
  if (debugTo != null && debugTo != "FALSE" && debugTo != "") { debug = true; }

  // get the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(searchSubjectLine);

  // get the data from the passed sheet
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

  // loop through all the rows of data
  obj.forEach(function(row, rowIdx){
    // only send emails is email_sent cell is blank and not hidden by filter
    if (row[EMAIL_SENT_COL] == ''){
      try {
        var msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
        var subjectString = fillInTemplateFromObject_(templateSubjectLine, row);
        if (debug) { subjectString = "[DEBUGGING] " + subjectString; }

        // See documentation for message options: https://developers.google.com/apps-script/reference/mail/mail-app#advanced-parameters_1
        var msgOptionsHash = {};
        msgOptionsHash['htmlBody'] = msgObj.html;
        msgOptionsHash['attachments'] = emailTemplate.attachments;
        if (name != null && name != "") { msgOptionsHash['name'] = name; }
        if (cc != null && cc != "") { msgOptionsHash['cc'] = cc; }
        if (bcc != null && bcc != "") { msgOptionsHash['bcc'] = bcc; }
        if (replyTo != null && replyTo != "") { msgOptionsHash['replyTo'] = replyTo; }

        // Use MailApp (over GmailApp) that allows sending of Emojis.
        // See https://developers.google.com/apps-script/reference/mail/mail-app
        // GmailApp.sendEmail(row[RECIPIENT_COL], subjectString, msgObj.text, msgOptionsHash)
        MailApp.sendEmail(row[RECIPIENT_COL], subjectString, msgObj.text, msgOptionsHash)

        // modify cell to record email sent date
        out.push([new Date()]);
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  // updating the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);

  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
    try {
      // get drafts
      const drafts = GmailApp.getDrafts();
      // filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      // get the message object
      const msg = draft.getMessage();
      // getting attachments so they can be included in the merge
      const attachments = msg.getAttachments();
      return {message: {subject: subject_line, text: msg.getPlainBody(), html:msg.getBody()},
              attachments: attachments};
    } catch(e) {
      throw new Error("Oops - can't find Gmail draft");
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
    */
    function subjectFilter_(search_subject_line){
      return function(element) {
        if (element.getMessage().getSubject() === search_subject_line) {
          return element;
        }
      }
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

function checkQuota() {

}
