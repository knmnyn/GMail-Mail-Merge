---
title: Yet another mail merge using Gmail and Google Sheets
description: Yet another Gmail Mail Merge, based on the script by Martin Hawksey.  https://github.com/gsuitedevs/solutions/tree/master/mail-merge
labels: Sheets, Gmail
material_icon: merge_type
create_time: 2020-05-22
update_time: 2020-05-22
---

Yet another Gmail Mail Merge, based on the script by Martin Hawksey.
https://github.com/gsuitedevs/solutions/tree/master/mail-merge

* It uses a separate sheet for metadata to control the merge process, so that you can list and reference multiple merges, using the "Merge Sheet Name" variable. Just use this one sheet for all of your merges, and keep historic ones around.  There are two modes of operation:
  * _Send Scheduled Emails_: You can have it send all merge mails that need to be sent after a certain date.  
  * _Run Specific Row Mail Merge_: Run a specific mail merge corresponding to a row of the Metadata sheet.
* The above is also useful for connecting with Google Form output, where responses come in a separate "Form Responses 1" sheet.
* It also draws the template email from the user's emails using standard Gmail search queries.  This allows you to use both sent-mail, mail using a particular label (_e.g._, `label:mail-merge`) or drafts (_e.g.,_ `in:drafts`) all as ways for finding the template to use.
* It adds a debugging variable which should send the email only to a debugging email address, with pop-up alerts for checking the status of the merge process.
* It enables the sending of emails with emoji in both the text and subject lines.
* It can customize the subject line, feeding the subject line through the mail merge variable replacement process.
* Local (per email) merge sheet overrides of values from the metadata sheet, for "CC" (appends), "Reply To" and "Sender Name" columns.

## Known Bugs
* BCC doesn't seem to work currently.

## Development Plans
* Add BCC.
* Debug the prompt box for cancel button.
* Allow Triggers to send Scheduled Emails on cron job.
* Needs better installation instructions.
