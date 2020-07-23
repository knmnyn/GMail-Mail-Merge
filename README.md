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

## Try it
Create a copy of the sample [Gmail Mail Merge++](https://docs.google.com/spreadsheets/d/11dS1-kunj-sHA49WVtyACIqmCYOGn3Y5N1lIPPIQZoU/edit?usp=sharing) spreadsheet.

Update the **Recipient Email Address** in the column with email addresses you would like to use in the mail merge in the **Mail Merge** sheet.

Create a message in your Gmail account using markers like `{{First name}}`, which correspond to column names, to indicate text you’d like to be replaced with data from the copied spreadsheet.

In the **Metadata** worksheet, specify the template email (commonly in Drafts, specified by `in:draft` using the **Search Restriction** and the **Search Subject Line**) that you want to use a source for the mass mailing.

In the copied spreadsheet, click on custom menu item `Mail Merge` > `Run Specific Row Mail Merge`.  For example, to run the example mail merge partially completed, type `2`. 

A dialog box will appear and tell you that the script requires authorization. Read the authorization notice and continue.

The Email Sent column will update with the message status, in both the **Metadata** and **Mail Merge** sheets.

## Next steps

Additional columns can be added to the spreadsheet with other data you would like to use. Using the `{{}}` annotation and including your column name as part of your Gmail draft will allow you to include other data from your spreadsheet. If you change the name of the Recipient or Email Sent columns this will need to be updated by opening `Tools` > `Script Editor`.

The source code includes a number of additional parameters, currently commented out, which can be used to control the name of the account email is sent from, reply to email addresses, as well as bcc and cc'd email addresses. 

## Enabling Scheduled Sending

The script also has the function to send scheduled emails, so that you can plan your mass customized mailings in advance.  You need to specific a date/time correctly in the **Schedule Send** column in the **Metadata** worksheet.  When the `sendScheduledEmails` function is run, any row's email whose time is in the past and whose **Status** column is not blank is run.  By default any row that has this column blank is ignored by the Scheduled Send facility.  This facility is similar to the "crontab" means in Un\*x.  You can also trigger this function directly on the `Mail Merge` > `Send Scheduled Email` dropdown.  

There are two steps to get this functionality to run. Let's look at these.

1. You must install an Installable Trigger to the script and specify the `sendScheduledEmails` as the target.  To do this, select `Tools` > `Script Editor` to get to the Script Editor.  In the Script Editor, select `Edit` > `Current Project's Triggers`.  Press the blue `Add Trigger` button in the lower right to add a trigger.  You can set the trigger as you like; for example to check on an hourly basis, follow the set of images below.  

![Step 1.](https://github.com/knmnyn/GMail-Mail-Merge/blob/master/trigger-install-1.png "Go to the current project's triggers in the Script Editor.")
![Step 2.](https://github.com/knmnyn/GMail-Mail-Merge/blob/master/trigger-install-2.png "Specify the interval and logging.")
![Step 3.](https://github.com/knmnyn/GMail-Mail-Merge/blob/master/trigger-install-3.png "Trigger installed!")

2. You must change the `SpreadsheetID` variable on Line .  This is because when the function is run by the trigger it needs to know which worksheet to reference.  The SpreadsheetID is displayed in the URL of the spreadsheet in your browser.  The ID should look something like the information in bold. https://docs.google.com/a/yourdomain.com/spreadsheets/d/**11dS1-kunj-sHA49WVtyACIqmCYOGn3Y5N1lIPPIQZoU**/edit?usp=sharing

With these two steps completed, the **Send Scheduled Email** function will run at the interval you have specified.

## Known Bugs
* BCC doesn't seem to work currently.

## Development Plans
* Add BCC.
* Debug the prompt box for cancel button.
* Needs better installation instructions.
