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

* It uses a separate sheet for metadata to control the merge process.
* It also draws the template email from the user's emails using standard Gmail search queries.  You just have to make sure that
* It adds a debugging variable which should send the email only to a debugging email address.
* It can customize the subject line, feeding the subject line through the mail merge variable replacement process.

This can be useful for connecting with Google Form output, where responses come in a separate "Form Responses 1" sheet.

## Known Bugs
* BCC doesn't seem to work currently.

## Development Plans
* Scheduled sending and multiple merges in the works.
* Needs better installation instructions.
* Aliases for group sending.
