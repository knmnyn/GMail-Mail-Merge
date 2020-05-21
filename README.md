# GMail Mail Merge

Yet another Gmail Mail Merge.

This is based on the script

* It uses a separate sheet for metadata to control the merge process.
* It also draws the template email from the user's emails using standard Gmail search queries.  You just have to make sure that
* It adds a debugging variable which should send the email only to a debugging email address.
* It can customize the subject line, feeding the subject line through the mail merge variable replacement process,

BUG: BCC doesn't seem to work currently.
BUG: Scheduled sending and multiple merges in the works.

This can be useful for connecting with Google Form output, where responses come in a separate "Form Responses 1" sheet.
