# VBA Macro: Automated Report #1
I used the following VBA code to automate the production of one of my Weekly Reports:
Since this Macro produces a report for the previous week's performance, and since I save the reports in Year folders,I automated the process where the dynamic path (RPATH) I use to save the report is adjusted for last week's year date (i.e: On the first week of January you want last week's report to be saved in the previous year's folder).
Then the code scans Outlook for all emails received on the current week and looks for a specific attachment which is automatically sent by IBM COGNOS via email.
Then it pastes the targeted attachment in an excel template that I use to automate the production of the final Report.
The final report is saved as values into the RPATH and send via email witht he appropriate subject, date and destinataries.
<br>
Note: Due to Privacy, the real Path and Emails have been shortened using ***
