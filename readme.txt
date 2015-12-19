The purpose of this program is to automate uploading of standardized test scores into a database at the Berea College Education Studies Department. When I first began my job at Berea College, test scores were entered by hand which was extremely time consuming.  This program automates this entire task by doing the following:

-Obtaining score reports Educational Testing Service (ETS) via a SOAP api.
-Determine which score reports are new (if any).
-Uploads the scores into the database.
-Emails the appropriate stakeholders (student test taker, student's advisor and department chair) if any tests were not passing

#Background Information
When I first began this position, Microsoft Access was the only database solution available. This program in its present form assumes a Microsoft Access backend. As of December  2015, we have been given permission to develop a more sophisticated web application using Ruby on Rails and a MySQL backend. The code for that application will be shared when it is available.  The update code for this program made to run on our new system will also be shared.

Throughout the codebase, names like bnum and bnums are used. This is short for B Number which is another name for student ID number at Berea College.