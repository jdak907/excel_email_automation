# excel_outlook_automation
Save an Excel sheet based on cell data and email as an Outlook attachment with VBA macros.
This will help save some time with complicated file naming schemes, reduce redundant data entry while reducing errors.


Download the Excel sheet and enable macros.
Customize the quick access toolbar in Excel.
Choose commands from 'Macros', add the Report_Save and Report_Send to the toolbar.

Add data to the fields on the spreadsheet; date, report, serial, part number, company and model.

Click the Report_Save button that has been installed on the quick access toolbar:
This will save the spreadsheet with the filename based on cell data.

Click the Report_Send button that has been installed on the quick access toolbar:
This will create an email with subject and message with a non-macro (.xlsx) version of the spreadsheet attached.

The 'to email' is stored on Z9, and 'to first name' on Z10.
The 'cc email' is stored on Z11 and the 'cc first name' on Z12.
These cells have a white background and white text so they are invisible.

The workbook protection password is 'password'. In the 'EmailReport' form code, Private Sub ButtonSaveSetup_Click() utilizes this password
when saving the email settings to the Z9 through Z12 cells on the worksheet.

Feel free to use this as a starting point for automating some features in your own spreadsheets.
Always remember to lock cells that you don't want to be changed.

-Jeff
