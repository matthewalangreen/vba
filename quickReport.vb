'* Delete extra files
Application.DisplayAlerts = False

Sheets("Time").Delete
Sheets("Sheet1").Name = "Quick Report"
Sheets("3.csv").Name = "PLP"

'* Delete extra columns we don't need
Worksheets("Master").Activate
Range("B1:G1").EntireColumn.Delete
Range("G1:H1").EntireColumn.Delete
Range("H1:J1").EntireColumn.Delete
Worksheets("PLP").Activate
Range("A1").EntireColumn.Delete
Range("C1:L1").EntireColumn.Delete
Range("C1").EntireColumn.Delete

Application.DisplayAlerts = True


'***********************
'* Stuff to do next
'* 1. Calculate "up to date" information within the "Master" sheet
'*    then copy it over to the "Quick Report" sheet.  This will be
'*    tricky because not every kid is in math... maybe we fix this
'*    by removing kids from the contacts spread sheet we use...
'*    it would be best to do error handling within the sheet...
'*
'* 2. Merge data between "Master" & "Quick Report"
'*
'* 3. Calculate "up to date" information for "PLP" this will likely
'*    be based on if there are any missing "Complete" values in
'*
'*
'*
'*
'*
'*
