'Globals
'*****************************************
Dim weekNumber As Integer
Dim minutesPerDay As Integer
Dim totalStudents As Integer

weekNumber = 18
totalStudents = 185
minutesPerDay = 30

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

Application.DisplayAlerts = True              '* move this to the end

'* clean up the PLP sheet and sort it by last name
Worksheets("PLP").Activate
Range("C1").Value = "9"
Range("D1").Value = "10"
Columns("A:Z").Sort key1:=Range("B:B"), order1:=xlAscending, Header:=xlYes
Range("C1").EntireColumn.Insert
Range("C1").Value = "On Track"


'* loop over columns for each row, make sure not empty
Dim r As Integer
Dim c As Integer
Dim lastCol As Integer
Dim pass As Integer

r = 2
c = 4
lastCol = 50
pass = 0

Worksheets("PLP").Cells(r,c).Value = "complete"

For r = 2 To 190 'row
  For c = 4 To lastCol
    If Worksheets("PLP").Cells(r,c).Value = "complete" Then
      Worksheets("PLP").Cells(r,3).Value = "Yes"
    Else If Worksheets("PLP").Cells(r,c).Value <> "complete" Then
      Worksheets("PLP").Cells(r,3).Value = "No"
    End If
  Next c
Next r


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
