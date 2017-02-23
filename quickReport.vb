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
Range("B1:G1").EntireColumn.Delete 'delete struggling - L2
Range("G1:H1").EntireColumn.Delete 'delete video & skill mins
Range("H1").EntireColumn.Delete 'delete total topics
Range("I1").EntireColumn.Delete 'delete notes'
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


'* replace complete, working and off track with corresponding values 1 & 0
Worksheets("PLP").Columns("A:Z").Replace _
 What:="complete", Replacement:=0, _
 SearchOrder:=xlByColumns, MatchCase:=True

Worksheets("PLP").Columns("A:Z").Replace _
 What:="working", Replacement:=1, _
 SearchOrder:=xlByColumns, MatchCase:=True

 Worksheets("PLP").Columns("A:Z").Replace _
 What:="off track", Replacement:=1, _
 SearchOrder:=xlByColumns, MatchCase:=True

'* loop over columns for each row, make sure not empty
Dim r As Integer
For r = 2 To totalStudents 'row
  Worksheets("PLP").Cells(r, 3).Value = Application.sum(Range(Cells(r, 4), Cells(r, 30)))
Next r

'* if the value is 0, replace it with a "YES", otherwise replace with "NO"
For r = 2 To totalStudents
  If Worksheets("PLP").Cells(r,3).Value = 0 Then
    Worksheets("PLP").Cells(r,3).Value = "Yes"
  Else
    Worksheets("PLP").Cells(r,3).Value = "No"
  End If
Next r


' while loop to calculate averages
Dim row As Integer
Dim total As Integer
row = 2

Do While row <= totalStudents
  total = 0

  Do While Worksheets("Master").Cells(row,1).Value = Worksheets("Master").Cells(row+1,1).Value
    total = total + Worksheets("Master").Cells(row,7).Value
    Worksheets("Master").Cells(row,9).Value = total
    row = row + 1
  Loop

    total = total + Worksheets("Master").Cells(row, 7).Value
    Worksheets("Master").Cells(row, 9).Value = total
    row = row + 1
Loop



' while loop to make each row a duplicate, fill in averages
' the last duplicate has the stuff total we want so we need to get
' that value and also count how many duplicates we have
' we need to delete all but the last one.
Dim count As Integer
Dim t As Integer
Dim ave As Single
Dim tempR As Integer

ave = 0
r = 2
count = 1
t = 0

Do While r <= 190
  count = 1

  ' go through and count up duplicates
  Do While Worksheets("Master").Cells(r, 1).Value = Worksheets("Master").Cells(r + 1, 1).Value
    count = count + 1
    r = r + 1
  Loop

  ' calculate the average and stick it in the last row of duplicates
  t = Worksheets("Master").Cells(r, 9).Value
  ave = t / count
  Worksheets("Master").Cells(r, 7).Value = ave

  'Debugging
  Worksheets("Master").Cells(r, 11).Value = count
  Worksheets("Master").Cells(r, 12).Value = r

  ' Go through and copy total percentage calculated from Cells(r,9) and total
  ' this works by counting from the last duplicate value back up
  ' we will just delete duplicagtes at the end of this outter while loop'
  tempR = r
  Do While count >= 1
    Worksheets("Master").Cells(tempR, 7).Value = ave
    count = count - 1
    tempR = tempR - 1
    'Worksheets("Master").Cells(tempR, 1).EntireRow.Delete
  Loop


  Worksheets("Master").Cells(r, 10).Value = count
  r = r + 1
Loop


'****************************************************************************************************************************************'

' Delete duplicate entries'
r = 2
Do While r < 190
  Do While Worksheets("Master").Cells(r,1).Value = Worksheets("Master").Cells(r+1,1).Value
    Worksheets("Master").Cells(r+1,1).EntireRow.Delete
    r = r + 1
  Loop
  r = r + 1
Loop




'***********************
'* Stuff to do next
'* 1. Calculate "up to date" information within the "Master" sheet
'*    then copy it over to the "Quick Report" sheet.  This will be
'*    tricky because not every kid is in math... maybe we fix this
'*    by removing kids from the contacts spread sheet we use...
'*    it would be best to do error handling within the sheet...

'* 2.  Write a function to average the a2 & trig scores'
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
