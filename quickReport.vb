'**************************************************************************************************************************************''
'Globals & Setup
'**************************************************************************************************************************************''

Dim weekNumber As Integer
Dim totalStudents As Integer

weekNumber = 24
totalStudents = 190

Application.DisplayAlerts = False
Application.ScreenUpdating = False

' Delete extra sheets
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
For r = 2 To 157
  If Worksheets("PLP").Cells(r,3).Value <= 2 Then  ' change this value to change the acceptable range.  2 is within 2 of targer
    Worksheets("PLP").Cells(r,3).Value = "Yes"
  Else
    Worksheets("PLP").Cells(r,3).Value = "No"
  End If
Next r



' while loop to calculate averages
Worksheets("Master").Activate
Dim total As Integer
r = 2

Do While r <= totalStudents
  total = 0

  Do While Cells(r, 1).Value = Cells(r + 1, 1).Value
        total = total + Cells(r, 7).Value
        Cells(r, 9).Value = total
        r = r + 1
        If r > totalStudents Then
          Exit Do
        End If
  Loop

    total = total + Cells(r, 7).Value
    Cells(r, 9).Value = total
    r = r + 1

    If r > totalStudents Then
      Exit Do
    End If
Loop




' while loop to make each row a duplicate, fill in averages
' the last duplicate has the stuff total we want so we need to get
' that value and also count how many duplicates we have
Worksheets("Master").Activate

Dim rr As Integer
Dim clicker As Integer
Dim ave As Single
Dim tempR As Integer
rr = 2

Cells(1, 9).Value = "Total"
Cells(1, 10).Value = "Clicker"

Do While rr < totalStudents
    'set up a clicker to count duplicates.  Start at 1
    clicker = 1

    ' count up duplicate rows and stick in last row
    Do While Cells(rr, 1).Value = Cells(rr + 1, 1).Value
        'error checking to avoid long loop with emtpy cells
        If Cells(rr, 1).Value = "" Then
            Exit Do
        End If
        clicker = clicker + 1
        rr = rr + 1
    Loop

    'put the clicker amount in row 10
    Cells(rr, 10).Value = clicker

    rr = rr + 1
Loop

' calculate the averages and put them in place
Worksheets("Master").Activate
For rr = 2 To totalStudents
    If Cells(rr, 10).Value = "" Then
        Cells(rr, 7).Value = Cells(rr, 9).Value
    Else
        Cells(rr, 7).Value = Cells(rr, 9) / Cells(rr, 10)
    End If
Next rr



' Delete duplicate entries'
Worksheets("Master").Activate
Dim i As Integer
For r = 2 To totalStudents
    For i = 1 To 5
      If Cells(r, 1).Value = Cells(r + 1, 1).Value Then
        Rows(r).EntireRow.Delete
      End If
    Next i
Next r



'**************************************************************************************************************************************''
' Merge data'
'**************************************************************************************************************************************''


' bring over data from Khan Academy in Master to Quick Report'
Worksheets("Quick Report").Activate

' Old algorithm - no error checking'
'Columns("A:Z").Sort key1:=Range("D:D"), order1:=xlAscending, Header:=xlYes
'For r = 2 to totalStudents
' Pull data from last 30 days'
''  Worksheets("Quick Report").Cells(r,1).Value = Worksheets("Master").Cells(r,5).Value
' Pull percent complete data'
''  Worksheets("Quick Report").Cells(r,2).Value = Worksheets("Master").Cells(r,7).Value
'next r
''

' new algorithm which adjusts the index based on if kids are missing in the "Master" sheet
Columns("A:Z").Sort key1:=Range("D:D"), order1:=xlAscending, Header:=xlYes
Dim q As Integer
Dim x As Integer
q = 2

For x = 2 To totalStudents 'row
	Do While Worksheets("Quick Report").Cells(q,4).Value <> Worksheets("Master").Cells(x,1).Value
    q = q + 1
	Loop

	Worksheets("Quick Report").Cells(q,1).Value = Worksheets("Master").Cells(x,6) ' this congtrols which time to bring over -- col 5 is 30 days, col 6 is all time'
  Worksheets("Quick Report").Cells(q,2).Value = Worksheets("Master").Cells(x,7)
	q = q + 1
Next x



' bring over data from PLP into Quick Report'
Worksheets("Quick Report").Activate
Columns("A:Z").Sort key1:=Range("F:F"), order1:=xlAscending, Header:=xlYes

For r = 2 to totalStudents
  Worksheets("Quick Report").Cells(r,3).Value = Worksheets("PLP").Cells(r,3).Value
next r





'**************************************************************************************************************************************''
' formatting
'**************************************************************************************************************************************''


Range("E1").EntireColumn.Insert
For r = 2 To totalStudents
  Cells(r,5).Value = Cells(r,6).Value
next r

'**********************************************************  DEBUGGING ****************************************************************'
Range("F1:H1").EntireColumn.Delete
Range("D1").EntireColumn.Delete
'**************************************************************************************************************************************''

Worksheets("Quick Report").Activate
Range("A:A").NumberFormat = "#"
Range("B:B").NumberFormat = "#"
Range("A1:C1").ColumnWidth = 8
Range("A:Z").Font.Bold = True

Columns("A:Z").Sort key1:=Range("C:C"), order1:=xlDescending, Header:=xlYes

' conditional coloring based on formulas

' Change values in col 1 to be the  last 30 days'
Cells(1,1).Value = "30 Days"

' Deprecated - carry over from when we calculated average time per week'
'For r = 2 To totalStudents
''  Cells(r,1).Value = Cells(r,1).Value / weekNumber
'next r

' color based on if you are within 10% of KA target based on week number'
Dim almost As Integer
almost = weekNumber * 3 - 10
For r = 2 To totalStudents
  If Cells(r,2).Value >= almost Then
    Cells(r,2).Font.Color = RGB(34,139,34)
  Else
    Cells(r,2).Font.Color = RGB(255,0,0)
  End If
next r

' color based on if you are on track with POTW'

For r = 2 to totalStudents
  If Cells(r,3).Value = "Yes" Then
    Cells(r,3).Font.Color = RGB(34,139,34)
  Else
    Cells(r,3).Font.Color = RGB(255,0,0)
  End If
Next r


'Color alternating rows
'******************************************

ActiveSheet.Range("A:Z").Select
Set sh = Worksheets("Quick Report")
Dim flip As Integer
Dim arrs As Integer
arrs = ActiveSheet.UsedRange.Rows.Count


    For flip = 2 To arrs
        'If the row is an odd number (within the selection)...
        If flip Mod 2 = 0 Then
            Cells(flip, 17).Interior.Color = RGB(240, 240, 240)
            Selection.Rows(flip).Interior.Color = RGB(240, 240, 240)
        End If
    Next flip


    'Worksheets("Master").delete
    'Worksheets("PLP").delete


    '************************************************ THIS IS THE END ***********************************************************************'
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    '****************************************************************************************************************************************'
