
' Notes
' Summary is 30 days                  --YES
' Summary (2) is yesterday            --YES
' Summary (3) is last 2 days`         --YES
' Summary (4) is Trigonometry         --YES
' Summary (5) is Statistics           --YES
' Summary (6) is PreCalculus          --YES
' Summary (7) is PreAlgebra           --YES
' Summary (8) is Geometry             --YES
' Summary (9) is Calc B               --YES
' Summary (10) is Calc A              --YES
' Summary (11) is A2                  --YES
' Summary (12) is Algebra 1           --YES
' Summary (13) is ALL time            --YES

' ISSUES TO FIX'
' 1.  Make another sheet besides summmary 9 to populate with info... We're gonna
'     Have calculus B students with scores in there pretty soon.'
' 2.  Determine if the reason yesterday & today match is because I'm using the
'     Wrong sheet.  (I have them mislabeled)'

'Globals
'*****************************************
Dim weekNumber As Integer
Dim minutesPerDay As Integer
Dim totalStudents As Integer

weekNumber = 15
totalStudents = 176
minutesPerDay = 30

'Pull values from mission specific into summary
'*************************************************
Sheets("Summary").Range("D:G").Value = Sheets("Mission-specific").Range("D:G").Value
Sheets("Summary (2)").Range("D:G").Value = Sheets("Mission-specific (2)").Range("D:G").Value
Sheets("Summary (3)").Range("D:G").Value = Sheets("Mission-specific (3)").Range("D:G").Value
Sheets("Summary (4)").Range("D:G").Value = Sheets("Mission-specific (4)").Range("D:G").Value
Sheets("Summary (5)").Range("D:G").Value = Sheets("Mission-specific (5)").Range("D:G").Value
Sheets("Summary (6)").Range("D:G").Value = Sheets("Mission-specific (6)").Range("D:G").Value
Sheets("Summary (7)").Range("D:G").Value = Sheets("Mission-specific (7)").Range("D:G").Value
Sheets("Summary (8)").Range("D:G").Value = Sheets("Mission-specific (8)").Range("D:G").Value
Sheets("Summary (13)").Range("D:G").Value = Sheets("Mission-specific (13)").Range("D:G").Value


' Do the extra time calculations and get them into a single sheet
'************************************************************************
' get data from summary (9) - summary (12) into a single sheet by pulling data into
' specific and labeled columns


Worksheets("Summary (9)").Activate
Columns("B:G").EntireColumn.Delete
Columns("C:H").EntireColumn.Delete


Sheets("Summary (9)").Range("C:C").Value = Sheets("Summary (10)").Range("H:H").Value
Sheets("Summary (9)").Range("D:D").Value = Sheets("Summary (11)").Range("H:H").Value
Sheets("Summary (9)").Range("E:E").Value = Sheets("Summary (12)").Range("H:H").Value

Range("B1").Value = "Yesterday"
Range("C1").Value = "2 Days"
Range("D1").Value = "7 Days"
Range("E1").Value = "30 Days"

Worksheets("Summary (9)").Columns.AutoFit

Worksheets("Summary (9)").Name = "Time"

Worksheets("Time").Columns("A:E").Sort key1:=Range("A:A"), order1:=xlAscending, Header:=xlYes

Worksheets("Time").Move _
 after:=Worksheets("Summary (5)")


Application.DisplayAlerts = False
Sheets("Summary (10)").Delete
Sheets("Summary (11)").Delete
Sheets("Summary (12)").Delete
Application.DisplayAlerts = True




'Delete extra sheets''
'******************************************

Application.DisplayAlerts = False
Sheets("Mission-specific").Delete
Sheets("Mission-specific (2)").Delete
Sheets("Mission-specific (3)").Delete
Sheets("Mission-specific (4)").Delete
Sheets("Mission-specific (5)").Delete
Sheets("Mission-specific (6)").Delete
Sheets("Mission-specific (7)").Delete
Sheets("Mission-specific (8)").Delete
Sheets("Mission-specific (9)").Delete
Sheets("Mission-specific (10)").Delete
Sheets("Mission-specific (11)").Delete
Sheets("Mission-specific (12)").Delete
Sheets("Mission-specific (13)").Delete

Sheets("Videos").Delete
Sheets("Videos (2)").Delete
Sheets("Videos (3)").Delete
Sheets("Videos (4)").Delete
Sheets("Videos (5)").Delete
Sheets("Videos (6)").Delete
Sheets("Videos (7)").Delete
Sheets("Videos (8)").Delete
Sheets("Videos (9)").Delete
Sheets("Videos (10)").Delete
Sheets("Videos (11)").Delete
Sheets("Videos (12)").Delete
Sheets("Videos (13)").Delete

Sheets("Badges").Delete
Sheets("Badges (2)").Delete
Sheets("Badges (3)").Delete
Sheets("Badges (4)").Delete
Sheets("Badges (5)").Delete
Sheets("Badges (6)").Delete
Sheets("Badges (7)").Delete
Sheets("Badges (8)").Delete
Sheets("Badges (9)").Delete
Sheets("Badges (10)").Delete
Sheets("Badges (11)").Delete
Sheets("Badges (12)").Delete
Sheets("Badges (13)").Delete

Sheets("Points").Delete
Sheets("Points (2)").Delete
Sheets("Points (3)").Delete
Sheets("Points (4)").Delete
Sheets("Points (5)").Delete
Sheets("Points (6)").Delete
Sheets("Points (7)").Delete
Sheets("Points (8)").Delete
Sheets("Points (9)").Delete
Sheets("Points (10)").Delete
Sheets("Points (11)").Delete
Sheets("Points (12)").Delete
Sheets("Points (13)").Delete

Sheets("Exercises").Delete
Sheets("Exercises (2)").Delete
Sheets("Exercises (3)").Delete
Sheets("Exercises (4)").Delete
Sheets("Exercises (5)").Delete
Sheets("Exercises (6)").Delete
Sheets("Exercises (7)").Delete
Sheets("Exercises (8)").Delete
Sheets("Exercises (9)").Delete
Sheets("Exercises (10)").Delete
Sheets("Exercises (11)").Delete
Sheets("Exercises (12)").Delete
Sheets("Exercises (13)").Delete



Application.DisplayAlerts = True


'algorithm to replace column values in each student sheet
'*************************************************************

Dim sh As Worksheet
Dim rw As Range

'Algebra 2
'*************************************************************
Set sh = Worksheets("Summary")
For Each rw In sh.rows
    If sh.Cells(rw.Row, 13).Value = "" Then
        Exit For
    End If
    If sh.Cells(rw.Row, 13).Value = "Classes" Then
    Else
        sh.Cells(rw.Row, 13).Value = "Algebra 2"
    End If
Next rw

'Calculus A
'*************************************************************
Set sh = Worksheets("Summary (2)")
For Each rw In sh.rows
    If sh.Cells(rw.Row, 13).Value = "" Then
        Exit For
    End If
    If sh.Cells(rw.Row, 13).Value = "Classes" Then
    Else
        sh.Cells(rw.Row, 13).Value = "Calculus A"
    End If
Next rw

'Calculus B
'*************************************************************
Set sh = Worksheets("Summary (3)")
For Each rw In sh.rows
    If sh.Cells(rw.Row, 13).Value = "" Then
        Exit For
    End If
    If sh.Cells(rw.Row, 13).Value = "Classes" Then
    Else
        sh.Cells(rw.Row, 13).Value = "Calculus B"
    End If
Next rw

'Geometry
'*************************************************************
Set sh = Worksheets("Summary (4)")
For Each rw In sh.rows
    If sh.Cells(rw.Row, 13).Value = "" Then
        Exit For
    End If
    If sh.Cells(rw.Row, 13).Value = "Classes" Then
    Else
        sh.Cells(rw.Row, 13).Value = "Geometry"
    End If
Next rw

'Pre Algebra
'*************************************************************
Set sh = Worksheets("Summary (5)")
For Each rw In sh.rows
    If sh.Cells(rw.Row, 13).Value = "" Then
        Exit For
    End If
    If sh.Cells(rw.Row, 13).Value = "Classes" Then
    Else
        sh.Cells(rw.Row, 13).Value = "Pre Algebra"
    End If
Next rw

'Pre Calculus
'*************************************************************
Set sh = Worksheets("Summary (6)")
For Each rw In sh.rows
    If sh.Cells(rw.Row, 13).Value = "" Then
        Exit For
    End If
    If sh.Cells(rw.Row, 13).Value = "Classes" Then
    Else
        sh.Cells(rw.Row, 13).Value = "Pre Calculus"
    End If
Next rw

'Statistics
'*************************************************************
Set sh = Worksheets("Summary (7)")
For Each rw In sh.rows
    If sh.Cells(rw.Row, 13).Value = "" Then
        Exit For
    End If
    If sh.Cells(rw.Row, 13).Value = "Classes" Then
    Else
        sh.Cells(rw.Row, 13).Value = "Statistics"
    End If
Next rw

' Trigonometry
'*************************************************************
Set sh = Worksheets("Summary (8)")
For Each rw In sh.rows
    If sh.Cells(rw.Row, 13).Value = "" Then
        Exit For
    End If
    If sh.Cells(rw.Row, 13).Value = "Classes" Then
    Else
        sh.Cells(rw.Row, 13).Value = "Trigonometry"
    End If
Next rw

'Algebra 1
'*************************************************************
Set sh = Worksheets("Summary (13)")
For Each rw In sh.rows
    If sh.Cells(rw.Row, 13).Value = "" Then
        Exit For
    End If
    If sh.Cells(rw.Row, 13).Value = "Classes" Then
    Else
        sh.Cells(rw.Row, 13).Value = "Algebra 1"
    End If
Next rw


' Delete the empty calculus B class
'************************************
Application.DisplayAlerts = False
Sheets("Summary (3)").Delete
Application.DisplayAlerts = True



'Combine into single sheet
'******************************************

    Dim wrk As Workbook 'Workbook object - Always good to work with object variables
    Dim sht As Worksheet 'Object for handling worksheets in loop
    Dim trg As Worksheet 'Master Worksheet
    Dim rng As Range 'Range object
    Dim colCount As Integer 'Column count in tables in the worksheets

    Set wrk = ActiveWorkbook 'Working in active workbook

    For Each sht In wrk.Worksheets
        If sht.Name = "Master" Then
            MsgBox "There is a worksheet called as 'Master'." & vbCrLf & _
            "Please remove or rename this worksheet since 'Master' would be" & _
            "the name of the result worksheet of this process.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
        End If
    Next sht

     'We don't want screen updating
    Application.ScreenUpdating = False

     'Add new worksheet as the last worksheet
    Set trg = wrk.Worksheets.Add(After:=wrk.Worksheets(wrk.Worksheets.count))
     'Rename the new worksheet
    trg.Name = "Master"
     'Get column headers from the first worksheet
     'Column count first
    Set sht = wrk.Worksheets(1)
    colCount = sht.Cells(1, 255).End(xlToLeft).Column
     'Now retrieve headers, no copy&paste needed
    With trg.Cells(1, 1).Resize(1, colCount)
        .Value = sht.Cells(1, 1).Resize(1, colCount).Value
         'Set font as bold
        .Font.Bold = True
    End With

     'We can start loop
    For Each sht In wrk.Worksheets
         'If worksheet in loop is the last one, stop execution (it is Master worksheet)

       		If sht.Index = wrk.Worksheets.count Then
            	Exit For
            '**** I'm messing with this ****
            ElseIf sht.Name <> "Time" Then


         	'Data range in worksheet - starts from second row as first rows are the header rows in all worksheets
       	 	Set rng = sht.Range(sht.Cells(2, 1), sht.Cells(65536, 1).End(xlUp).Resize(, colCount))
         	'Put data into the Master worksheet
        	trg.Cells(65536, 1).End(xlUp).Offset(1).Resize(rng.rows.count, rng.Columns.count).Value = rng.Value
        	End If

    Next sht
     'Fit the columns in Master worksheet
    trg.Columns.AutoFit

     'Screen updating should be activated
    Application.ScreenUpdating = True


    Application.DisplayAlerts = False
    Sheets("Summary").Delete
    Sheets("Summary (2)").Delete
    ' this is BC Calculus'
    'Sheets("Summary (3)").Delete
    Sheets("Summary (4)").Delete
    Sheets("Summary (5)").Delete
    Sheets("Summary (6)").Delete
    Sheets("Summary (7)").Delete
    Sheets("Summary (8)").Delete
    Sheets("Summary (13)").Delete



    Application.DisplayAlerts = True







'Rotate label text
'******************************************
rows(1).Orientation = xlUpward
Range("A1").Orientation = xlHorizontal
Range("M1").Orientation = xlHorizontal
Range("N1").Orientation = xlHorizontal
Range("Q1").Orientation = xlHorizontal

'Add columns for values
'******************************************
Range("O1:O170").Value = " "
Range("L1").EntireColumn.Insert
Range("L1").Value = "% Complete"
Range("M1").EntireColumn.Insert
Range("M1").Value = "Total Topics"

'Format columns
'******************************************
  Range("P1").Value = "Student Profile"
  Range("Q1").Value = "Notes"
  Range("P1").Hyperlinks.Delete
  Range("P1").Font.Bold = True
  Range("Q1").Font.Bold = True
  Range("H:J").NumberFormat = "#,###.#"
  Range("L:L").NumberFormat = "###"
  Range("Q1").Orientation = xlHorizontal

'Set column widths
'******************************************

Range("B:G").ColumnWidth = 5
Range("H:J").ColumnWidth = 8
Range("K:M").ColumnWidth = 5
Range("N:N").ColumnWidth = 8
Range("O:O").ColumnWidth = 15
Range("P:P").ColumnWidth = 15
Range("Q:Q").ColumnWidth = 40

'Remove duplicates
'******************************************
Sheets("Master").Range("A:P").RemoveDuplicates Columns:=Array(1, 16), Header:=xlYes

'calculate total topics
'*****************************
Dim l As Integer 'row

For l = 2 To totalStudents
    If Cells(l, 15).Value = "Algebra 1" Then
       Cells(l, 13).Value = 185
    ElseIf Cells(l, 15).Value = "Algebra 2" Then
       Cells(l, 13).Value = 141
    ElseIf Cells(l, 15).Value = "Geometry" Then
        Cells(l, 13).Value = 104
    ElseIf Cells(l, 15).Value = "Pre Algebra" Then
       Cells(l, 13).Value = 182
    ElseIf Cells(l, 15).Value = "Calculus A" Then
       Cells(l, 13).Value = 87
    ElseIf Cells(l, 15).Value = "Pre Calculus" Then
       Cells(l, 13).Value = 106
    ElseIf Cells(l, 15).Value = "Calculus B" Then
       Cells(l, 13).Value = 65
    ElseIf Cells(l, 15).Value = "Statistics" Then
       Cells(l, 13).Value = 52
    ElseIf Cells(l, 15).Value = "Trigonometry" Then
       Cells(l, 13).Value = 35
    Else
       Cells(l, 13).Value = 1
     End If
Next l

'calculate percent complete
'*******************************
Dim r As Integer 'row
Dim mastered As Integer
Dim l2 As Integer
Dim l1 As Integer
Dim practiced As Integer
Dim total As Integer
Dim result As Integer

    For r = 2 To totalStudents
        mastered = Cells(r, 7).Value
        l2 = Cells(r, 6).Value
        l1 = Cells(r, 5).Value
        practiced = Cells(r, 4).Value
        total = Cells(r, 13).Value

        result = ((4 * mastered + 3 * l2 + 2 * l1 + practiced) / (4 * total)) * 100
        Cells(r, 12).Value = result

Next r

'Sort
'********************************

Columns("A:P").Sort key1:=Range("L:L"), order1:=xlAscending, Header:=xlYes

Columns("A:P").Sort key1:=Range("O:O"), order1:=xlAscending, Header:=xlYes

Columns("A:P").Sort key1:=Range("A:A"), order1:=xlAscending, Header:=xlYes


'Color alternating rows
'******************************************

ActiveSheet.Range("A:Q").Select
Set sh = Worksheets("Master")
Dim Counter As Integer
Dim arrs As Integer
arrs = ActiveSheet.UsedRange.rows.count


    For Counter = 2 To arrs
        'If the row is an odd number (within the selection)...
        If Counter Mod 2 = 0 Then
        	Cells(Counter,17).Interior.Color = RGB(240,240,240)
            Selection.rows(Counter).Interior.Color = RGB(240, 240, 240)
        End If
    Next



'Label total time based on minute of practice average
'******************************************
Dim i As Integer 'row
Dim studentPractice As Integer
Dim fiveHoursEachWeek As Integer
Dim threeHoursEachWeek As Integer
Dim minimumPractice As Integer
'minutesPerDay = 30


minimumPractice = minutesPerDay * 5 * weekNumber


For i = 2 To totalStudents 'iterate over rows
	studentPractice = Cells(i,8).Value

    If studentPractice >= minimumPractice Then
    	Cells(i,8).Font.Color = RGB(34,139,34) 'Mark green

    Else
    	Cells(i,8).Font.Color = RGB(255,0,0) 'Otherwise mark red

    End If

Next i

'Label names as behind percent complete or not
'******************************************
Dim x As Integer 'row
Dim comp As Integer 'how complete are you?
comp = weekNumber * 3

For x = 2 To totalStudents 'row
	'Cells(x,17).Value = "a thing"
    If Cells(x, 12).Value >= comp Then
            Cells(x, 12).Font.Color = RGB(34,139,34) 'green
            Cells(x, 1).Font.Color = RGB(34,139,34) 'green
       ElseIf Cells(x, 12).Value >= comp - 3 Then
            Cells(x, 12).Font.Color = RGB(255, 128, 0) 'orange
            Cells(x, 1).Font.Color = RGB(255, 128, 0) 'orange
        Else
            Cells(x, 12).Font.Color = RGB(255,0,0)
            Cells(x, 1).Font.Color = RGB(255,0,0)
        End If
Next x

'Migrate data from "Time" Sheet into main
'sheet before additional sorting, calculations and formatting
'******************************************


' make new columns
'******************************************
Range("H1").EntireColumn.Insert
Range("H1").EntireColumn.Insert
Range("H1").EntireColumn.Insert
Range("H1").EntireColumn.Insert
Range("H1").Value = "Yesterday"
Range("I1").Value = "2 Days"
Range("J1").Value = "7 Days"
Range("K1").Value = "30 Days"
Range("L1").Value = "All Time"


' Delete Badges, Points & Profile
'******************************************
Range("O1").EntireColumn.Delete
Range("Q1").EntireColumn.Delete
Range("R1").EntireColumn.Delete


' Migrate date data from other sheet
'******************************************

Dim clicker As Integer
clicker = 2

For x = 2 To totalStudents 'row
	Do While Worksheets("Master").Cells(x,1).Value <> Worksheets("Time").Cells(clicker,1).Value
  		clicker = clicker - 1
	Loop

	Worksheets("Master").Cells(x,8).Value = Worksheets("Time").Cells(clicker,2).Value
	Worksheets("Master").Cells(x,9).Value = Worksheets("Time").Cells(clicker,3).Value
	Worksheets("Master").Cells(x,10).Value = Worksheets("Time").Cells(clicker,4).Value
	Worksheets("Master").Cells(x,11).Value = Worksheets("Time").Cells(clicker,5).Value

	clicker = clicker + 1
Next x


' More formatting of column widths
'******************************************

Range("H:N").NumberFormat = "#####"
Range("H:N").ColumnWidth = 6


'  Add today's date to top right column of sheet
'******************************************

Dim str As String
str = Range("R1").Value
'Format (Now, “dd/MM/yyyy”)
str = str + " " + Format(Now, "MM/dd/yy")
Range("R1").Value = str

'Label time columns based on practice
'******************************************
Dim yDay As Integer
Dim twoDay As Integer
Dim sevenDay As Integer
Dim thirtyDay As Integer

yDay = minutesPerDay
twoDay = minutesPerDay * 2
sevenDay = minutesPerDay * 5
thirtyDay = minutesPerDay * 22

'Label yesterday
'********************************************
For i = 2 To totalStudents 'iterate over rows
	studentPractice = Cells(i,8).Value
    If studentPractice >= yDay Then
    	Cells(i,8).Font.Color = RGB(34,139,34) 'Mark green
    Else
    	Cells(i,8).Font.Color = RGB(255, 0, 0) 'Otherwise mark red
    End If
Next i

'Label 2 days ago
'********************************************
For i = 2 To totalStudents 'iterate over rows
	studentPractice = Cells(i,9).Value
    If studentPractice >= twoDay Then
    	Cells(i,9).Font.Color = RGB(34,139,34) 'Mark green
    Else
    	Cells(i,9).Font.Color = RGB(255, 0, 0) 'Otherwise mark red
    End If
Next i

'Label 7 days ago
'********************************************
For i = 2 To totalStudents 'iterate over rows
	studentPractice = Cells(i,10).Value
    If studentPractice >= sevenDay Then
    	Cells(i,10).Font.Color = RGB(34,139,34) 'Mark green
    Else
    	Cells(i,10).Font.Color = RGB(255, 0, 0) 'Otherwise mark red
    End If
Next i


'Label last 30 days
'********************************************
For i = 2 To totalStudents 'iterate over rows
	studentPractice = Cells(i,11).Value
    If studentPractice >= thirtyDay Then
    	Cells(i,11).Font.Color = RGB(34,139,34) 'Mark green
    Else
    	Cells(i,11).Font.Color = RGB(255, 0, 0) 'Otherwise mark red
    End If
Next i


'dark green RGB(34,139,34)
'light grey RGB(224,224,224)
'dark red RGB(204,0,0)
'yellow RGB(255,255,0)
' orange RGB(255,128,0)
'Bold
Worksheets("Master").Range("A:R").Font.Bold = True
