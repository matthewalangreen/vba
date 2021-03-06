'Globals
'*****************************************
Dim weekNumber As Integer
weekNumber = 2
Dim totalStudents As Integer
totalStudents = 164

'Pull values from mission specific into summary
'*************************************************
Sheets("Summary").range("D:G").value = Sheets("Mission-specific").range("D:G").value
Sheets("Summary (2)").range("D:G").value = Sheets("Mission-specific (2)").range("D:G").value
Sheets("Summary (3)").range("D:G").value = Sheets("Mission-specific (3)").range("D:G").value
Sheets("Summary (4)").range("D:G").value = Sheets("Mission-specific (4)").range("D:G").value
Sheets("Summary (5)").range("D:G").value = Sheets("Mission-specific (5)").range("D:G").value
Sheets("Summary (6)").range("D:G").value = Sheets("Mission-specific (6)").range("D:G").value
Sheets("Summary (7)").range("D:G").value = Sheets("Mission-specific (7)").range("D:G").value

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

Sheets("Videos").Delete
Sheets("Videos (2)").Delete
Sheets("Videos (3)").Delete
Sheets("Videos (4)").Delete
Sheets("Videos (5)").Delete
Sheets("Videos (6)").Delete
Sheets("Videos (7)").Delete

Sheets("Badges").Delete
Sheets("Badges (2)").Delete
Sheets("Badges (3)").Delete
Sheets("Badges (4)").Delete
Sheets("Badges (5)").Delete
Sheets("Badges (6)").Delete
Sheets("Badges (7)").Delete

Sheets("Points").Delete
Sheets("Points (2)").Delete
Sheets("Points (3)").Delete
Sheets("Points (4)").Delete
Sheets("Points (5)").Delete
Sheets("Points (6)").Delete
Sheets("Points (7)").Delete

Sheets("Exercises").Delete
Sheets("Exercises (2)").Delete
Sheets("Exercises (3)").Delete
Sheets("Exercises (4)").Delete
Sheets("Exercises (5)").Delete
Sheets("Exercises (6)").Delete
Sheets("Exercises (7)").Delete

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
        End If
         'Data range in worksheet - starts from second row as first rows are the header rows in all worksheets
        Set rng = sht.Range(sht.Cells(2, 1), sht.Cells(65536, 1).End(xlUp).Resize(, colCount))
         'Put data into the Master worksheet
        trg.Cells(65536, 1).End(xlUp).Offset(1).Resize(rng.Rows.count, rng.Columns.count).Value = rng.Value
    Next sht
     'Fit the columns in Master worksheet
    trg.Columns.AutoFit
     
     'Screen updating should be activated
    Application.ScreenUpdating = True
    

    Application.DisplayAlerts = False
    Sheets("Summary").Delete
    Sheets("Summary (2)").Delete
    Sheets("Summary (3)").Delete
    Sheets("Summary (4)").Delete
    Sheets("Summary (5)").Delete
    Sheets("Summary (6)").Delete
    Sheets("Summary (7)").Delete
    
    
    Application.DisplayAlerts = True

'Add hyperlinks
'******************************************
'Worksheets("Master").Activate
'Columns("N").Select
'Range("N1").Value = "Hyperlinks"
' Written by: Michael Milette
' Copyright 2011-2012 TNG Consulting Inc.
' Purpose: Converts the selected text into hyperlinks.
' Note: HTTP is assumed if not specified in the text.

  'Dim Cell As Range
 ' For Each Cell In Intersect(Selection, ActiveSheet.UsedRange)
  '  If Trim(Cell) > "" Then
  '    If Left(Trim(Cell), 4) = "http" Then  ' handles http and https
  '      ActiveSheet.Hyperlinks.Add Cell, Trim(Cell.Value)
  '    Else ' Default to http if no protocol was specified.
  '      ActiveSheet.Hyperlinks.Add Cell, "http://" & Trim(Cell.Value)
  '    End If
  '  End If
'  Next






'Rotate label text
'******************************************
Rows(1).Orientation = xlUpward
Range("A1").Orientation = xlHorizontal
Range("M1").Orientation = xlHorizontal
Range("N1").Orientation = xlHorizontal

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
  Range("P1").Hyperlinks.Delete
  Range("P1").Font.Bold = True
  Range("H:J").NumberFormat = "#,###.#"
  Range("L:L").NumberFormat = "###"

'Set column widths
'******************************************

Range("B:G").ColumnWidth = 5
Range("H:J").ColumnWidth = 8
Range("K:M").ColumnWidth = 5
Range("N:N").ColumnWidth = 8
Range("P:P").ColumnWidth = 12

'Remove duplicates
'******************************************
Sheets("Master").Range("A:P").RemoveDuplicates Columns:=Array(1, 16), Header:=xlYes

'calculate total topics
'*****************************
Dim l As Integer 'row

For l = 2 To totalStudents
    If Cells(l, 15).Value = "RA Algebra 1" Then
       Cells(l, 13).Value = 185
    ElseIf Cells(l, 15).Value = "RA Algebra 2/Trig" Then
       Cells(l, 13).Value = 141
    ElseIf Cells(l, 15).Value = "RA Geometry" Then
        Cells(l, 13).Value = 104
    ElseIf Cells(l, 15).Value = "RA Pre Algebra" Then
       Cells(l, 13).Value = 182
    ElseIf Cells(l, 15).Value = "RA Calculus A" Then
       Cells(l, 13).Value = 87
    ElseIf Cells(l, 15).Value = "RA Precalculus" Then
       Cells(l, 13).Value = 106
    ElseIf Cells(l, 15).Value = "RA Algebra 2/Trig, RA Trigonometry" Then
       Cells(l, 13).Value = 141
    ElseIf Cells(l, 15).Value = "RA Calculus A, RA Geometry" Then
       Cells(l, 13).Value = 104
    ElseIf Cells(l, 15).Value = "RA Algebra 1, RA Geometry" Then
       Cells(l, 13).Value = 185
    Else
       Cells(l, 13).Value = 265
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

'Label names as "on track" or not
'******************************************
Dim i As Integer 'row
Dim j As Integer 'column
Dim count As Integer
Dim rowSum As Integer
'Dim weekNumber As Integer
'weekNumber = 2
Dim minPracticedTotal As Integer


minPracticedTotal = weekNumber * 15 'fifteen items practiced total per week
count = 0

For i = 2 To totalStudents 'row
    For j = 3 To 7 'column
        count = count + Cells(i, j).Value
    Next j
    'Cells(i, 9).Value = count
    
    If count < minPracticedTotal Then
    Cells(i, 4).Interior.Color = RGB(255, 255, 153)
    Cells(i, 5).Interior.Color = RGB(255, 255, 153)
    Cells(i, 6).Interior.Color = RGB(255, 255, 153)
    Cells(i, 7).Interior.Color = RGB(255, 255, 153)
    Else
    Cells(i, 4).Interior.Color = RGB(152, 251, 152)
    Cells(i, 5).Interior.Color = RGB(152, 251, 152)
    Cells(i, 6).Interior.Color = RGB(152, 251, 152)
    Cells(i, 7).Interior.Color = RGB(152, 251, 152)
    End If
    count = 0
Next i

'Label names as behind percent complete or not
'******************************************
Dim x As Integer 'row
Dim comp As Integer 'how complete are you?
comp = weekNumber * 3

For x = 2 To totalStudents 'row
    If Cells(x, 12).Value >= comp Then
            Cells(x, 12).Interior.Color = RGB(152, 251, 152)
       ElseIf Cells(x, 12).Value >= comp - 2 Then
            Cells(x, 12).Interior.Color = RGB(255, 255, 152)
        Else
            Cells(x, 12).Interior.Color = RGB(255, 91, 79)
        End If
Next x

'Sort
'********************************

Columns("A:P").Sort key1:=Range("L:L"), order1:=xlAscending, Header:=xlYes

Columns("A:P").Sort key1:=Range("O:O"), order1:=xlAscending, Header:=xlYes

Columns("A:P").Sort key1:=Range("A:A"), order1:=xlAscending, Header:=xlYes

'Borders
'********************************
 Range("L2:L137").Borders(xlEdgeLeft).LineStyle = xlContinuous
 Range("L2:L137").Borders(xlEdgeRight).LineStyle = xlContinuous
 Range("L2:L137").Borders(xlEdgeTop).LineStyle = xlContinuous
 Range("L2:L137").Borders(xlEdgeBottom).LineStyle = xlContinuous


