Dim r As Integer
Application.DisplayAlerts = False
Application.ScreenUpdating = False

ActiveWorkbook.Sheets.Add Before:=Worksheets("Master")
Worksheets("Sheet1").Name = "Printable"

'***************************  Copy cell titles ********************************************'
Worksheets("Printable").Activate
Cells(1,1).Value = "Name"
Cells(1,5).Value = "Name"
Cells(1,2).Value = "Khan Mins"
Cells(1,6).Value = "Khan Mins"
Cells(1,3).Value = "Khan %"
Cells(1,7).Value = "Khan %"
Cells(1,4).Value = "POTW"
Cells(1,8).Value = "POTW"




'********************** Move data from other sheet ******************************************'
For r = 2 To 80
  Worksheets("Printable").Cells(r,1).Value = Worksheets("Quick Report").Cells(r,4).Value
  Worksheets("Printable").Cells(r,5).Value = Worksheets("Quick Report").Cells(r+80,4).Value
Next r


For r = 2 To 80
  Worksheets("Printable").Cells(r,2).Value = Worksheets("Quick Report").Cells(r,1).Value
  Worksheets("Printable").Cells(r,3).Value = Worksheets("Quick Report").Cells(r,2).Value
  Worksheets("Printable").Cells(r,4).Value = Worksheets("Quick Report").Cells(r,3).Value

  Worksheets("Printable").Cells(r,6).Value = Worksheets("Quick Report").Cells(r+80,1).Value
  Worksheets("Printable").Cells(r,7).Value = Worksheets("Quick Report").Cells(r+80,2).Value
  Worksheets("Printable").Cells(r,8).Value = Worksheets("Quick Report").Cells(r+80,3).Value
Next r

'*************  Formatting ****************'
For r = 2 To 170
  Cells(r,1).Font.Color = RGB(0,0,0)
  Cells(r,2).Font.Color = RGB(0,0,0)
  Cells(r,3).Font.Color = RGB(0,0,0)
  Cells(r,4).Font.Color = RGB(0,0,0)
  Cells(r,5).Font.Color = RGB(0,0,0)
  Cells(r,6).Font.Color = RGB(0,0,0)
  Cells(r,7).Font.Color = RGB(0,0,0)
  Cells(r,8).Font.Color = RGB(0,0,0)
Next r

Worksheets("Printable").Range("A:Z").Font.Bold = True

Range("A:A").ColumnWidth = 18
Range("E:E").ColumnWidth = 18
Range("F:F").NumberFormat = "#"

Worksheets("Printable").Activate
For r = 2 To 170
  If Cells(r,4).Value = "Yes" Then
    Cells(r,4).Font.Color = RGB(34,139,34)
  Else
    Cells(r,4).Font.Color = RGB(255,0,0)
  End If

  If Cells(r,8).Value = "Yes" Then
    Cells(r,8).Font.Color = RGB(34,139,34)
  Else
    Cells(r,8).Font.Color = RGB(255,0,0)
  End If
Next r

'**************************   Color alternating rows ***************************************'

ActiveSheet.Range("A:Z").Select
Set sh = Worksheets("Printable")
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


    '******************************************
    Range("E1").EntireColumn.Insert
    Range("E:E").ColumnWidth = 5




        '************************************************ THIS IS THE END ***********************************************************************'
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        '****************************************************************************************************************************************'
