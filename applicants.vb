' *** Setup *** '

Application.ScreenUpdating = False
Application.DisplayAlerts = False


' *** Sorting ***'
Worksheets("Master").Activate
Columns("A:Z").Sort key1:=Range("A:A"), order1:=xlAscending, Header:=xlYes
Worksheets("Combined").Activate
Columns("A:Z").Sort key1:=Range("A:A"), order1:=xlAscending, Header:=xlYes


' *** loop over rows to combine sheets ***'
Dim i As Integer
Dim clicker As Integer
clicker = 2


For i = 2 To 200 'iterate over rows
    Do While Worksheets("Combined").Cells(i,1).Value <> Worksheets("Master").Cells(clicker,1).Value
    		clicker = clicker + 1
        If Worksheets("Master").Cells(clicker,1).Value = "" Then
          Exit Do
        End If
  	Loop

    Sheets("Master").Cells(clicker,4).Copy Sheets("Combined").Cells(i,13)
    Sheets("Master").Cells(clicker,5).Copy Sheets("Combined").Cells(i,14)
    Sheets("Master").Cells(clicker,6).Copy Sheets("Combined").Cells(i,15)
    Sheets("Master").Cells(clicker,7).Copy Sheets("Combined").Cells(i,16)
    Sheets("Master").Cells(clicker,8).Copy Sheets("Combined").Cells(i,17)
    Sheets("Master").Cells(clicker,9).Copy Sheets("Combined").Cells(i,18)
    Sheets("Master").Cells(clicker,10).Copy Sheets("Combined").Cells(i,19)
    Sheets("Master").Cells(clicker,11).Copy Sheets("Combined").Cells(i,20)
    Sheets("Master").Cells(clicker,12).Copy Sheets("Combined").Cells(i,21)
    Sheets("Master").Cells(clicker,13).Copy Sheets("Combined").Cells(i,22)
    clicker = 2
Next i

'Worksheets("Master").Cells(1, 3).Copy Worksheets("Combined").Cells(1, 13)'



' Conclusion'
Application.ScreenUpdating = True
Application.DisplayAlerts = True
