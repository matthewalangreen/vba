'calculate percent complete
'*******************************
Dim r As Integer 'row
Dim num As Integer
Dim denom As Integer
	For r = 2 To 137
	num = Cells(r, 7).Value
	denom = Cells(r, 13).Value
	Cells(r, 12).Value = num / denom * 100
Next r


Dim r As Integer 'row