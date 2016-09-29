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


Dim r As Integer 'rowDim mastered As IntegerDim l2 As IntegerDim l1 As IntegerDim practiced As IntegerDim total As IntegerDim result As Integer    For r = 2 To 147        mastered = Cells(r, 7).Value        l2 = Cells(r, 6).Value        l1 = Cells(r, 5).Value        practiced = Cells(r, 4).Value        total = Cells(r, 13).Value                result = ((4 * mastered + 3 * l2 + 2 * l1 + practiced) / (4 * total)) * 100        Cells(r, 12).Value = result        Next r
