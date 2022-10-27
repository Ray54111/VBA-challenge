Sub Max_Summary()
Dim maxyearlychange As Double
Dim minyearlychange As Double
Dim maxvolume As Double

maxyearlychange = 0
minyearlychange = 0
maxvolume = 0

lastrowsum = Cells(Rows.Count, 9).End(xlUp).Row



For i = 2 To lastrowsum

If Cells(i, 11).Value > maxyearlychange Then

maxyearlychange = Cells(i, 11).Value

Cells(2, 15) = maxyearlychange

Cells(2, 14) = Cells(i, 9).Value

ElseIf Cells(i, 11).Value < minyearlychange Then

minyearlychange = Cells(i, 11).Value
Cells(3, 15) = maxyearlychange
Cells(3, 14) = Cells(i, 9).Value


ElseIf Cells(i, 12).Value > maxvolume Then

maxvolume = Cells(i, 12).Value

Cells(4, 15) = maxvolume

Cells(4, 14) = Cells(i, 9).Value


End If


Next i





End Sub
