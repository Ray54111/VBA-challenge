Sub VBA_Challenge_try2()

'Creat last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

x = Range("a2: g1000").Value

    'Set variables
    Dim Ticker As String
    Dim VolumeTotal As Double
    Dim SummaryRow As Integer
    Dim Yearlychange As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PercentChange As Double
    
    
    
    VolumeTotal = 0
    SummaryRow = 2
    Yearlychange = 0
    OpenPrice = 0
    ClosePrice = 0
    PercentChange = 0
For i = 2 To lastrow
'For i = 2 To 1000
        
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then



'Define Ticker name
Ticker = Cells(i, 1)

ClosePrice = Cells(i, 6).Value

Yearlychange = ClosePrice - OpenPrice
PercentChange = Format((Yearlychange / OpenPrice) * 100, "0.00")






'add Total
VolumeTotal = VolumeTotal + Cells(i, 7).Value



'Transfer info to Summary Row

Range("I" & SummaryRow) = Ticker
Range("J" & SummaryRow) = Yearlychange
Range("K" & SummaryRow) = PercentChange & "%"
Range("L" & SummaryRow) = VolumeTotal


'Set Summary counter +1
SummaryRow = SummaryRow + 1

'Reset Volume Total
VolumeTotal = 0

ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then

OpenPrice = Cells(i, 3).Value

VolumeTotal = VolumeTotal + Cells(i, 7).Value




Else
    VolumeTotal = VolumeTotal + Cells(i, 7).Value
    
    

 

                End If
    
    
      
Next i

    


End Sub
