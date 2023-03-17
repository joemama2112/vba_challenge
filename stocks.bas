Attribute VB_Name = "Module1"
Sub stocks()


Dim r As Long
Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim day As Long
Dim FirstOpen As Long
FirstOpen = 20200102
Dim LastClose As Long
LastClose = 20201231

Dim FirstOpenVal As Double
Dim LastCloseVal As Double

Dim TickSymb As String
Dim TickCount As Long
TickCount = 1

Dim TotVol As Long
Dim NewTotVol As Long
        
Range("J1").Value = "Ticker Symb."
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Volume"
        
For r = 2 To LastRow
day = Cells(r, 2).Value
If day = FirstOpen Then
TickSymb = Range("A" & r).Value
FirstOpenVal = Cells(r, 3).Value
TickCount = TickCount + 1
TotVol = Cells(r, 7).Value
ElseIf day = LastClose Then
LastCloseVal = Cells(r, 6).Value
TotVol = TotVol + Cells(r, 7).Value
'ElseIf day <> FirstOpen Or LastClose Then
'NewTotVol = TotVol + Cells(r, 7).Value
End If
Range("J" & TickCount).Value = TickSymb
Range("K" & TickCount).Value = LastCloseVal - FirstOpenVal
Range("L" & TickCount).Value = (FirstOpenVal - LastCloseVal) / FirstOpenVal
'Range("M" & TickCount).Value = NewTotVol
Next r


       
    
       




End Sub
