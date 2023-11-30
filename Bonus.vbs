Sub bonus()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As LongLong

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

If ws.Range("L" & i).Value > Greatest_Increase Then
Greatest_Increase = ws.Range("L" & i).Value

ws.Cells(2, 16).Value = ws.Cells(i, 10).Value
ws.Cells(2, 17).Value = Greatest_Increase
ws.Range("Q2").NumberFormat = "0.00%"

End If

Next i

For i = 2 To LastRow

If ws.Range("L" & i).Value < Greatest_Decrease Then
Greatest_Decrease = ws.Range("L" & i).Value

ws.Cells(3, 16).Value = ws.Cells(i, 10).Value
ws.Cells(3, 17).Value = Greatest_Decrease
ws.Range("Q3").NumberFormat = "0.00%"
End If

Next i

For i = 2 To LastRow

If ws.Range("M" & i).Value > Greatest_Volume Then
Greatest_Volume = ws.Range("M" & i).Value

ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
ws.Cells(4, 17).Value = Greatest_Volume
End If

Next i

ws.Columns("O:R").AutoFit
Next ws

End Sub

