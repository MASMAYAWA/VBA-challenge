Sub stockvalue()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"

Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Value As LongLong

Yearly_Change = 0
Percent_Change = 0
Total_Stock_Value = 0

Dim Summary_Row As Integer
Summary_Row = 2
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker = ws.Cells(i, 1).Value
Yearly_Change = Yearly_Change + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)
Percent_Change = Percent_Change + ((ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 3))
Total_Stock_Value = Total_Stock_Value + ws.Cells(i, 7).Value

ws.Range("J" & Summary_Row).Value = Ticker
ws.Range("K" & Summary_Row).Value = Yearly_Change
ws.Range("L" & Summary_Row).Value = Percent_Change
ws.Range("M" & Summary_Row).Value = Total_Stock_Value
ws.Range("L" & Summary_Row).NumberFormat = "0.00%"

Summary_Row = Summary_Row + 1

Yearly_Change = 0
Percent_Change = 0
Total_Stock_Value = 0

Else

Yearly_Change = Yearly_Change + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)
Percent_Change = Percent_Change + ((ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 3))
Total_Stock_Value = Total_Stock_Value + ws.Cells(i, 7).Value

End If

If ws.Range("K" & Summary_Row).Value > 0 Then
ws.Range("K" & Summary_Row).Interior.ColorIndex = 4

Else
ws.Range("K" & Summary_Row).Interior.ColorIndex = 3

End If

Next i
ws.Columns("J:M").AutoFit

Next ws

End Sub