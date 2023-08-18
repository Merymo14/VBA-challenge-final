# VBA-challenge-final
Week 2 Challenge completed
Sub Stock_Analysis()

Dim total As Double
Dim rowIndex As Long
Dim Change As Double
Dim Columnindex As Integer
Dim Star As Long
Dim rowcount As Long
Dim percentChange As Double
Dim Days As Integer
Dim DailyChange As Single
Dim Averagechange As Double
Dim ws As Worksheet

For Each ws In Worksheets
Columnindex = 0
total = 0
Change = 0
Start = 2
DailyChange = 0
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row

For rowIndex = 2 To rowcount

If ws.Cells(rowIndex + 1, 1).Value <> ws.Cells(rowIndex, 1).Value Then
total = total + ws.Cells(rowIndex, 7).Value
If total = 0 Then

ws.Range("I" & 2 + Columnindex).Value = Cells(rowIndex, 1).Value
ws.Range("J" & 2 + Columnindex).Value = 0
ws.Range("K" & 2 + Columnindex).Value = "%" & 0
ws.Range("L" & 2 + Columnindex).Value = 0

Else
If ws.Cells(Start, 3) = 0 Then
For find_value = Start To rowIndex
If ws.Cells(find_value, 3).Value <> 0 Then

Start = find_value
Exit For
End If
Next find_value
End If

Change = (ws.Cells(rowIndex, 6) - ws.Cells(Start, 3))


percentChange = Change / ws.Cells(Start, 3)

Start = rowIndex + 1

ws.Range("I" & 2 + Columnindex) = ws.Cells(rowIndex, 1).Value
ws.Range("J" & 2 + Columnindex) = Change
ws.Range("J" & 2 + Columnsindex).NumberFormat = "0.00"
ws.Range("K" & 2 + Columnindex).Value = percentChange
ws.Range("K" & 2 + Columnindex).NumberFormat = "0.00%"
ws.Range("L" & 2 + Columnindex).Value = total


Select Case Change
Case Is > 0
ws.Range("J" & 2 + Columnindex).Interior.ColorIndex = 4

Case Is < 0
ws.Range("J" & 2 + Columnindex).Interior.ColorIndex = 3

Case Else
ws.Range("J" & 2 + Columnindex).Interior.ColorIndex = 0

End Select


End If

total = 0
Change = 0
Columnindex = Columnindex + 1
Days = 0
DailyChange = 0

Else
total = total + ws.Cells(rowIndex, 7).Value


End If

Next rowIndex

ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowcount)) * 100
ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowcount)) * 100
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowcount))

increase_Number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)

Decrease_Number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)

Volume_Number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowcount)), ws.Range("L2:L" & rowcount), 0)


ws.Range("P2") = ws.Cells(increase_Number + 1, 9)
ws.Range("P3") = ws.Cells(Decrease_Number + 1, 9)
ws.Range("P4") = ws.Cells(Volume_Number + 1, 9)



Next ws



End Sub

