Sub Assignment_VBA()

Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate
Debug.Print ws.Name

Dim WorksheetsName As String
WorksheetsName = ws.Name
Dim TickerName As String
Dim TotalVolume As Double

TotalVolume = 0

Dim RowReference As Integer
RowReference = 2

worksheetname = ws.Name

Cells(1, "I").Value = "TickerName"
Cells(1, "J").Value = "TotalVolume"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
TotalVolume = 0

For i = 2 To LastRow

If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then

TickerName = ws.Cells(i, 1).Value
TotalVolume = TotalVolume + ws.Cells(i, 7).Value

ws.Range("I" & RowReference).Value = TickerName
ws.Range("J" & RowReference).Value = TotalVolume

RowReference = RowReference + 1
TotalVolume = 0

Else

TotalVolume = TotalVolume + ws.Cells(i, 7).Value

End If
Next i
Next ws
End Sub
