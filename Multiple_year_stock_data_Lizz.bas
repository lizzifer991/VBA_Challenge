Attribute VB_Name = "Module1"
Sub StockData()
Dim Ticker As String
Dim Stock_Volume As Double
Stock_Volume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
For i = 2 To 70926
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Stock_Volume = Stock_Volume + Cells(i, 7).Value
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("L" & Summary_Table_Row).Value = Stock_Volume
        Summary_Table_Row = Summary_Table_Row + 1
        Stock_Volume = 0
    Else
        Stock_Volume = Stock_Volume + Cells(i, 7).Value
    End If
Next i
        
End Sub