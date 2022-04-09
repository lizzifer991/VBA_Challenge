Attribute VB_Name = "Module1"
Sub StockData()
'Labeling variables
    Dim Ticker As String
    Dim Stock_Volume As Double
        Stock_Volume = 0
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    Dim lr As Long
        lr = Cells(Rows.Count, 1).End(xlUp).Row
    Dim Starting As Double
        Starting = Cells(2, 3)
    Dim Closing As Double
        For i = 2 To lr
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
'Ticker population in I column
                Ticker = Cells(i, 1).Value
'Stocks total volume popluates in L column
                Stock_Volume = Stock_Volume + Cells(i, 7).Value
'Calculate yearly change
                Closing = Cells(i, 6).Value
                Yearlychange = Closing - Starting
'Calculate percent change & formate to percentage
                Percentchange = Yearlychange / Starting * 100 & "%"
'Identifying where to put responses
                    Range("I" & Summary_Table_Row).Value = Ticker
                    Range("J" & Summary_Table_Row).Value = Yearlychange
                    Range("K" & Summary_Table_Row).Value = Percentchange
                    Range("L" & Summary_Table_Row).Value = Stock_Volume
                    Summary_Table_Row = Summary_Table_Row + 1
                    Stock_Volume = 0
                Else
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                End If
            Next i
End Sub
