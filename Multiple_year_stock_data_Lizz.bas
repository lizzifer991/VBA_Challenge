Attribute VB_Name = "Module1"
Sub StockData()
'Run the code across all worksheets
    For Each ws In Worksheets
'Labeling variables
    Dim lr As Long
        lr = Cells(Rows.Count, 1).End(xlUp).Row
    Dim Ticker As String
    Dim Stock_Volume As Double
        Stock_Volume = 0
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    Dim Starting As Double
        Starting = Cells(2, 3)
    Dim Closing As Double
        For i = 2 To lr
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
'Ticker population in I column
                Ticker = Cells(i, 1).Value
'Calculate yearly change
                Starting = Cells(i + 1, 3).Value
                Closing = Cells(i, 6).Value
                yearlychange = Closing - Starting
'Calculate percent change & formate to percentage
                Percentchange = yearlychange / Starting * 100 & "%"
'Identifying where to put responses
                    Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    Range("I" & Summary_Table_Row).Value = Ticker
                    Range("J" & Summary_Table_Row).Value = yearly_change
                    Range("K" & Summary_Table_Row).Value = percent_change
                    Range("L" & Summary_Table_Row).Value = Stock_Volume
                    Summary_Table_Row = Summary_Table_Row + 1
                    Stock_Volume = 0
                Else
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                End If
            Next i
    Next ws
End Sub
