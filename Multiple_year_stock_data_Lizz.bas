Attribute VB_Name = "Module1"
Sub Stocks_Bond()
    For Each ws In Worksheets
        Dim LR As Long
        LR = Cells(Rows.Count, 1).End(xlUp).Row
            Dim Ticker As String
            Dim Stock_Volume As Double
            Dim close_val As Integer
            Dim open_val As Integer
            Dim yearly_change As Double
            Stock_Volume = 0
            yearly_change = 0
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
            For i = 2 To LR
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    Ticker = Cells(i, 1).Value
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    close_val = Cells(i, 6).Value
                    open_val = Cells(i, 3).Value
                    yearly_change = close_val - open_val
                    Range("I" & Summary_Table_Row).Value = Ticker
                    Range("J" & Summary_Table_Row).Value = yearly_change
                    Range("L" & Summary_Table_Row).Value = Stock_Volume
                    Summary_Table_Row = Summary_Table_Row + 1
                    yearly_change = 0
                    Stock_Volume = 0
                Else
                    yearly_change = close_val - open_val
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                End If
            Next i
        Next ws
End Sub
