Attribute VB_Name = "Module1"
Sub StockData_Bonus()
'Labeling variables
    Dim Ticker As String
    Dim Stock_Volume As Double
        Stock_Volume = 0
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    Dim lr As Long
        lr = Cells(Rows.Count, 1).End(xlUp).Row
    Dim Starting As Double
        Starting = 0
    Dim Closing As Double
        Closing = 0
    Dim Yearlychange As Double
        Yearlychange = 0
    Dim Percentchange As Double
        Percentchnage = 0
    Dim MaxCompany As String
    Dim MinCompany As String
    Dim MaxPercent As Double
        MaxPercent = 0
    Dim MinPercent As Double
        MinPercent = 0
    Dim MaxVolTicker As String
    Dim MaxVolume As Double
        MaxVolume = 0
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
                If Starting <> 0 Then
                    Percentchange = Yearlychange / Starting
                End If
'Identifying where to put responses
                    Range("I" & Summary_Table_Row).Value = Ticker
                    Range("J" & Summary_Table_Row).Value = Yearlychange
'Color conditional
                    If (Yearlychange > 0) Then
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    ElseIf (Yearlychange <= 0) Then
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                    Range("K" & Summary_Table_Row).Value = Percentchange
                    Range("L" & Summary_Table_Row).Value = Stock_Volume
                    Summary_Table_Row = Summary_Table_Row + 1
                    Starting = Cells(i + 1, 3).Value
                    If (Percentchange > MaxPercent) Then
                        MaxPercent = Percentchange
                        MaxCompany = Ticker
                    ElseIf (Percentchange < MinPercent) Then
                        MinPercent = Percentchange
                        MinCompany = Ticker
                    End If
                    Percentchange = 0
                    Stock_Volume = 0
                Else
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                End If

'Identifying where to put Percentage increase & decrease
                Range("N2").Value = "Max % Increase"
                Range("N3").Value = "Max % Decrease"
                Range("N4").Value = "Max Total Volume"
                Range("O2").Value = MaxCompany
                Range("O3").Value = MinCompany
                Range("O3").Value = MaxVolume
                Range("P2").Value = (CStr(MaxPercent) & "%")
                Range("P3").Value = (CStr(MinPercent) & "%")
                Range("P4").Value = MaxVolume
    Next i
End Sub
