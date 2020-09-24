Sub Stockmarket():
    For Each ws In Worksheets
        'Naming header
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Assigning the variables
        Dim i As Long
        Dim ticker As String
        Dim OpenPrice As Double
        OpenPrice = 0
        Dim ClosePrice As Double
        ClosePrice = 0
        Dim StockVolume As Double
        StockVolume = 0
        
        Dim YearlyChange As Double
        YearlyChange = 0
        Dim PercentChange As Double
        PercentChange = 0
        
        Dim TickerRow As Long
        TickerRow = 2
        Dim Lastrow As Long
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To Lastrow
        
        OpenPrice = ws.Cells(2, 3).Value

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Ticker Symbol
            ticker = ws.Cells(i, 1).Value

            'Calcualting Yearly Change Change amd Percent Change
            ClosePrice = ws.Cells(i, 6).Value
            YearlyChange = ClosePrice - OpenPrice

        If OpenPrice <> 0 Then
            PercentChange = (YearlyChange / OpenPrice) * 100
            StockVolume = StockVolume + ws.Cells(i, 7).Value

            'Printing values into the cells
            ws.Range("I" & TickerRow).Value = ticker
            ws.Range("J" & TickerRow).Value = YearlyChange
            ws.Range("K" & TickerRow).Value = (Str(PercentChange) & "%")
            ws.Range("L" & TickerRow).Value = StockVolume
            TickerRow = TickerRow + 1
            
            'Reset the values
            YearlyChange = 0
            OpenPrice = ws.Cells(i + 1, 3).Value
            ClosePrice = 0
            PercentChange = 0
            StockVolume = 0
            
            End If

            Else
            StockVolume = StockVolume + ws.Cells(i, 7).Value
        End If

    'Color formatting
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i
Next ws
End Sub
