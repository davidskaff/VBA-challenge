Sub StockDataAnalysis()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim LastRow As Long
        Dim i As Long
        Dim j As Integer

        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double
        Dim IncreaseTicker As String
        Dim DecreaseTicker As String
        Dim VolumeTicker As String

        j = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        OpenPrice = ws.Cells(2, 3).Value

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                If OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

                ws.Cells(j, 9).Value = Ticker
                ws.Cells(j, 10).Value = YearlyChange
                ws.Cells(j, 11).Value = PercentChange
                ws.Cells(j, 12).Value = TotalStockVolume

                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    IncreaseTicker = Ticker
                End If

                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    DecreaseTicker = Ticker
                End If

                If TotalStockVolume > GreatestVolume Then
                    GreatestVolume = TotalStockVolume
                    VolumeTicker = Ticker
                End If

                j = j + 1
                OpenPrice = ws.Cells(i + 1, 3)
                TotalStockVolume = 0

            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = IncreaseTicker
        ws.Cells(2, 17).Value = GreatestIncrease

        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = DecreaseTicker
        ws.Cells(3, 17).Value = GreatestDecrease

        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = VolumeTicker
        ws.Cells(4, 17).Value = GreatestVolume

        ' Formatting
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To LastRow
            ' Convert to percentage format
            ws.Cells(i, 11).NumberFormat = "0.00%"
            ' Color yearly change column
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0) ' Green
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0) ' Red
            End If
        Next i

    Next ws

End Sub
