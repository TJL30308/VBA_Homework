Sub StockTicker()

For Each ws In Worksheets

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Volume"
    ws.Range("N2") = "Greatest Percent Increase"
    ws.Range("N3") = "Greatest Percent Decrease"
    ws.Range("N4") = "Greatest Total Volume"

    Dim Ticker As String

    Dim TotalStockVolume
        TotalStockVolume = 0

    Dim SummaryTableRow
      SummaryTableRow = 2

    Dim LastRow
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        Dim YearClose As Double
    
        Dim YearOpen As Double

    For i = 2 To LastRow

         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                ws.Range("I" & SummaryTableRow).Value = Ticker
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                SummaryTableRow = SummaryTableRow + 1
                TotalStockVolume = 0
                YearClose = ws.Cells(i, 6).Value

            Else
        
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                YearOpen = ws.Cells(i, 3).Value
                ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
                    
                    If Cells(SummaryTableRow, 11) >= 0 Then
                        ws.cells(SummaryTableRow, 11) = (YearOpen - YearClose) / YearOpen 
                        ws.Cells(SummaryTableRow, 11).Interior.Color = RGB(0, 255, 0)
                    Else
                        ws.cells(SummaryTableRow, 11) = 0
                        ws.Cells(SummaryTableRow, 11).Interior.Color = RGB(255, 0, 0)
                    End If

        End If
        
        ws.Cells(SummaryTableRow, 10) = YearOpen - YearClose

    Next i
    
Next ws

End Sub