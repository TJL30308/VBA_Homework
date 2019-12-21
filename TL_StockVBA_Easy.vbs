Attribute VB_Name = "Module1"
Sub StockTicker()

    Range("I1") = "Ticker"
    Range("J1") = "Volume"

    Dim Ticker As String

    Dim TotalStockVolume
        TotalStockVolume = 0

    Dim SummaryTableRow
      SummaryTableRow = 2

    Dim LastRow
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow

         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
                Range("I" & SummaryTableRow).Value = Ticker
                Range("J" & SummaryTableRow).Value = TotalStockVolume
                SummaryTableRow = SummaryTableRow + 1
                TotalStockVolume = 0

            Else
        
                TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            
        End If
        
    Next i

End Sub

