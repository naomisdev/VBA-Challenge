Sub stock():
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
     
    Dim tickerRow As Double
    Dim lastRow As Double
    Dim opening As Double
    Dim change As Double
    Dim total As Double
    
    total = 0
    opening = 0
    tickerRow = 2
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
    
        If opening = 0 Then opening = Cells(i, 3)
        total = total + Cells(i, 7)
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            Cells(tickerRow, 9) = Cells(i, 1)
            change = (Cells(i, 6) - opening)
            If change > 0 Then Cells(tickerRow, 10).Interior.ColorIndex = 4
            If change < 0 Then Cells(tickerRow, 10).Interior.ColorIndex = 3
            Cells(tickerRow, 10).Borders().LineStyle = xlContinuous
            Cells(tickerRow, 10) = change
            If change = 0 Then
                Cells(tickerRow, 11) = 0
            Else
                Cells(tickerRow, 11) = (change / opening)
            End If
            Cells(tickerRow, 12) = total
            total = 0
            tickerRow = tickerRow + 1
            opening = Cells(i + 1, 3).Value
        End If
    Next i
End Sub
