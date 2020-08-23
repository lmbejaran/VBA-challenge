Sub stock_reports()

Dim yearlychange As Double
Dim percentchange As Double
    
   For Each ws In Worksheets
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "YearlyChange"
    Range("K1").Value = "PercentChange"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Total = 0
    j = 1
    Start = 2
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   
   For i = 2 To RowCount
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Range("I" & j + 1).Value = ws.Cells(i, 1).Value
            Total = Total + ws.Cells(i, 7).Value - Total
            ws.Range("L" & j + 1).Value = Total
            yearlychange = ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value
            ws.Range("J" & j + 1).Value = yearlychange
            percentchange = yearlychange / ws.Cells(Start, 3).Value
            ws.Range("K" & j + 1).Value = percentchange
            ws.Range("K" & j + 1).NumberFormat = "0.00%"
            
            If yearlychange >= 0 Then
            ws.Range("J" & j + 1).Interior.ColorIndex = 4
            ElseIf yearlychange < 0 Then
                ws.Range("J" & j + 1).Interior.ColorIndex = 3
            End If
            j = j + 1
    End If
    
Next i
    
    Next ws
End Sub