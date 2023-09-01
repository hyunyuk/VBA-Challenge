Sub multiple_year()

'Give variables
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    TotalVolume = 0
    Dim LastRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim SummaryRow As Long
    Dim ws As Worksheet
    Dim MaxPercentIncrease As Double
    Dim MaxPercentDecrease As Double
    Dim MaxTotalVolume As Double
    Dim SummaryLastRow As Long
    
    For Each ws In Worksheets
    
    SummaryRow = 2
     OpeningPrice = ws.Cells(2, 3).Value
'Give Column Names
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
'Find Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To LastRow
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        Ticker = ws.Cells(I, 1).Value
        
        ClosingPrice = ws.Cells(I, 6).Value
        YearlyChange = ClosingPrice - OpeningPrice
        TotalVolume = TotalVolume + ws.Cells(I, 7).Value
       
        If OpeningPrice <> 0 Then
            PercentChange = (YearlyChange / OpeningPrice)
            Else
            PercentChange = 0
        End If
        'Print the ticker, yearlychange, percentchange, and total volume to summary table
        ws.Range("I" & SummaryRow).Value = Ticker
        ws.Range("J" & SummaryRow).Value = YearlyChange
        ws.Range("K" & SummaryRow).Value = PercentChange
        ws.Range("L" & SummaryRow).Value = TotalVolume

        'Format PercentChange as percentage
        ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
        OpeningPrice = ws.Cells(I + 1, 3).Value
        'Color the Yearly Change
        If YearlyChange >= 0 Then
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
        ElseIf YearlyChange < 0 Then
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
        End If
        'Add one to the summary table
        SummaryRow = SummaryRow + 1
        'Reset the total volume
        TotalVolume = 0
    Else
    TotalVolume = TotalVolume + ws.Cells(I, 7).Value
    End If
    Next I
    SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For j = 2 To SummaryLastRow
        If ws.Range("L" & j).Value >= MaxTotalVolume Then
            MaxTotalVolume = Range("L" & j).Value
            ws.Range("P4").Value = ws.Cells(j, 9).Value
        End If
        ws.Range("Q4").Value = MaxTotalVolume
        
    Next j
    For k = 2 To SummaryLastRow
        If ws.Range("K" & k).Value >= MaxPercentIncrease Then
            MaxPercentIncrease = ws.Range("K" & k).Value
             ws.Range("P2").Value = ws.Cells(k, 9).Value
        End If
        ws.Range("Q2").Value = MaxPercentIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
    
    Next k
    For l = 2 To SummaryLastRow
        If ws.Range("K" & l).Value <= MaxPercentDecrease Then
            MaxPercentDecrease = ws.Range("K" & l).Value
                ws.Range("P3").Value = ws.Cells(l, 9).Value
        End If
        ws.Range("Q3").Value = MaxPercentDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
     
    Next l
    Next ws
            
End Sub
