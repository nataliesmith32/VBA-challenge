Attribute VB_Name = "stockdatavba"
Sub stockanalysis():

    For Each ws In Sheets
    
     ' declare variables values and names
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
     
        Dim TotalStockVol As Double
        TotalStockVol = 0
     
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
     
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
     
        Dim GreatestTotalVol As Double
        GreatestTotalVol = 0
        
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
        
        Dim PreviousAmountInteger
        PreviousAmount = 2
     
     ' set column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
     
     'grab the last row
     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
      ' loop through to print ticker
        For i = 2 To LastRow
        
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryTableRow).Value = Ticker
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                TotalStockVolume = 0

    ' printing for yearly stock changes
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SummaryTableRow).Value = YearlyChange

                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If

                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
            
            Next i

' loop for summary table, grab last row first
   LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If
                
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("Q3").NumberFormat = "0.00%"

            Next i
            
        ws.Columns("I:Q").AutoFit

    Next ws

End Sub

