Attribute VB_Name = "Module1"
' vba-challange - 'The VBA of Wall Street'
Sub MultipleYearStockData():

    ' Loop through all the stocks for one year
    For Each ws In Worksheets

        ' Output the following as new columns:

            ' Ticker symbol
                ws.Range("I1").Value = "Ticker"
            ' Yearly change from opening price at beginning of a given year to closing price at EOY
                ws.Range("J1").Value = "Yearly Change"
            ' % change from opening price at beginning of a given year to closing price at EOY
                ws.Range("K1").Value = "Percent Change"
            ' Total stock volume (of each stock)
                ws.Range("L1").Value = "Total Stock Volume"
            ' Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
            ' Greatest % increase
                ws.Range("O2").Value = "Greatest % Increase"
            ' Greatest % decrease
                ws.Range("O3").Value = "Greatest % Decrease"
            ' Greatest total volume
                ws.Range("O4").Value = "Greatest Total Volume"
            ' Determine new Ticker symbol data
                ws.Range("P1").Value = "Ticker"
            ' Determine new Value data
                ws.Range("Q1").Value = "Value"

            ' Apply conditional formatting to highlight positive change (green) & negative change (red)

    ' Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once

        ' Declare variables and set default variables
        Dim TickerName As String
        Dim LastRow As Long
        Dim TotalTickerVolume As Double
            TotalTickerVolume = 0
        Dim SummaryTableRow As Long
            SummaryTableRow = 2
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PreviousAmount As Long
            PreviousAmount = 2
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
            GreatestIncrease = 0
        Dim GreatestDecrease As Double
            GreatestDecrease = 0
        Dim LastRowValue As Long
        Dim GreatestTotalVolume As Double
            GreatestTotalVolume = 0

        ' Determine last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            ' Add to "Ticker Total Volume"
            TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
            ' Make sure data/loop is within "Total Ticker Volume" range
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Assign ticker name
                TickerName = ws.Cells(i, 1).Value
                ' Print ticker name in summary table
                ws.Range("I" & SummaryTableRow).Value = TickerName
                ' Print ticker total amount to summary table
                ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
                ' Reset ticker total
                TotalTickerVolume = 0

                ' Yearly open
                YearlyOpen = ws.Range("C" & PreviousAmount)
                ' Yearly close
                YearlyClose = ws.Range("F" & i)
                ' Yearly change
                YearlyChange = YearlyClose - YearlyOpen
                ' Print to summary table
                ws.Range("J" & SummaryTableRow).Value = YearlyChange

                ' Calculate % change
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
                ' Format Double with % symbol and to two decimal places
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                ' Apply conditional formatting to highlight positive change (green)
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                ' & negative change (red)
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                ' Add one to summary table row
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
            Next i

            ' Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            ' Start loop for final results
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

            Next i
        ' Format Double to with % symbol and to two decimal places
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
        ' Auto fit tables
        ws.Columns("I:Q").AutoFit

    Next ws

End Sub
