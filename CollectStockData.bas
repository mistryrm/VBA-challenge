Attribute VB_Name = "Module1"
Sub CollectStockData():

    For Each wSheet In Worksheets
        Dim currentRow As Long ' keeps track of which row I am currently looking at
        Dim startTickerRow As Long ' keeps track of starting row of the Ticker I am looking at
        Dim recordTickerRow As Long ' keep track of the row which I am recording the ticker calculations for
        Dim lastRowA As Long ' the very last row so I know when to stop
        Dim percentChange As Double ' to hold percent change for ticker
        
        ' Set Header Values
        wSheet.Cells(1, 9).Value = "Ticker"
        wSheet.Cells(1, 10).Value = "Yearly Change"
        wSheet.Cells(1, 11).Value = "Percent Change"
        wSheet.Cells(1, 12).Value = "Total Stock Volume"
        wSheet.Cells(1, 16).Value = "Ticker"
        wSheet.Cells(1, 17).Value = "Value"
        
        ' Initialize start values
        lastRowA = wSheet.Cells(Rows.Count, 1).End(xlUp).Row ' set last row
        startTickerRow = 2 ' Starting from 2nd row
        recordTickerRow = 2 ' Recordng value of first ticker on 2nd row
        
        ' Loop through all rows
        For currentRow = 2 To lastRowA
            
            ' when the following row ticker symbol does not match  currentRow ticker symbol -> as in they change then record the
            If wSheet.Cells(currentRow + 1, 1).Value <> wSheet.Cells(currentRow, 1).Value Then

                wSheet.Cells(recordTickerRow, 9).Value = wSheet.Cells(currentRow, 1).Value ' record the name of ticker to column J (#9) (Ticker)

                ' record yearly change in column I (#10) (Yearly Change)
                ' which is the currentRow (the last value of the current ticker we are looking at) - startTickerRow (opening value of the current ticker we are looking at)
                wSheet.Cells(recordTickerRow, 10).Value = wSheet.Cells(currentRow, 6).Value - wSheet.Cells(startTickerRow, 3).Value

                ' Setting Yearly Change cell colour
                If wSheet.Cells(recordTickerRow, 10).Value < 0 Then
                    wSheet.Cells(recordTickerRow, 10).Interior.ColorIndex = 3 ' Set cell colour to red
                Else
                    wSheet.Cells(recordTickerRow, 10).Interior.ColorIndex = 4 ' Set call colour to green
                End If
                
                ' Calculate and record percent change in column K (#11), condition is needed to valud dividing by 0 error
                If wSheet.Cells(startTickerRow, 3).Value <> 0 Then
                    PerChange = ((wSheet.Cells(currentRow, 6).Value - wSheet.Cells(startTickerRow, 3).Value) / wSheet.Cells(startTickerRow, 3).Value)
                    wSheet.Cells(recordTickerRow, 11).Value = Format(PerChange, "Percent") ' formating to percentage
                Else
                    wSheet.Cells(recordTickerRow, 11).Value = Format(0, "Percent")
                End If

                ' Calculate and record total volume in column L (#12)
                ' By getting the sum of the of closing prices from the first day to the last
                wSheet.Cells(recordTickerRow, 12).Value = WorksheetFunction.Sum(Range(wSheet.Cells(startTickerRow, 7), wSheet.Cells(currentRow, 7)))
                
                recordTickerRow = recordTickerRow + 1 ' go to next row to calculate and record the next ticker values
                
                startTickerRow = currentRow + 1 ' next ticker symbol is in the following currentRow

            End If
            
        Next currentRow
        

        ' TICKER SUMMARY

        Dim GreatestIncrease As Double ' for greatest increase calculation
        Dim GreatestDecrease As Double ' for greatest decrease calculcation
        Dim GreatestVolume As Double ' for greatest volumn calculation

        ' Row headings
        wSheet.Cells(2, 15).Value = "Greatest % Increase"
        wSheet.Cells(3, 15).Value = "Greatest % Decrease"
        wSheet.Cells(4, 15).Value = "Greatest Total Volume"

         'Find last non-blank cell in column I
        LastRowI = wSheet.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Prepare for summary
        GreatestVolume = wSheet.Cells(2, 12).Value
        GreatestIncrease = wSheet.Cells(2, 11).Value
        GreatestDecrease = wSheet.Cells(2, 11).Value

        For currentRow = 2 To LastRowI
        
            ' Calculation for greatest total volume
            ' if next value is larger then update greatest total volume
            If wSheet.Cells(currentRow, 12).Value > GreatestVolume Then
                GreatestVolume = wSheet.Cells(currentRow, 12).Value ' update greatest volume
                wSheet.Cells(4, 16).Value = wSheet.Cells(currentRow, 9).Value ' record ticker value
            End If
            
            ' Calculation for greatest increase
            ' if next value is has larger then update greatest increase
            If wSheet.Cells(currentRow, 11).Value > GreatestIncrease Then
                GreatestIncrease = wSheet.Cells(currentRow, 11).Value ' update greatest increase
                wSheet.Cells(2, 16).Value = wSheet.Cells(currentRow, 9).Value ' record ticker value
            End If
            
            ' Calculation for greatest decrease
            ' if next value is smaller then update greated decrease
            If wSheet.Cells(currentRow, 11).Value < GreatestDecrease Then
                GreatestDecrease = wSheet.Cells(currentRow, 11).Value ' update greatest decrease
                wSheet.Cells(3, 16).Value = wSheet.Cells(currentRow, 9).Value ' record ticker value
            End If
            
        ' Record summary results
        wSheet.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
        wSheet.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
        wSheet.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")
        
        Next currentRow

        'Resize column widths automatically content for better visual
        Dim WorksheetName As String
         'Get the WorksheetName
        WorksheetName = wSheet.Name
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
    Next wSheet
End Sub



