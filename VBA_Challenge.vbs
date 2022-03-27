Sub AllStocksAnalysisRefactored()
' Declaring start time and end time variables in order to start and end timer later
Dim startTime As Single
Dim endTime As Single

' provides an Input Box, whereupon the inputter types what year they want to run analysis on, and that input replaces "yearValue" variable
yearValue = InputBox("What year would you like to run the analysis on?")

' set start time, "Timer" is actually a function within VBA, so our variable startTime now represents this function
startTime = Timer

'Format the Output sheet on All Stocks Analysis worksheet, this is where all our code will put what we want
Worksheets("All Stocks Analysis").Activate

' Sets A1 to display "All Stocks (yearValue)"
Range("A1").Value = "All Stocks(" + yearValue + ")"

'Create a Header Row
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

'Initialize array of all tickers, in this tickers(x), x = total count of indexes, then in the array, tickers(n) n = each index
Dim tickers(12) As String

tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"

'Activate data worksheet to pull data from
Worksheets(yearValue).Activate

'Get the number of rows to loop over
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'Create a ticker Index set equal to 0 before iterating through the rows
Dim tickerIndex As Single
tickerIndex = 0

'Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

'Create a for loop to initialize the tickerVolumes to zero
For i = 0 To 11
    tickerVolumes(i) = 0
Next i
'Loop over all the rows in the spreadsheet

For i = 2 To RowCount

    'Increase the Volume for current ticker
    If Cells(i, 1).Value = tickerIndex Then
        
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    End If
    'Check if the current row is the first row that contains that ticker
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If

    'Check if the current row is the last row with the selected Ticker
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
    'If the next row's ticker doesn't match, increase the tickerIndex

        tickerIndex = tickerIndex + 1
        
    End If

Next i

'Loop through your arrays to output the Ticker, Total Daily Volume, and Returns
For i = 0 To 11
'Activate which worksheet to put analysis in
    Worksheets("All Stocks Analysis").Activate
    'script to output in the right columns and rows
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
Next i

'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
