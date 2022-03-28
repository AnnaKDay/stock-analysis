# stock-analysis
 Conducting stock analysis on alternative energy related industries

## Overview
The original analysis was conducted in order to help Steve understand at a glance what stocks his parents should invest in, and which ones to avoid.The refactored analysis is a test to see if editing the code to be more compact would successfully increase efficiency. This is so Steve can see that there are verified ways of optimizing the code, as he intends to apply this code to a larger dataset of the stock market and wants the code to run as quickly as possible, given the increased volume of data. 
## Analysis and Results
In order to increase the efficiency (aka decrease run time) of our code, we created a variable named "tickerIndex," that is initialized to zero, and will serve as a index reference for the original input array of tickers, as well as three output tickers that we made for the refactoring process. The three output arrays are as follows: "tickerVolumes", "tickerStartingPrices", and "tickerEndingPrices". 
```
'Create a ticker Index set equal to 0 before iterating through the rows
Dim tickerIndex As Single
tickerIndex = 0

'Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```
We initialized "tickerVolumes" to zero within its own ```for``` loop, effectively setting each row's initial volume count to zero so that it could be counted in our next for loop, which will be explained next. While we did not initialize "tickerStartingPrices", and "tickerEndingPrices," this is also an option.
```
'Create a for loop to initialize the tickerVolumes to zero
For i = 0 To 11
    tickerVolumes(i) = 0
Next i
```
We then started a new ```for``` loop and specified three ```if``` statements to handle the counters for "tickerVolumes", "tickerStartingPrice", and "tickerEndingPrice", using our "tickerIndex" variable (which is set to increase at the end of each iteration to change the applicable ticker), we iterated through each ticker in the dataset and consolidated their total volume, their starting prices, and ending prices without using the nested ```for``` loop we used in the original code.
```
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
```
Finally, we used some simple Cells().Value code to populate the consolidated data into a new sheet. This was followed by a calculation of the returns (in percentage) that each ticker experienced in 2017 or 2018. This was followed by some formatting to make the visualization more impactful. 
```
'Loop through your arrays to output the Ticker, Total Daily Volume, and Returns
For i = 0 To 11
'Activate which worksheet to put analysis in
    Worksheets("All Stocks Analysis").Activate
    'script to output in the right columns and rows
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
Next i
```
## Summary

### Pros and Cons of Refactoring

### Pros and Cons of Refactoring our Green Stocks Analysis
