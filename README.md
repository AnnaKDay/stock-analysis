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
![Average Efficiency](https://user-images.githubusercontent.com/100614266/160314127-0df2d67e-fcdc-4644-ae82-2783d029e223.png)
After testing the original code and refactored code 10 times for each year in the workbook, we found that while there are small discrepancies in each instance of the code running, the average run time for the original code in both years was 1.06 seconds. Likewise, for the refactored code, we found that the code ran in 0.125 seconds for both years. This results in an 88% increase in effienciency (or 88% decrease in runtime). 
![2017 Original](https://user-images.githubusercontent.com/100614266/160314314-d05481ca-6ac1-481b-b610-285716691cc9.png)
![2017 Refactor ](https://user-images.githubusercontent.com/100614266/160314319-5554c136-b0eb-476f-b782-cd454ec64087.png)
![2018 Original](https://user-images.githubusercontent.com/100614266/160314324-da7eacc0-40c3-471c-8412-12c97f3de6cc.png)
![2018 Refactor](https://user-images.githubusercontent.com/100614266/160314326-ea10f5ef-e1e9-4d05-b662-684b628b3983.png)

### Pros and Cons of Refactoring
Refactoring is a process that takes a drafted, or even completed code, and edits it to optimize runtime and performance. So, the obvious advantage is the running of cleaner, more efficient code. However, this "cleaner" code is also less understandable to any newcomer or person who is unfamiliar with the code. It can be quite difficult to understand what the original coder was doing, and how, which are important considerations for anyone trying to work with that code going forward. Furthermore, the process of refactoring can sometimes lead to accidentally breaking the code and rendering it inoperable. This is extremely annoying to debug and fix if you lack a backup of the original working code. In our case, spending 3 hours to gain less than a second of efficiency is less useful. 
### Pros and Cons of Refactoring our Green Stocks Analysis
We proved that refactoring the code does indeed make the runtime faster, which has good implications for applying this code to a larger dataset of the stock market. That being said, it is uncertain if the 88% efficiency gained by refactoring will translate when the code has to handle a much larger volume of data. If the efficiency gained is marginal for what Steve wants to do, then this effort was not as helpful as it seems in the small scale. There are other factors to take into consideration as well, such as the hardware of whatever system is parsing the code. Some computers do not run as well as others, and this can influence runtime beyond the refactored code. Furthermore, since we only tested one refactoring method, it is uncertain if this is the best refactoring method, or if there are even better ways to optimize the code.
