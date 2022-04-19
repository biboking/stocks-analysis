# All Stocks Analysis with VBA in Excel

## Overview of Project

### Purpose
Steve wanted to analyze the stock of a company called DAQO New Energy Corp (DQ) because his parents were passionate about green energy and wanted to invest all their money into it. Steve asked for our help on an excel he created for analysis of both DAQO's stock and all other green energy stocks. We used the VBA script and collected 12 different green energy companies' stock performances in the year 2017 and 2018, including DAQO's. The purpose of this project is to refactor this VBA script to increase the efficiency of the original code so it takes less time for Steve to run through all stocks information and pull out the information he wants.

## Results

- The refactored code did increase the efficiency by reducing the run time of the code. The original code run time for 2017 and 2018 were 0.605 seconds and 0.578+ seconds. The refactored one only takes 0.121 seconds for 2017 and 0.133 for 2018.

- The original code used nested for loops and printed out the matched tickers. The outer loop ran through the 12 tickers and the inner loop through the rows. The refactored code, however, does not use nested for loops, but only several independent for loops. Two variables were used, i and j, referring to the rows and the 12 tickers. The code details are listed below.

![2017refactor](VBA_Challenge_2017.png)

![2017refactor](VBA_Challenge_2018.png)
   
    
     '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
        tickerVolumes(j) = 0
        tickerStartingPrices(j) = 0
        tickerEndingPrices(j) = 0
    Next j
        
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
      
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
        
        '3c) check if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
           
         '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
            
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For j = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + j, 1).Value = tickers(j)
        Cells(4 + j, 2).Value = tickerVolumes(j)
        Cells(4 + j, 3).Value = tickerEndingPrices(j) / tickerStartingPrices(j) - 1
    
    Next j



## Summary
### Pros and Cons of Refactoring Code In General

- Refactoring Code in general gives a second thought to the original code we write. It will increase the efficiency of our code, refine logic loopholes and make the script look clean and concise. This not only helps the writer understand his code better in the future but also benefits the users to run programs with less time and computer energy.

- However, every coin has two sides. Refactoring the original code, which has been proved to be working, can bring in new bugs or conflicts with other programs in the main branch system. It also sometimes requires better mathematics logic and higher coding skills to get a lean script.

### Pros and Cons of the original All Stock Analysis VBA Script

- The original All Stock Analysis VBA Script looks long and complicated, it has nested loops that require a longer time for the computer to process. It also could lead to more bugs because of the complication there. However, the pros are that it's easier for the writer to put the code down because it's the flow of the writer's thoughts. It might be tedious but it didn't require the writer to organize the codes, but only follow the writer's thoughts and write the script step by step.

### Pros and Cons of the refactored All Stock Analysis VBA Script

- The refactored script looks clean and concise. One of the biggest pros is that it significantly reduces the time for the computer to process the macro than the original one. The cons are also obvious. The writer needs to have a very clear picture of the whole script and think thoroughly before putting the code down. Otherwise, in the worst-case scenario, one careless move would forfeit the whole game.
