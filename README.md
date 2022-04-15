# Green Energy Stock Analysis with VBA in Excel

## Overview of Project

### Purpose
Steve wanted to analyze the stock of a company called DAQO New Energy Corp (DQ) because his parents were passionate about green energy and wanted to invest all their money into it. Steve asked our help on an excel he created for analysis of both DAQO's stock and all other green energy stocks.  We used VBA scriopt and collected 12 different green enegry companies' stock performance in the year 2017 and 2018, including DAQO's. The purpose of this project is to refactor this VBA script to increase the efficiency of the original code so it takes less time for Steve to run through all stocks information and pull out the information he wants.

## Results

- The refactored code did increase the efficiency by reduce the run time of the code. The original code run time for 2017 and 2018 were 0,605 seconds and 0.578+ seconds. The refactored one only takes 0.121 seconds for 2017 and 0.133 for 2018.

The original code was using nested for loops and print out the matched tickers. The outer loop run through the 12 tickers and the inner loop

![2017refactor](VBA_Challenge_2017.png)

## Summary
### Pros and Cons of Refactoring Code
- Refactoring Code in general gives a second thought to the original code we write. It will increase the efficiency of our code, refine logic loopholes and make the script look clean and organized. This not only helps the writer understand his code better in the future, but also benefits the users to run  programs with less time and computer energy.
- However, every coin has two sides. Refactoring the original code, which has been proved to be working, can bring in new bugs or conflicts with other programes in the main branch system.

### Pros and Cons of of the original All Stock Analysis VBA Script
### Pros and Cons of of the refactored All Stock Analysis VBA Script
