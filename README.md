# Stock Analysis Tool - Data Analysis and Refactor

## Purpose

This repository is a comparative analysis of a tool designed in Excel with supporting VBA macros for consumption of stock ticker data and calculation of  annual metrics on total trading volume and annual return in 2017 and 2018.  Intent is to demonstrate the impact of a refactor in significantly reducing burden in a small data set before scaling to accomodate larger data sets.  The original design, while accurate, already showed signs of slowness in a small data set.

## Tool

The tool is built to consume tabular data featuring the following data points:
* Stock ticker
* Trading Date
* Daily prices - Open/High/Low/Closing/AdjustedClosing
* Trading volume

Tabular data is stored in Excel spreadsheet (vba_challenge.xlsm), with data being separated by year with a corresponding name for each worksheet.  Macros attached to the spreadsheet are built to:
1. Prompt user for year input to determine data set for analysis
2. Identify all tickers in data set (currently hardcoded to anticipate 12 tickers)
3. Iterate over full data set and identify the following attributes per ticker:
    * Starting price
    * Ending Price
    * Aggregate trade volume
4. Calculate annual return by comparing start and end price
5. Write ticker name, aggregate trade volume, and annual return for each ticker to the All Stocks Analysis sheet in the same workbook
6. Apply conditional formatting for legibility

## Design

### Original

Original design is contained in the Module 1 subroutine `allStockAnalysis()`.  The below snippet contains the original design for the actual process of iterating through each ticker.  Design iterates for each ticker in the `tickers()` array, which contains 12 separate stock tickers.  Within each ticker's respective for loop, each row is independently analyzed to see if it contains data related to that ticker, and collects the data if it is true.

```
For x = 0 To 11
    Worksheets(yearValue).Activate
    ticker = tickers(x)
    totalVolume = 0
    
    For y = rowStart To rowend
        If Cells(y, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(y, 8).Value
        End If
    
        If Cells(y, 1).Value = ticker And Cells(y - 1, 1).Value <> ticker Then
            startingPrice = Cells(y, 6).Value
        End If
        
        If Cells(y, 1).Value = ticker And Cells(y + 1, 1).Value <> ticker Then
            endingPrice = Cells(y, 6).Value
        End If
    Next y
            
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + x, 1).Value = ticker
    Cells(4 + x, 2).Value = totalVolume
    Cells(4 + x, 3).Value = (endingPrice / startingPrice) - 1
Next x
```

### Refactor

Refactored design is contained in the Module 2 subroutine `AllStocksAnalysisRefactored()`. Refactor eliminated nested for loops in favor of a single loop that iterated over every row only once, identfied whether that stock 

```
''2b) Loop over all the rows in the spreadsheet.
For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        '3d Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
    End If
        
Next i
```
## Analysis

### Results

![2017 good](resources/VBA_Challenge_2017.png)
![2017 bad](resources/VBA_Challenge_2018.png)

### Observations

ENPH stands out as a star performerin return. DQ saw most dramatic downward swing in return. Daily volume does not appear to have an impact on performance.

### Data limitations

Design of current framework is built on major assumptions about value of data, including:
* Design assumes sequential data, ordered by stock ticker and date.
* Design assumes a fixed quantity of stock tickers (12 total)

## Refactoring

### Benefits and risks

Refactoring of complex analytical calculations is often mandatory given a large or growing data set.  The larger the data's volume and the more complex the required calculations are, the more that inefficient code can contribute to slowness and even full process failure by way of timeouts.

In other cases, refactors can assist legibility and maintainability, and allow for expansion of existing functionality to include new use cases.

Primary risk of refactor is the addition of new unintended defects based on a failure to understand algorithm in full or misundertanding of how the current design affected edge casess.  Without sufficient testing in place, refactors can easily create new edge case failures etc.

Something about technical debt.

### Refactor Performance Testing Methodology

Time is recorded for full execution of both versions of the code using the VBA `Timer` function, which is called at the beginning and end of the subroutines to establish total duration.

Time for each was run three times, for both subroutines, with the year 2017 applied (operating assumption is that specified year does not matter as the data sets are roughly identicial).  See attached screenshots for sample output from before and after:

### Refactor Performance Testing Outcome

Reduction of for loops reduces number of row processing events by 91.67%.  
Assuming a linear value for processing, this should reduce overall burden of the code above by a similar percentage, and additionally 

## Limitations of current design / Opportunities for future refactors

Test framework is limited due to reliance on manual execution to test performance and subsequent small sample size.  For more accurate or robust data set, suggest building a data testing subroutine within the existing test framework by doing the following:
* Parameterize existing functions by adding**FIXME**
* Build for loops to test *x* years with *y* times repeated, and store the data from these runs on the same sheet by adding a print function to end of existing methodologies.

Scale this further with larger data set?  Assuming a linear progression, the original one would potentially take x seconds and the new one would take y seconds.