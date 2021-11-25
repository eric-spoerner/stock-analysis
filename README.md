# Stock Analysis Tool and Refactor

## Purpose

The purpose of this item is a comparative analysis of a tool designed in Excel with VBA macros for analyzing and processing stock ticker data and calculating annual metrics on total trading volume and annual return.  Intent is to demonstrate the impact of a refactor in significantly reducing burden in a small data set before scaling to accomodate larger data sets.  The original design, while accurate, already showed signs of slowness in a small data set.

## Tool

The tool is built to internalize tabular data featuring the following data points:
* Stock ticker
* Trading Date
* Daily prices - Open/High/Low/Closing/AdjustedClosing
* Trading volume

Tabular data is stored in Excel spreadsheet (vba_challenge.xlsm), with data being separated by year with a corresponding name in the .  Macros attached to the spreadsheet are built to:
1. Prompt user for year input to determine data set for analysis
2. Identify all tickers in data set (currently hardcoded to anticipate 12 tickers)
3. Iterate over full data set and identify the following attributes per ticker:
    1. Starting price
    2. Ending Price
    3. Aggregate trade volume
4. Calculate annual return by comparing start and end price
5. Write ticker name, aggregate trade volume, and annual return for each ticker to the All Stocks Analysis sheet in the same workbook
6. Apply conditional formatting for legibility

## Design

### Original

Original design is contained in the Module 1 subroutine `yearValueAnalysis()`.  The below snippet contains the original design for the actual process of iterating through each ticker.  Design iterates for each ticker in the `tickers()` array, which contains 12 values.

```
    For x = 0 To 11
        Worksheets("2018").Activate
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

Original design is contained in the Module 2 subroutine `AllStocksAnalysisRefactored()`.

## Refactor analysis

**ES NOTE: MAKE SURE TO DO SUB ANALYSIS OF CORE LOOP TO SEE WHAT COMPONENT OF THE FULL REFACTOR IS THIS LOOP**

## Future work