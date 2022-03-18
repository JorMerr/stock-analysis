# stock-analysis
# Overview of Project

The purpose of this analysis is to determine whether a completed VBA code loop will scale effectively to larger data sets. The initial Macro met expectations when analyzing 12 different stock tickers for Total Daily Volume for a given year, and the resulting Annual Return value. 

The VBA code loop has been refactored in this project to analyze a greater number of stock tickers for the same values of Total Daily Volume in a given year, and the Annual Return for the same year. Using the same dataset as the initial code loop, our comparison will be made based on code performance with a built-in timer to determine how long each code block takes in order to complete the loops.


---
# Results: 

## Comparing Stock Performance

Through analysis of the years 2017 and 2018, we can see a average trend of better performance in 2017 for the stock tickers included in our dataset. The notable exception to this trend is stock ticker "RUN", which resulted in a Total Annual Return for 2017 of 5.5%, while the Total Annual Return for stock ticker "Run" in the year 2018 resulted in an 84.0% increase in value.

![Stock Performance 2017](https://github.com/JorMerr/stock-analysis/blob/main/Resources/Stock_Performance_2017.PNG)

![Stock Performance 2018](https://github.com/JorMerr/stock-analysis/blob/main/Resources/Stock_Performance_2018.PNG)


It is unclear from our dataset whether this trend extends to stock tickers throughout the market, or if there was an average downtrend in stock market performance for the year 2018. With our refactored VBA script, we may be able to quickly make minor adjustments to our `tickers` array and the array length of our `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` in order to quickly analyze performance of other stock tickers.


## Execution times of Original Script vs Refactored Script
The original code was used to loop through the data several times in order to collect the information required for each stock ticker. The completed original code, including the macro used to format cells is included below:

```
Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer
'1) Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    ' Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'2) Initialize an array of all tickers.
    Dim tickers(11) As String
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
    
'3a) Initialize variables for the starting price and ending price.
    Dim startingPrice As Double
    Dim endingPrice As Double
    
'3b) Activate the data worksheet.
    Worksheets(yearValue).Activate
    
'3c) Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'4) Loop through the tickers with initial volume set to 0
    For i = 0 To 11
        Ticker = tickers(i)
        totalVolume = 0
        '5) Loop through rows in the data.
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
        '5a) Find the total volume for the current ticker.
            If Cells(j, 1).Value = Ticker Then
        
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
        
        '5b) Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
        
                startingPrice = Cells(j, 6).Value
            End If
        '5c) Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
            
                endingPrice = Cells(j, 6).Value
            End If
        Next j
         
    '6) Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = Ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
    endTime = Timer
    MsgBox "this code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub

Sub formatAllStocksAnalysis()
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.Color = RGB(0, 100, 180)
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
    'Conditional Formatting
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed
        Else
            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone
        End If
    Next i
End Sub
```

The refactored code allowed for us to loop through to code a single time, saving information to output arrays which were updated based on the `tickerIndex` created as each row was analyzed. Relevant portions of refactored code are as below:

```
'1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For v = 0 To 11
        tickerVolumes(v) = 0
        tickerStartingPrices(v) = 0
        tickerEndingPrices(v) = 0
    Next v
        
    '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
            
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            End If
            
        Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
```

As can be seen when comparing the refactored code, we used the `tickerIndex` variable as a means of storing information for each string of the `tickers` array. By storing the information to `tickerIndex` this allowed for a single loop of the information to be completed, greatly improving the performance of our VBA Script. We can see the increase in performance when comparing the original script performance to the refactored script performance as outlined in the images below.


![Refactored Code runtime for 2017:](https://github.com/JorMerr/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

![Original Script runtime for 2017:](https://github.com/JorMerr/stock-analysis/blob/main/Resources/VBA_Challenge_original_2017.PNG)



![Refactored Code runtime for 2018:](https://github.com/JorMerr/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

![Original Script runtime for 2018:](https://github.com/JorMerr/stock-analysis/blob/main/Resources/VBA_Challenge_original_2018.PNG)


---

# Summary:

## What are the advantages or disadvantages of refactoring code?

The primary advantage of refactoring code is that the code can be made to run more efficiently. The code may not introduce any new functionality, but refactoring code will allow a faster loop, or a lower memory utilization. This increase in efficiency can be especially important when dealing with larger datasets.
A main disadvantage of refactoring code is that it can be difficult to ensure the code continues to work as intended without introducing errors or inconsistent syntax.

## How do these pros and cons apply to refactoring the original VBA script?

The advantage of refactoring the code in this exercise is to allow for improved efficiencies with a larger dataset than the original VBA script provided. As seen above, the refactored script for each of the years 2017 and 2018 complete in less than one fifth (1/5^th, or 0.20) seconds. The original script ran at greater than one (1.00) second for 2017 and 2018 years respectively.

As can be clearly seen when comparing performance, the refactored code operates at an efficiency greater than five times the speed of the original code. The refactored code also completed the formatting of cells as a component of the script, rather than having a separate subroutine.

The disadvantage of completing the refactorting in this exercise was the time taken to debug script when errors would occur. At times in the completion of refactoring the code as instructed, I had encountered several errors and had to debug them before moving on to the next step.

The debugging process was the most time intensive portion of this exercise. In an employment setting there may be some consideration taken to determine whether refactoring the working VBA script to run more efficiently will be worth the time taken for employees to complete the refactor.