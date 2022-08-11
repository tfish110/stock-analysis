# stock-analysis
Using VBA and Excel to analyze stock data over two years

## Overview

Steve wishes to help his parents with their investments by examining a selection of green energy stocks, a business sector they are passionate about. They currently own shares of a company called DAQO, but Steve thinks it would be wise for them to diversify their investments. By analyzing other green energy stocks and comparing their performance against DAQO for the years 2017-2018, we can make some recommendations about which stocks might be good choices for their investment. We were able to use VBA in Excel to quickly and efficiently transform a large quantity of daily stock ticker data for our stocks of interest from those two years into a table summarizing how the stock performed that year.

## Results
### Different coding approaches

For this analysis, we used two different methods of VBA scripting to get the same results. In both methods, the code for building the structure and formatting for the results table remained the same. However, the methodology for drawing information from the raw data into our table was different. In our first coding methodology, the modules from the class materials walked us through using nested for loops to identify which lines of data belonged to each stock ticker (outer loop). We then calculated the total daily volume and annual return for each stock (inner loop):

'''

    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
        Next j
            
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
'''

For this challenge, we were tasked with refactoring this original code into a more efficient format so that the program runs quicker. We were given some starting points and hints for the steps along the way, but had to figure out the correct way to code each of the hints we were given. This refactored code used an additional "Index" variable that we created and a series of arrays to structure the data in the results table. This method instead had us build entirely separate loops rather than nesting one loop within another:

'''

    Dim tickerIndex As Integer
    tickerIndex = 0

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    For i = 2 To RowCount
    
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
       
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
          
    Next i

'''

