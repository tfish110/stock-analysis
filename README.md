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

### Tables of results

For 2017, only one of the selected stocks returned a loss in value, TERP. But, it seems that 2018 must have been a rough year for the green energy sector, as that trend nearly reversed and only two of the selected stocks return a gain in value, ENPH and RUN. The way that this information might influence Steve and his parents' choices for diversifying their investments would depend largely on their investment strategies and risk tolerence. All we can say for sure is that ENPH is the only stock that returned gains for both years, and TERP is the only stock that returned losses for both years. The main advantage that we gain in this analysis is the ease with which we can quickly get a sense for the directionality and magnitude of the stocks' changes in value each year. The tables for both years are displayed here:

![Table_2017](https://github.com/tfish110/stock-analysis/blob/main/Resources/Table_2017.png)

![Table_2018](https://github.com/tfish110/stock-analysis/blob/main/Resources/Table_2018.png)

### Execution times

By using the 'Timer' function in VBA, we were able to see how long it took for each of our scripts to run and fill in the tables discussed above. We then used the 'MsgBox' function to display the time it took to run the code after it completed. For the original code, these were the resulting times that were displayed for each year:

![Original_2017_Time](https://github.com/tfish110/stock-analysis/blob/main/Resources/Original_Code_2017.png)

![Original_2018_Time](https://github.com/tfish110/stock-analysis/blob/main/Resources/Original_Code_2018.png)

We included this same functionality for timing the code when we refactored the original code. These new, faster times can be seen here:

![Refactored_2017_Time](https://github.com/tfish110/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![Refactored_2018_Time](https://github.com/tfish110/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary
### What are the advantages or disadvantages of refactoring code?

Refactoring code can have some clear benefits when it comes to improving the efficiency and speed that it takes to run a script. However, there are some pitfalls which could make the process more of a burden than a benefit. If you end up making mistakes, you may break code that was functioning perfectly fine before. There is also a possibility that your refactored solution might simply be a different way of running a script, without actually making any efficiency improvements at all. You may also end up impacting the readability of the code in either direction; different approaches to refactoring could be either a benefit or a hindrance in that respect, especially if there are multiple people working on the code who need to work together as a team.

### How do these pros and cons apply to refactoring the original VBA script?

In this case, there was a clear advantage in the speed and efficiency of the refactored script. While both scripts ran in less than a second, the refactored script was still about 8 times faster than the original. With this volume of data, that difference wouldn't make much of an impact, but it is a good demonstration of how much of an impact more efficient code can have when working with larger datasets. As for the pitfalls, I must admit that I did run into a few when refactoring my code. I knew that it would be useful to copy and paste some code from the original script to make some minor changes for the new version. However, I ended up getting confused at times, and needed to backtrack to undo some changes. Even when I thought that I had gotten to the end of the script successfully the first time, it turned out I had made some mistakes and my code would not run. I tried debugging, which resulted in making the problems worse it seemed, and so I ended up cutting it all out and starting over again from the beginning. This was frustrating, but it was a good learning opportunity. By going through each step a second time, I had a better understanding of what I needed to do along the way because of the mistakes that I made the first time.
