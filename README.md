# stock-analysis

## Overview of Project

The purpose of this project is to refactor code to analyze stock data from all stock information provided from 2017 and 2018. The analysis will compare the performance between 2017 and 2018. It will also show how this the refactored code differs from the original code and provide run timestamps. Lastly the pros and cons of refactoring original VBA scripts will be explained, along with detailing advantages and disadvantages to refactoring code.

## Analysis

The first portion of the Analysis provided will be comparing the performance between the stocks provided in 2017 and 2018.Below the data and visualization charts are displayed.

 <img src="Resources\AllStockComparison.png" style="zoom: 80%;" />
 
 Above is a comparison of a complete analysis of all stocks by year. The "Total Daily Volume" is the number of trades of the stock in a given day. In the tables the value represented for each stock is a sum of the daily volue, or number of trades in a day. In the "Return" column, the stock's price at the end of the year is divided by the price at the beginning of the year, and converted to show percentage growth or loss. This indicates how much return on investment a stock recieved, with positive (green) values indicating increased value and negative (red) indicating losses. The color coding od the return percentages is achieved with conditional formmating.
 
 Conditional formatting allows the differences between stock performance in 2017 and 2018 to be displayed easily and clearly. In 2017,most of the stocks have a positve return, indicated by the green, while the ticker ("TERP") shows a negative return (colored red). Whereas, in 2018 most of the selected tickers showed a negative return except for "ENPH" and "RUN", both colored green.
 
The **All Stock Analysis** was able to be executed by code refactoring. The initial macro was written using a nested for loop, an example is available below:

 'initiate ticker loop and totalVolume
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
       
            'activate data worksheet
            Worksheets(yearValue).Activate
            
            'initiate nested row loop
            For j = 2 To RowCount
            
                'find total volume for current ticker
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                
                End If
                
                
The way a nested for loop functions is the computer will run through each "data row" once for each ticker category. Although nested for loops can run through any "data set" this can take the computer a longer time to complete especially if the data set is longer. Therefore, creating an index and using an array will save time. By creating this the computer is now only have to go through each data row once. By sorting the data for the computer beforehand, the computer recognizes that is has checked for that specific identify and does not need to go through the rows and check again. Therefore, it cuts down on the computer execution time.

 
 <img src="Resources\VBA_Challenge_2018.png" style=";" /> 
 
 
 
 
 
