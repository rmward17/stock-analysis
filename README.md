# stock-analysis
## Overview
Overview of the prject, explain the purpose of this analysis
Steve came to me with Green Energy stocks and their prices for 2017 and 2018 so that we could figure out analyze them in order to help diversify his parent's portfolio. With the stock data, we looked at the total daily colume of the stocks, their starting prices for the year, ending prices for the year, and their return for the year. Using this data, we are able to give his parents some advice on what stocks they can invest in and set up this worksheet to get that same information if they were to add data for these stocks from other years.

We ran this analysis using VBA in excel and had a couple different ways of doing that. 
## Analysis
To being, we added a new sheet just to run an analysis on the "DQ" stock. We created a header row with labels for the year, for the total daily volume, and the yearly return.
Using a for loop, we are able to run through the rows on the sheet with the 2018 data. The loop goes through each ticker to look for "DQ". When a ticker does have "DQ", the loop increased the total daily volume. The loop also has conditionals in it to determine what the starting and ending prices for the year were. The conditional checks the paramter and if it is true, it stores the correct value. For example, here is the conidtional fo rthe starting price:

    If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
        startingPrice = Cells(i, 6).Value
    End If

This conditional looks at the cell in the ticker column and if the ticker in the cell above it is different, then that means it is the starting price for the year. 
Once all of those values are stored, we wrote code to output the year, the total daily colume and were able to calculate the yearly return with the following formula

    Ending Price / Starting Price - 1

Using that code to start, we refactored it to be more data driven, dynamic, and user friendly! the first change we made was with the addition of an Input Box. When you runthe macro, you are prompted to input a year. Once you put the year in, the year is assigned to a variable that is used within the code. This makes it so that if you want to change the year that the analysis is run on, you can rerun the macro, change the year. and be done! 

We also created an **array**. An array is a list of data and the index for an array starts at 0. So, for our data, we have 12 tickers so they are indexed 0 to 11 rather than 1 to 12. In the code, we created the array for the tickers and gave each index a value.

        Dim tickers(12) As String
    
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
        
Here you can see that we filled up our array with the ticker abbreviations and assgined them to an index value. We also created arrays for the total daily columes, starting prices, and ending prices to hold those values for each ticker. Now we can use a for loop to assign the total daily volume, starting price, and ending price for each ticker using the index! In the code below you can see the loop as well as the conditionals I mentioned before being used to run through the data and pull what we need into the arrays.

    'this is saying "for each row, do the following" and once the action is complete, it will move onto the next row until it has reached all of the rows
    For i = 2 To RowCount 
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value 'increases the Total Daily Volume for each ticker
        
        'Checks to see if the ticker above the current row is different. If it is, then we store that closing price as the starting price for that ticker
        If Cells(i, 1).Value <> Cells(i - 1, 1) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value   
        End If
        
        'Checks to see if the ticker below the current row is different. If it is, then we store that closing price as the ending price for that ticker
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
       
       'Checks to see if this is the last row for the ticker we are on. If it is, we increase our ticker index by 1 so that we can move on and do the loop for the next ticker in the list.
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerIndex = tickerIndex + 1
         End If
    
    Next i

This code did all f the work for us and stored everythign into it's arrays and essentially created lists of the data for each ticker by connecting them via index. We were also able to do formatting in VBA as well so it is easier to look at the results. 

## Results


## Summary
1. What are advantages or disadvantages of refactoring code
2. 2. How do these pros and cons apply to refactoring the original VBA script
