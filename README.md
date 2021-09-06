# Stock-analysis
week 2 module over VBA

## Overview: 
The goal of this analysis was to use excels macros developed through VBA, to provide and accurate analysis of the the difference is stock values over the course of 2017 and 2018 to provide the clients with a better view of how their DQ stocks are preforming. 

## Results: 
The results of the stocks analysis showed that the performance of all the stocks in 2017, except for RUN, decreased their returns over the course of 2018, and that the stock DQ, has a 262% decrease over the course of the year, the most significant decrease out of all the stock’s return. It is recommended with such a downturn in returns to withdrawl and buy stocks, but with better prospects to maximize returns.

2017 stock returns

![2017 stock returns](https://github.com/ChristopheGarcia1/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

2018 stock returns

![2018 stock returns](https://github.com/ChristopheGarcia1/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The results of the refactored all stocks analysis macro vs. the original showed a significant increase in speed by going from approximately 0.67 seconds per run in the original code to 0.14 seconds for the refactored code.

Original Macro time:

![Time of original Code](https://github.com/ChristopheGarcia1/Stock-analysis/blob/main/Resources/original_Time.png)

Refactored Macro time:

![Time of refactored code](https://github.com/ChristopheGarcia1/Stock-analysis/blob/main/Resources/Refactored_time.png)

This improvement in run time was accomplished by removing the nested for loops and instead using arrays to bypass the nested for loops.  

'''
       Refactored code
       
       For i = tickerIndex To 11
            ticker = tickers(i)
            tickerVolume(i) = 0
        Next i
    2b) Loop over all the rows in the spreadsheet.
            For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
                tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.

                If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                    tickerStartingPrices(tickerIndex) = Cells(j, 6)
                End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rows ticker doesnt match, increase the tickerIndex.
         
                If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                   tickerEndingPrices(tickerIndex) = Cells(j, 6)
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
            End If
    
    Next j
    
    '''
'''Original Code

    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        Sheets(yearValue).Activate
    For j = 2 To RowCount
    
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        End If
        
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            'set starting price
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

The result leads to less complicated processes and pushes the codes optimization significantly.  


## Advantages and disadvantages of refactoring: 

Refactoring is a crucial part of coding for the fact that an initial script or program will never be as optimized as it can be at its first draft. You have to set up the variables and the logic so it flows in a concise and straightforward manner. However, our logic can be flawed or inefficient so one of the main advantages of refactoring is that it offers a chance to refine the code to be more efficient and easier on the computer. Inefficient code can cause simple tasks to take a needlessly long amount of time and resources during the process. Refactoring allows for optimization, reducing run time and simplifying task. This can be seen in the  difference in all stocks analysis and all stocks analysis refactored code. This refactoring has also helped with the readability. 


A disadvantage of refactoring code is that changing the initial arguments can cause the original skeleton of the code to not function. Refactoring can mean reconstructing the script from scratch and making the original code obsolete. A majority of the disadvantages of coding revolve around the fact that a lot of its original functionality has to be reworked to accommodate the coding. This can be seen with removing the nest for loops and replacing them with arrays.

