# Stock-Analysis

## Overview of Project
Steve is a recent finance graduate whose parents have invested in DAQO New Energy Corporation stocks. The goal of this project was to measure the stocks performance as well as to look into other clean energy stocks to diversify their portfolio.

I analyzed the stocks based on their past performance during 2017 and 2018. These were the metrics I used to measure performance:    
  - How actively were they traded in the past?
  - What were their yearly returns? (percentage change in price over 1 year.)


## Results
I used VBA in excel to calculate trading activity and yearly returns. The dataset I used included two sheets for the years 2017 and 2018. The data included 12 clean energy stocks and was organized by categories such as ticker name, date, closing amount and volume.
I quantified the performance of the stocks by:
  - Calculating trading activity for each stock (ticker) by adding their daily volume amounts.
  - Calculating yearly returns for each stock(ticker) by calculating the percentage change in price over the year. I used the starting and ending closing price for this calculation.

### 2017 performance

**Trading Activity**
  - DAQO stock (DQ) had the lowest trading volume among all the green energy stocks as shown below (The stocks are sorted by highest to lowest trading value):
  <img src ="https://github.com/Kee2u/Stock-Analysis/blob/main/README%20Images/2017%20Stock%20Performance%20by%20daily%20volume.png?raw=true" width = "400">

**Yearly Return**
  - However, DAQO had the highest yearly return (The stocks are sorted by highest to lowest return):
  <img src ="https://github.com/Kee2u/Stock-Analysis/blob/main/README%20Images/2017%20Stock%20Performance%20sorted%20by%20return.png?raw=true" width = "400">

Overall 2017 was a good year for clean energy with most stocks exhibiting a positive yearly return.

Note that these results show that trading volume is not a good indication of the yearly return of a stock. A stock's high trading volume doesn't imply a good return. TERP's trading volume was higher than DQ's but it had a negative return.

### 2018 performance

**Trading Activity**
  - DAQO stock had the third lowest trading volume as shown below (The stocks are sorted by highest to lowest trading value):
  <img src ="https://github.com/Kee2u/Stock-Analysis/blob/main/README%20Images/2018%20Stock%20Performance%20by%20daily%20volume.png?raw=true" width = "400">

**Yearly Return**
  - This year, DQ stock had the lowest return (The stocks are sorted by highest to lowest return):
  <img src ="https://github.com/Kee2u/Stock-Analysis/blob/main/README%20Images/2018%20Stock%20Performance%20by%20return.png?raw=true" width = "400">

Overall 2018 was a bad year for clear energy stocks with most stocks exhibiting a negative return.

These results also show that past performance does not imply future success. Many stocks did well in 2017 but did poorly in 2018.

### Original Code

Initially, I approached this problem by creating nested for loops to go through all the rows in the table and perform the calculations for one ticker. Then it updated the ticker and went through all the rows again. 

**Here is the code I used. The values of the ticker array were initialized with the values of all 12 tickers before these lines:**
  
          'Loop through the tickers
             
             For i = 0 To 11
    
                 ticker = tickers(i)
                 Totalvolume = 0
                 
                 'Loop through rows in the table
          
                 Worksheets(yearvalue).Activate
            
                        For j = 2 To RowCount
            
                            'Find the total volume for the current ticker.
                             
                             If Cells(j, 1).Value = ticker Then
                
                                  Totalvolume = Totalvolume + Cells(j, 8).Value
                
                             End If
                
                             'Find the starting price for the current ticker.
               
                              If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                   
                              startprice = Cells(j, 6).Value
                
                              End If
                
                              'Find the ending price for the current ticker.
               
                               If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                   
                               endprice = Cells(j, 6).Value
                
                               End If
            
                           Next j
            
              'Output the data for the current ticker.
              Worksheets("All Stocks Analysis").Activate
              Cells(4 + i, 1).Value = ticker
              Cells(4 + i, 2).Value = Totalvolume
              Cells(4 + i, 3).Value = endprice / startprice - 1
              
          Next i

**Computation Time**
  - The computation time for this code was 0.99s for 2017 and 0.81s for 2018:

<img src = "https://github.com/Kee2u/Stock-Analysis/blob/main/README%20Images/Green_Stocks_2017.PNG?raw=true" width = "300">   <img src = "https://github.com/Kee2u/Stock-Analysis/blob/main/README%20Images/Green_Stocks_2018.PNG?raw=true" width = "283">

### Refactored Code  

I then refactored the code to make it more efficient. Previously, the code went through all the rows 12 times for each ticker. This time the code ran through the rows only once and updated the ticker as it went along. 
I did this by updating the ticker once the code detected the ending price of the current ticker. This was possible because the rows were sorted chronologically: The ending price of one ticker was immediately before the starting price of the next as the code looped through the rows.

In addition to creating an array for the ticker index like before, I created arrays for ticker volumes, starting price and ending price.

**Here is the code I used. The values of the ticker array were initialized with the values of all 12 tickers before these lines:**
  
    'Create a ticker Index
    
    tickerindex = 0
    
    'Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEnding(12) As Single
    
    'Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    'Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount

        'Increase volume for current ticker
        
        If Cells(i, 1).Value = tickers(tickerindex) Then
        
            tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        'Check if the current row is the first row with the selected tickerIndex.
        
             If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            
                 tickerStartingPrices(tickerindex) = Cells(i, 6).Value
            
             End If
        
        'Check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
         
             If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
                 tickerEnding(tickerindex) = Cells(i, 6).Value
                 
                 'Increase the tickerIndex.
                 
                 tickerindex = tickerindex + 1
             
             End If
        End If
    
    Next i
    
    'Loop through the arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEnding(i) / tickerStartingPrices(i) - 1
        
    Next i
    
**Computation Time**
  - The computation time for this code was 0.42s for 2017 and 0.40s for 2018. The refactored code is faster than the original code.:

<img src = "https://github.com/Kee2u/Stock-Analysis/blob/main/resources/VBA_Challenge_2017.png?raw=true" width = "300">   <img src = "https://github.com/Kee2u/Stock-Analysis/blob/main/resources/VBA_Challenge_2018.png?raw=true" width = "290">

## Code Summary

Refactoring code restructures and optimizes code without changing its behaviour.

Depending on the approach, its advantages can be:
 - To improve Maintainability
 - To increase performance by decreasing computation speed
 - To create extensible code (makes it easier to add more functions)
 
It disadvantages are:
 - It is time consuming because it entails reworking existing code
 - It may introduce code bugs
 
 In our case, the advantage of refactoring the code is decreasing computational speed. 
 The disadvantage is that it consumed more memory by storing more data in arrays than the original code. In contrast, the original code used variables that updated their value. 
