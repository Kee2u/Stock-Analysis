# Stock-Analysis

## Overview of Project
Steve is a recent finance graduate whose parents have invested in DAQO energy stock. The goal of this project was to see how well the stock compared to other clean energy stocks. 

I analyzed the stocks based on their past perfomance during 2017 and 2018. These were the metrics I used to measure performance:    
  - How actively were they traded in the past?
  - What were their yearly returns? (percentage change in price over 1 year.)


## Results
I used VBA in excel to calculate trading activity and yearly returns. The dataset I used included two sheets for the years 2017 and 2018. The data included 12 clean energy stocks and was organized by categories such as ticker name, date, closing amount and volume.
I quantified the performance of the stocks by:
  - Calculating trading activity for each stock (ticker) by adding their daily volume amounts.
  - Calculating yearly returns for each stock(ticker) by calculating the percentage change in price over the year. I used the starting and ending closing price for this calculation.

### 2017 performance

**Trading Activity**
DAQO stock had the lowest trading volume among all the green energy stocks as shown below:
<img src ="https://github.com/Kee2u/Stock-Analysis/blob/main/resources/2017%20Stock%20Performance%20by%20daily%20volume.png?raw=true" width = "400">

**Yearly Return**
However, DAQO had the highest yearly return:
<img src ="https://github.com/Kee2u/Stock-Analysis/blob/main/resources/2017%20Stock%20Performance%20sorted%20by%20return.png?raw=true" width = "400">

This shows that trading volume is not a good indication of a the yearly return of a stock. Just because a stock's trading volume is high doesnt mean its return will be good. TERP's trading volume was higher than DQ's but it had a negative return.


### Original Code

Initially, I approached this problem by making an array for the tickers ( tickers(12) ). Then I created nested for loops to first go through all the rows in the table and perform the calculations for one ticker. Then it went to the next ticker and went through all the rows again. 

**Here is the code I used. The values of the ticker array was initalized with the values of all 12 tickers before these lines:**
  
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
  The computation time for this code was 0.99s for 2017 and 0.8s for 2018:

<img src = "https://github.com/Kee2u/Stock-Analysis/blob/main/resources/Green_Stocks_2017.PNG?raw=true" width = "300">   <img src = "https://github.com/Kee2u/Stock-Analysis/blob/main/resources/Green_Stocks_2018.PNG?raw=true" width = "283">

### Refactored Code  

I then refactored the code to make it more efficient. Previously, the  code went through all the rows 12 times for each ticker. This time the code ran through the rows only once and updated the ticker as it went along. 
I did this by updating the ticker once the code detected the ending price of the current ticker. This was possible because the rows were sorted chronologically: The ending price of one ticker was immediately before the starting price of the next as the code looped through the rows.

In addition to creating an array for the ticker index like before, I created arrays for ticker volumes, starting price and ending price.

**Here is the code I used. The values of the ticker array was initalized with the values of all 12 tickers before these lines:**
  
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
    

