

# Stock Analysis Utilizing VBA
## Overview of Project
My client needs a user-friendly Excel sheet that calculates the total daily volume and returns for 2017 and 2018 stocks.
I have modified the sheet to do so, however, if more data is added, analysis can become tedious at the current run time. 
###Purpose
In this project, I’ve been tasked with optimizing the current run time and maintaining the integrity of the analysis.
I will refractor code to ensure a more efficient stock analysis performance.

### Preprocessing 
The data contains two charts for the years 2017 and 2018 of 12 select stocks.
Each sheet contains the stock's ticker, date, opening price, high, low, closing price, adjusted closing price, and daily volume.

## Results
In my initial attempt, I created a nested loop that went through each individual cell to determine the ticker, sum the total volume of the same tickers, and calculate the 
difference in staring price and closing price for each ticker. This means the code ran 3012 times to complete the stock analysis.
The worksheets were activated upon the input box, so both 2017 and 2018 had the same run time with this code.

### Nested Loop Runtime 2017



### Nested Loop Runtime 2018



## Refractored code
When refactoring the code, I didn’t create any nested loops. I created a code that went through the tickers, starting price, ending price, and total volume at once. Indexing the arrays together enables a quicker analysis. This modification saved 0.828125 seconds in processing time.



   
   

    
    '1a) Create a ticker Index
    Dim tickerindex As Integer
    tickerindex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single
        
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerindex) Then
            tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        End If
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            tickerStartingPrice(tickerindex) = Cells(i, 6).Value
            
        End If
        
            
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
                tickerEndingPrice(tickerindex) = Cells(i, 6).Value
                        

            '3d Increase the tickerIndex.
                tickerindex = tickerindex + 1
            
            End If
            
        'End If
    
    Next i
    
    '4) Loop through arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1
        
        
   
### Improved Analysis Runtime 2017 


### Improved Analysis Runtime 2018 

## Summary

Refactoring code is beneficial because it gives a foundation to start with. 
It allows a way to think through a problem without spending hours on code that may not work. 
Overall, refactoring code provides organized guidance and can save time. In contrast, refactoring code can be detrimental, as it risks adding bugs that did not exist before, causing a tedious debugging process.
Refactoring was remarkable for simplifying the process of the stock analysis. The issues I experienced refactoring the VBA code were due to my ignorance of the “tickerIndex” variable and its utilization.  That variable was essential to going through all the arrays simultaneously.  The original code was fine because I understood the thought process for solving the problem step by step. The original VBA code wasn’t efficient and was time-consuming to run as such.

## Resources
[Microsoft](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/array-function)
