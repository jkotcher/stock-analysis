# Analysis of Alternative Energy Stock Data


## Overview

### Purpose
The purpose of this project was to refactor code that loops through stock data and allows to quickly run analyses on stock data.  Steve was looking to do these analyses for his parents so they would be able to make a good decision about what stocks to invest in.  So he came to us to help him out and use VBA to automate repetitive tasks.  Refactoring code is a common part of the process of programming and putting together a block of code is not always perfect the first time.  Refactoring it helps to improve the code without adding any additional functionality.


## Results

### Code Script

In the process of refactoring our code some of the differences to improve the code including creating a ticker index and three output arrays that the ticker index would be used for.

 tickerIndex = tickers(i)
        
        '1b) Create three output arrays
        Dim tickerVolumes As Long
    
        Dim tickerStartingPrices As Single, tickerEndingPrices As Single
        
Using the ticker index we used If-then statements to move the index with price columns and used it to increase the total daily volume.

 If Cells(j, 1).Value = tickerIndex Then
        
        tickerVolumes = tickerVolumes + Cells(j, 8).Value
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(j, 1).Value = tickerIndex And Cells(j - 1, 1).Value <> tickerIndex Then
            
            tickerStartingPrices = Cells(j, 6).Value
            
            
         End If


In addition the formatting and code to loop through all the stock data was also included in the macro.
 Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
### Stock Performance

A comparison of stock performance between 2017 and 2018.

<img width="264" alt="Stock_performance" src="https://user-images.githubusercontent.com/29406929/173357560-95962238-e00e-4190-a87f-b52487a5dabe.png">

2017



<img width="263" alt="Stock_performance_2018" src="https://user-images.githubusercontent.com/29406929/173357627-f98ed6a0-d088-4bf9-b433-369f93a1569e.png">

2018


### Code Performance

Refactored script run times

<img width="260" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/29406929/173362432-ec393dfe-e9ba-422d-8000-49d2b2347c01.png">



<img width="261" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/29406929/173362480-c1c96a11-f40b-4562-b887-4b19e328f6fc.png">

After running several tests on both the original and refactored code the refactored code runs faster than the original.  The original code was usually running in about 0.60 or 0.61 seconds and sometimes a little longer.  The above images are of the refactored code where they are running slower than they normally would.  Usually the refactored code runs in about 0.57 or 0.59 seconds.  I do think that refactoring the code made it faster and that what we specifically improved was the efficiency with which the code does the work.

## Summary

### Advantages and Disadvantages of Refactoring
There are several advantages and disadvantages to refactoring code.  Some of the advantages are you improve the logical flow, efficiency, how quickly code runs a particular analysis.  By doing this though you can take code that works and is doing something and make it better.  Making it better you improve how the code works and ultimately could run through larger analyses much faster.  One disadvantge is added time spent editing code to try and make it better.  Not all code has to be refactored because to do so might make it work not as well.  Other disadvantages are increased chance of making a mistake and increased complexity making it harder for others to read your code should they want to look at it to see some of the component parts.

### Advantanges and Disadvantages of Refactoring applied to VBA script
The advantages and disadvantages outlined earlier apply to the VBA script just as much as they do to any other script that may be refactored.  First we had a code block that completed a task and was working.  However, it was running the analysis slower than we would have liked and so refactoring was needed to make the code run faster.  One main disadvantage associated with this project was that there were several mistakes made mainly typos and the like that caused the code to break at first.  Once those issues were fixed the code ran better than it did initially.  In conclusion, I would say that refactoring code as a fundamental step in programming should be used more often than not.
