# Stock-analysis
 Refactoring the stocks data analysis code to compare yearly return percentage and formatting the results with VBA.

# Overview of Project
This project is an assignmnet from Data Analytics program - VBA Module<br>
- Deliverable 1:  Refactor VBA code and measure performance<br>
  This deliverable will include an updated workbook and a folder with PNGs of the pop-ups with script run time <br>

# Results

Original code has inner loop which increased runs count compare to the refactored code. 
<h3>Original Code</h3>
<code>
 
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
 </code>
 
<h3>Refactorted Code</h3>
<code>
 
  
        For i = 0 To 11
            TickerVolumes(i) = 0
        Next i
        
        For i = 2 To RowCount
            TickerVolumes(tickerindex) = TickerVolumes(tickerindex) + Cells(i, 8).Value
        
            If Cells(i - 1, 1).Value <> tickers(tickerindex) Then
                TickerStartingPrice(tickerindex) = Cells(i, 6).Value
            End If
            
            If Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                TickerEndingPrice(tickerindex) = Cells(i, 6).Value
                tickerindex = tickerindex + 1
            End If
            
        Next i
 </code>

Run time comparison <br>
Both 2017 and 2018 refactored code process time is faster than the original code. <br>
Array is applied instead of pulling cells from data for loop, which has shorten the time for the code system.<br>
<br>
 
Original 2017 Code<br>
![date](Original%202017.png)
Refactored 2017 Code<br>
![date](Refactored%202017.png)
<br>
<br>
Original 2018 Code<br>
![date](Original%202018.png)
Refactored 2018 Code<br>
![date](Refactored%202018.png)




# Summary
- Advantages or disadvantages of refactoring code<br>
The advantage of refactoring is to boost up system performance, as it could made the script run efficiently,<br> 
using less memory and more simplify code for future users.<br>
The downside of refactoring is it takes time to restructure the system which might increase the risk of error when messing with exsiting format.<br>

- Pros and cons apply to refactoring the original VBA script<br>
The refactorted code lowered the process time with less storage than the original code. But the refactorting time and result doesn't make a significant changes. It depends on the usage of the code.
 

