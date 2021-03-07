## Overview
Steve’s parents are very passionate about Green energy and they are looking forward to investing money in it. Without conducting much research, they decide to invest all their money in “DAQO New Energy Corp”. Steve is concerned about their parents not diversifying their stocks, so he decided to build an Excel sheet for 2017 and 2018 stocks of 12 different Green energy companies. To help Steve analyze this data, an efficient tool more complex than Excel formulas is required. VBA which is an extension to Excel has been used for this purpose. An Excel macro has already been provided to do the stock analysis; however, the code is neither optimized nor easily readable. The original macro has been refactored by adding new variables and arrays which resulted in more readability and faster performance of the macro. 

## Results
### Refactoring the code
To improve efficiency, the original code has been refactored by switching the order of the loops. 

1.	Created a tickerIndex variable as a Single data type and assigned a value of Zero.

~~~
Dim tickerIndex As Single
tickerIndex = 0
~~~

2.	Three output arrays are created:
   *	The tickerVolume array as a Long data type
   *	The tickerStartingPrices array as a Single data type
   *	The tickerEndingPrices array as a Single data type

~~~
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
~~~
    
3.	Created a For loop to initialize the tickerVolumes to Zero.

~~~
For i = 0 To 11
   tickerVolumes(i) = 0
Next i 
~~~

4.	Created a For loop over all the rows in a spreadsheet by increasing the tickerVolumes using the tickerIndex. Using If-Then statements:
   *	If the current row is the first row of the selected tickerIndex, then assign the current closing price to the tickerStartingPrices.
   *	If the current row is the last row of the selected tickerIndex, then assign the current closing price to the tickerEndingPrices.
   *	Increase the tickerIndex if the next row ticker does not match the previous row ticker.

~~~
For i = 2 To RowCount  
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then         
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value           
         End If         
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then       
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value                   
            tickerIndex = tickerIndex + 1         
          End If 
Next i
~~~

5.	Created a For loop using tickerIndex to display the outputs as Ticker, Total daily volume, Return.
 
 ~~~
For i = 0 To 11  
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1     
Next i
 ~~~
 
### Stock Performance of 2017 and 2018
*	Stock returns for 2017 are positive except for “TERP”.
*	Stock returns for 2018 are mostly negative except for “ENPH”, “RUN”.
*	DQ ticker has a 199.4% return in 2017 whereas in 2018 it has -62.6%.

<img width="400" alt="All_Stocks_2017_2018" src="https://user-images.githubusercontent.com/76491891/110210216-982c5580-7e5e-11eb-9968-be81e09a4df5.png">

### Execution Times
The execution time of original code for the years 2017 and 2018:

![Original_2017](https://user-images.githubusercontent.com/76491891/110228503-a797b700-7ecf-11eb-90f6-1fc0a27022b3.png)
![Original_2018](https://user-images.githubusercontent.com/76491891/110228505-abc3d480-7ecf-11eb-8003-9fa6ad6bbef7.png)

The execution time of refactored code for the years 2017 and 2018:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/76491891/110210239-b8f4ab00-7e5e-11eb-9de4-2d528ecdd120.png) 
![VBA_Challenge_2018](https://user-images.githubusercontent.com/76491891/110210246-beea8c00-7e5e-11eb-9715-1d2a892a8f1e.png)

## Summary
### Advantages or Disadvantages of Refactoring code:
*	The refactored code is well structured and easy to read. 
*	The refactored code works more efficiently than the original code when working with large sets of data.
*	The run time for the refactored code is faster than the original code run time.
*	A disadvantage of refactoring code is new bugs can be introduced.

### Pros and cons for refactoring the original VBA script:
*	Refactoring the original code makes the macro run faster. The run times of the original code in 2017 and 2018 are 0.731562500004657, 0.758749999993597. After refactoring the code, run times of 2017 and 2018 are 0.1171875, 0.1328125 which are much faster (~600%) compared to original code run times.
*	The refactored code is more readable with appropriate variable names and resequencing of ‘For’ loops.
