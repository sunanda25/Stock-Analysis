## Overview
Steve’s parents are very passionate about Green energy and they are looking forward to investing money in it. Without conducting much research, they decide to invest all their money in “DAQO New Energy Corp”. Steve is concerned about their parents not diversifying their stocks, so he decided to build an Excel sheet for 2017 and 2018 stocks of 12 different Green energy companies. To help Steve analyze this data, an efficient tool more complex than Excel formulas is required. VBA which is an extension to Excel has been used for this purpose. An Excel macro has already been provided to do the stock analysis; however, the code is neither optimized nor easily readable. The original macro has been refactored by adding new variables and arrays which resulted in more readability and faster performance of the macro. 

## Results
### Refactoring the code
To improve efficiency, the original code has been refactored by switching the order of the loops. 

1.	Created a tickerIndex variable as a Single data type and assigned a value of Zero.

![tickerIndex](https://user-images.githubusercontent.com/76491891/110210096-23f1b200-7e5e-11eb-8550-8537e2602cca.png)


2.	Three output arrays are created:
   *	The tickerVolume array as a Long data type
   *	The tickerStartingPrices array as a Single data type
   *	The tickerEndingPrices array as a Single data type
![Arrays](https://user-images.githubusercontent.com/76491891/110210112-32d86480-7e5e-11eb-90ae-9ac284cb3fca.png)

3.	Created a For loop to initialize the tickerVolumes to Zero.
![TickerVolumes](https://user-images.githubusercontent.com/76491891/110210125-4388da80-7e5e-11eb-9b1f-bf2f33c03612.png)

4.	Created a For loop over all the rows in a spreadsheet by increasing the tickerVolumes using the tickerIndex. Using If-Then statements:
   *	If the current row is the first row of the selected tickerIndex, then assign the current closing price to the tickerStartingPrices.
   *	If the current row is the last row of the selected tickerIndex, then assign the current closing price to the tickerEndingPrices.
   *	Increase the tickerIndex if the next row ticker does not match the previous row ticker.
 ![For loop](https://user-images.githubusercontent.com/76491891/110210136-4c79ac00-7e5e-11eb-92c0-88114216ae91.png)

5.	Created a For loop using tickerIndex to display the outputs as Ticker, Total daily volume, Return.
 ![Output](https://user-images.githubusercontent.com/76491891/110210154-5a2f3180-7e5e-11eb-9a80-e03e0b46d586.png)

### Stock Performance of 2017 and 2018
*	Stock returns for 2017 are positive except for “TERP”.
*	Stock returns for 2018 are mostly negative except for “ENPH”, “RUN”.
*	DQ ticker has a 199.4% return in 2017 whereas in 2018 it has -62.6%.
<img width="335" alt="All_Stocks_2017_2018" src="https://user-images.githubusercontent.com/76491891/110210216-982c5580-7e5e-11eb-9968-be81e09a4df5.png">

### Execution Times
The execution time of original code for the years 2017 and 2018:
![Original_2017](https://user-images.githubusercontent.com/76491891/110210221-a5e1db00-7e5e-11eb-85c5-cfc32246730a.png)

![Original_2018](https://user-images.githubusercontent.com/76491891/110210229-abd7bc00-7e5e-11eb-9efa-727cc6cba695.png)

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
